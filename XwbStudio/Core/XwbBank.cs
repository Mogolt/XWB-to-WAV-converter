using System.Buffers.Binary;
using System.IO;

namespace XwbStudio.Core;

/// <summary>
/// Parses an XACT wave bank (.xwb) and exposes its track table.
/// Handles both little- and big-endian banks, versions 1 through 46,
/// compact banks, and banks embedded inside container files.
/// </summary>
public sealed class XwbBank
{
    public const int CopyChunk = 65536;

    public string FilePath { get; }
    public uint Version { get; }
    public bool IsBigEndian { get; }
    public bool IsCompact { get; }
    public IReadOnlyList<XwbTrack> Tracks { get; }

    private XwbBank(string filePath, uint version, bool bigEndian, bool compact, List<XwbTrack> tracks)
    {
        FilePath = filePath;
        Version = version;
        IsBigEndian = bigEndian;
        IsCompact = compact;
        Tracks = tracks;
    }

    public static XwbBank Open(string path)
    {
        using var f = File.OpenRead(path);

        var sig = new byte[4];
        f.ReadExactly(sig);

        bool big;
        if (sig.AsSpan().SequenceEqual("WBND"u8))
        {
            big = false;
        }
        else if (sig.AsSpan().SequenceEqual("DNBW"u8))
        {
            big = true;
        }
        else
        {
            // Signature not at the start — scan the first 1 MB for an embedded bank.
            f.Position = 0;
            var scan = new byte[Math.Min(1024 * 1024, f.Length)];
            f.ReadExactly(scan);
            int found = -1;
            for (int i = 0; i < scan.Length - 8; i++)
            {
                var s = scan.AsSpan(i, 4);
                if ((s.SequenceEqual("WBND"u8) && scan[i + 7] == 0) ||
                    (s.SequenceEqual("DNBW"u8) && scan[i + 4] == 0))
                {
                    found = i;
                    break;
                }
            }
            if (found < 0)
                throw new InvalidDataException("Not a valid XWB file");
            big = scan[found] == (byte)'D';
            f.Position = found + 4;
        }

        uint version = ReadU32(f, big);
        int lastSegment = version <= 3 ? 3 : 4;
        if (version >= 42)
            Skip(f, 4); // dwHeaderVersion

        var segments = new (uint Offset, uint Length)[5];
        for (int i = 0; i <= lastSegment; i++)
            segments[i] = (ReadU32(f, big), ReadU32(f, big));

        long bankDataOffset = version == 1 ? f.Position : segments[0].Offset;
        f.Position = bankDataOffset;

        uint flags = ReadU32(f, big);
        uint entryCount = ReadU32(f, big);
        bool compact = (flags & 0x00020000) != 0;
        Skip(f, version is 2 or 3 ? 16 : 64); // szBankName

        uint metaElementSize;
        uint alignment;
        uint compactFormat;
        long wavebankOffset;

        if (version == 1)
        {
            wavebankOffset = f.Position;
            metaElementSize = 20;
            alignment = 4;
            compactFormat = 0;
        }
        else
        {
            metaElementSize = ReadU32(f, big);
            ReadU32(f, big); // dwEntryNameElementSize
            alignment = ReadU32(f, big);
            wavebankOffset = segments[1].Offset;
            compactFormat = compact ? ReadU32(f, big) : 0;
        }

        long playRegionOffset = segments[lastSegment].Offset;
        if (playRegionOffset == 0)
            playRegionOffset = wavebankOffset + entryCount * metaElementSize;

        var tracks = new List<XwbTrack>((int)entryCount);

        for (int entryIdx = 0; entryIdx < entryCount; entryIdx++)
        {
            long ep = wavebankOffset + (long)entryIdx * metaElementSize;
            f.Position = ep;

            uint fmt;
            long playOff;
            long playLen;

            if (compact)
            {
                uint rawVal = ReadU32(f, big);
                fmt = compactFormat;
                playOff = (rawVal & 0x1FFFFF) * (long)alignment;
                if (entryIdx == entryCount - 1)
                {
                    playLen = segments[lastSegment].Length - playOff;
                }
                else
                {
                    f.Position = ep + metaElementSize;
                    playLen = (ReadU32(f, big) & 0x1FFFFF) * (long)alignment - playOff;
                }
            }
            else
            {
                fmt = 0;
                playOff = 0;
                playLen = 0;
                if (version == 1)
                {
                    fmt = ReadU32(f, big);
                    playOff = ReadU32(f, big);
                    playLen = ReadU32(f, big);
                }
                else
                {
                    if (metaElementSize >= 4) Skip(f, 4); // dwFlagsAndDuration
                    if (metaElementSize >= 8) fmt = ReadU32(f, big);
                    if (metaElementSize >= 12) playOff = ReadU32(f, big);
                    if (metaElementSize >= 16) playLen = ReadU32(f, big);
                }
                if (metaElementSize < 24 && playLen == 0)
                    playLen = segments[lastSegment].Length;
            }

            playOff += playRegionOffset;

            int codec, chans;
            if (version == 1)
            {
                codec = (int)(fmt & 0x01);
                chans = (int)((fmt >> 1) & 0x07);
            }
            else
            {
                codec = (int)(fmt & 0x03);
                chans = (int)((fmt >> 2) & 0x07);
            }
            int rate = (int)((fmt >> 5) & 0x3FFFF);
            int align = (int)((fmt >> 23) & 0xFF);
            int bits = (int)((fmt >> 31) & 0x01);

            if (playLen == 0)
                continue;

            double duration = ComputeDuration((XwbCodec)codec, playLen, rate, chans, bits, align);

            tracks.Add(new XwbTrack
            {
                Index = entryIdx,
                Offset = playOff,
                Size = (uint)playLen,
                Codec = (XwbCodec)codec,
                Channels = chans,
                SampleRate = rate,
                Bits = bits,
                Align = align,
                DurationSeconds = duration,
                EntryPosition = ep,
            });
        }

        return new XwbBank(path, version, big, compact, tracks);
    }

    private static double ComputeDuration(XwbCodec codec, long size, int rate, int channels, int bits, int align)
    {
        try
        {
            int ch = Math.Max(channels, 1);
            switch (codec)
            {
                case XwbCodec.Pcm when rate > 0:
                    int bps = 8 << bits;
                    return size / (rate * (bps / 8.0) * ch);

                case XwbCodec.Adpcm when rate > 0:
                    int blockAlign = (align + WavCodecs.AdpcmBlockAlignOffset) * ch;
                    if (blockAlign <= 0)
                        return 0;
                    int samplesPerBlock = ((blockAlign / ch - 7) * 2) + 2;
                    return (double)size / blockAlign * samplesPerBlock / rate;

                default:
                    return 0;
            }
        }
        catch
        {
            return 0;
        }
    }

    /// <summary>Extracts a single track to <paramref name="outPath"/> (WAV for PCM/XMA/ADPCM, raw for WMA).</summary>
    public void ExtractTrack(XwbTrack track, string outPath)
    {
        using var f = File.OpenRead(FilePath);
        f.Position = track.Offset;
        using var fout = File.Create(outPath);
        if (track.Codec != XwbCodec.Wma)
        {
            var header = WavCodecs.MakeWavHeader(track.Codec, track.Channels, track.SampleRate, track.Bits, track.Align, track.Size);
            if (header is not null)
                fout.Write(header);
        }
        CopyBytes(f, fout, track.Size);
    }

    /// <summary>
    /// Extracts every track into <paramref name="outDir"/>. Returns the written file paths.
    /// Stops between tracks when <paramref name="ct"/> is cancelled.
    /// </summary>
    public List<string> ExtractAll(string outDir, IReadOnlyDictionary<string, string>? trackNames = null, CancellationToken ct = default)
    {
        Directory.CreateDirectory(outDir);
        var outputs = new List<string>();

        using var f = File.OpenRead(FilePath);
        foreach (var track in Tracks)
        {
            if (ct.IsCancellationRequested)
                break;

            string baseName = track.HexName;
            if (trackNames is not null && trackNames.TryGetValue(baseName, out var friendly))
                baseName = friendly;

            string ext = track.Codec == XwbCodec.Wma ? ".wma" : ".wav";
            string outPath = Path.Combine(outDir, baseName + ext);

            f.Position = track.Offset;
            using (var fout = File.Create(outPath))
            {
                if (track.Codec != XwbCodec.Wma)
                {
                    var header = WavCodecs.MakeWavHeader(track.Codec, track.Channels, track.SampleRate, track.Bits, track.Align, track.Size);
                    if (header is not null)
                        fout.Write(header);
                }
                CopyBytes(f, fout, track.Size);
            }

            outputs.Add(outPath);
        }

        return outputs;
    }

    internal static void CopyBytes(Stream input, Stream output, long size)
    {
        var buf = new byte[CopyChunk];
        long remaining = size;
        while (remaining > 0)
        {
            int read = input.Read(buf, 0, (int)Math.Min(CopyChunk, remaining));
            if (read <= 0)
                break;
            output.Write(buf, 0, read);
            remaining -= read;
        }
    }

    internal static uint ReadU32(Stream s, bool bigEndian)
    {
        Span<byte> b = stackalloc byte[4];
        s.ReadExactly(b);
        return bigEndian
            ? BinaryPrimitives.ReadUInt32BigEndian(b)
            : BinaryPrimitives.ReadUInt32LittleEndian(b);
    }

    private static void Skip(Stream s, int count) => s.Seek(count, SeekOrigin.Current);
}
