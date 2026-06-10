using System.Buffers.Binary;
using System.IO;

namespace XwbStudio.Core;

public sealed record WavInfo(
    string Path,
    int Channels,
    int SampleRate,
    int BitsPerSample,
    int BlockAlign,
    byte[] Data);

/// <summary>Plain RIFF/WAVE reading helpers.</summary>
public static class WavFile
{
    /// <summary>Parses a WAV file's format chunk and returns its raw PCM data.</summary>
    public static WavInfo Parse(string path)
    {
        string name = System.IO.Path.GetFileName(path);
        using var f = File.OpenRead(path);

        if (!ReadFourCc(f).AsSpan().SequenceEqual("RIFF"u8))
            throw new InvalidDataException($"{name} is not a valid WAV file");
        Skip(f, 4); // RIFF size
        if (!ReadFourCc(f).AsSpan().SequenceEqual("WAVE"u8))
            throw new InvalidDataException($"{name} is not a valid WAV file");

        int channels = 0, sampleRate = 0, bitsPerSample = 0, blockAlign = 0;
        byte[] audioData = [];

        while (true)
        {
            var chunkId = new byte[4];
            int got = f.Read(chunkId, 0, 4);
            if (got < 4)
                break;
            uint chunkSize = ReadU32(f);

            if (chunkId.AsSpan().SequenceEqual("fmt "u8))
            {
                ReadU16(f); // format tag
                channels = ReadU16(f);
                sampleRate = (int)ReadU32(f);
                Skip(f, 4); // avg bytes per sec
                blockAlign = ReadU16(f);
                bitsPerSample = ReadU16(f);
                int remaining = (int)chunkSize - 16;
                if (remaining > 0)
                    Skip(f, remaining);
            }
            else if (chunkId.AsSpan().SequenceEqual("data"u8))
            {
                audioData = new byte[chunkSize];
                f.ReadExactly(audioData);
            }
            else
            {
                Skip(f, (int)chunkSize);
            }
        }

        if (audioData.Length == 0)
            throw new InvalidDataException($"No audio data in {name}");
        if (bitsPerSample is not (8 or 16))
            throw new InvalidDataException($"{name}: only 8-bit and 16-bit PCM supported");

        return new WavInfo(path, channels, sampleRate, bitsPerSample, blockAlign, audioData);
    }

    /// <summary>
    /// Returns only the raw audio bytes of a WAV file (no header).
    /// Non-RIFF files are returned verbatim.
    /// </summary>
    public static byte[] StripHeader(string path)
    {
        using var f = File.OpenRead(path);

        var sig = new byte[4];
        int got = f.Read(sig, 0, 4);
        if (got < 4 || !sig.AsSpan().SequenceEqual("RIFF"u8))
        {
            f.Position = 0;
            using var ms = new MemoryStream();
            f.CopyTo(ms);
            return ms.ToArray();
        }

        Skip(f, 4); // RIFF size
        Skip(f, 4); // WAVE

        while (true)
        {
            var chunkId = new byte[4];
            if (f.Read(chunkId, 0, 4) < 4)
                break;
            uint chunkSize = ReadU32(f);
            if (chunkId.AsSpan().SequenceEqual("data"u8))
            {
                var data = new byte[chunkSize];
                f.ReadExactly(data);
                return data;
            }
            Skip(f, (int)chunkSize);
        }

        return [];
    }

    private static byte[] ReadFourCc(Stream s)
    {
        var b = new byte[4];
        s.ReadExactly(b);
        return b;
    }

    private static uint ReadU32(Stream s)
    {
        Span<byte> b = stackalloc byte[4];
        s.ReadExactly(b);
        return BinaryPrimitives.ReadUInt32LittleEndian(b);
    }

    private static ushort ReadU16(Stream s)
    {
        Span<byte> b = stackalloc byte[2];
        s.ReadExactly(b);
        return BinaryPrimitives.ReadUInt16LittleEndian(b);
    }

    private static void Skip(Stream s, int count) => s.Seek(count, SeekOrigin.Current);
}
