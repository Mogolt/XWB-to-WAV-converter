using System.Buffers.Binary;
using System.IO;

namespace XwbStudio.Core;

/// <summary>
/// Rebuilds an XWB file with one track's audio replaced.
/// All play-region offsets are recalculated from scratch.
/// </summary>
public static class XwbInjector
{
    public static void Rebuild(string srcPath, int replaceIndex, string wavPath, string outPath)
    {
        byte[] src = File.ReadAllBytes(srcPath);

        byte[] newAudio = WavFile.StripHeader(wavPath);
        if (newAudio.Length == 0)
            throw new InvalidDataException("Could not read audio data from WAV file");

        var bank = XwbBank.Open(srcPath);
        var tracks = bank.Tracks;
        if (tracks.Count == 0)
            throw new InvalidDataException("No tracks found in XWB");

        bool big = src.AsSpan(0, 4).SequenceEqual("DNBW"u8);

        uint ReadU32(int off) => big
            ? BinaryPrimitives.ReadUInt32BigEndian(src.AsSpan(off, 4))
            : BinaryPrimitives.ReadUInt32LittleEndian(src.AsSpan(off, 4));

        uint version = ReadU32(4);
        int lastSegment = version <= 3 ? 3 : 4;
        int hdrOffset = 8; // after sig + version
        if (version >= 42)
            hdrOffset += 4;

        var segOffsets = new uint[lastSegment + 1];
        for (int i = 0; i <= lastSegment; i++)
            segOffsets[i] = ReadU32(hdrOffset + i * 8);

        uint waveDataStart = segOffsets[lastSegment];

        // Collect all track audio blobs in order, swapping in the replacement.
        var audioBlobs = new byte[tracks.Count][];
        for (int i = 0; i < tracks.Count; i++)
        {
            var t = tracks[i];
            audioBlobs[i] = t.Index == replaceIndex
                ? newAudio
                : src.AsSpan((int)t.Offset, (int)t.Size).ToArray();
        }

        // Everything before the wave data is copied, then patched in place.
        var head = src.AsSpan(0, (int)waveDataStart).ToArray();

        void WriteU32(int off, uint value)
        {
            if (big)
                BinaryPrimitives.WriteUInt32BigEndian(head.AsSpan(off, 4), value);
            else
                BinaryPrimitives.WriteUInt32LittleEndian(head.AsSpan(off, 4), value);
        }

        // New offsets relative to the wave data segment.
        var newOffsets = new uint[tracks.Count];
        uint current = 0;
        for (int i = 0; i < tracks.Count; i++)
        {
            newOffsets[i] = current;
            current += (uint)audioBlobs[i].Length;
        }
        uint newWaveDataSize = current;

        // Patch each track entry's offset and size in the metadata table.
        for (int i = 0; i < tracks.Count; i++)
        {
            long ep = tracks[i].EntryPosition;
            // v1 layout: fmt(4)+playoff(4)+playlen(4); later: flags(4)+fmt(4)+playoff(4)+playlen(4)
            int playOffField = (int)ep + (version == 1 ? 4 : 8);
            int playLenField = (int)ep + (version == 1 ? 8 : 12);

            WriteU32(playOffField, newOffsets[i]);
            WriteU32(playLenField, (uint)audioBlobs[i].Length);
        }

        // Patch segment table: wave data length.
        WriteU32(hdrOffset + lastSegment * 8 + 4, newWaveDataSize);

        using var f = File.Create(outPath);
        f.Write(head);
        foreach (var blob in audioBlobs)
            f.Write(blob);
    }
}
