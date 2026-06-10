using System.IO;
using System.Text;

namespace XwbStudio.Core;

/// <summary>Bundles WAV files into a brand-new XWB wave bank (PCM only, version 43).</summary>
public static class XwbBuilder
{
    public static void Create(IReadOnlyList<string> wavPaths, string outPath, string bankName = "CustomBank")
    {
        if (wavPaths.Count == 0)
            throw new ArgumentException("No WAV files provided");

        var tracks = wavPaths.Select(WavFile.Parse).ToList();

        int numTracks = tracks.Count;
        const int metaElementSize = 24; // standard WAVEBANKENTRY size

        int hdrSize = 4 + 4 + 4 + 5 * 8;                 // sig + version + hdrver + 5 segments = 52
        int bankDataSize = 4 + 4 + 64 + 4 + 4 + 4 + 4 + 4; // = 92
        int metaSize = numTracks * metaElementSize;

        int bankDataOff = hdrSize;
        int metaOff = bankDataOff + bankDataSize;
        int waveOffRaw = metaOff + metaSize;
        int waveOff = (waveOffRaw + 3) & ~3;             // align to 4 bytes

        var audioOffsets = new int[numTracks];
        int cur = 0;
        for (int i = 0; i < numTracks; i++)
        {
            audioOffsets[i] = cur;
            cur += tracks[i].Data.Length;
        }
        int totalAudio = cur;

        using var f = File.Create(outPath);
        using var w = new BinaryWriter(f);

        // WAVEBANKHEADER
        w.Write("WBND"u8);                 // LE signature
        w.Write(43u);                      // version
        w.Write(1u);                       // dwHeaderVersion
        w.Write((uint)bankDataOff); w.Write((uint)bankDataSize); // seg 0 BANKDATA
        w.Write((uint)metaOff); w.Write((uint)metaSize);          // seg 1 ENTRYMETADATA
        w.Write(0u); w.Write(0u);                                  // seg 2 SEEKTABLES
        w.Write(0u); w.Write(0u);                                  // seg 3 ENTRYNAMES
        w.Write((uint)waveOff); w.Write((uint)totalAudio);         // seg 4 ENTRYWAVEDATA

        // WAVEBANKDATA
        w.Write(0u);                       // dwFlags
        w.Write((uint)numTracks);          // dwEntryCount
        var nameBytes = new byte[64];
        var ascii = Encoding.ASCII.GetBytes(bankName);
        Array.Copy(ascii, nameBytes, Math.Min(ascii.Length, 64));
        w.Write(nameBytes);                // szBankName[64]
        w.Write((uint)metaElementSize);    // dwEntryMetaDataElementSize
        w.Write(0u);                       // dwEntryNameElementSize
        w.Write(4u);                       // dwAlignment
        w.Write(0u);                       // CompactFormat
        w.Write(0u);                       // BuildTime

        // WAVEBANKENTRY per track
        for (int i = 0; i < numTracks; i++)
        {
            var t = tracks[i];
            uint codec = 0; // PCM
            uint chans = (uint)t.Channels;
            uint rate = (uint)t.SampleRate;
            uint bits = t.BitsPerSample == 16 ? 1u : 0u;
            uint align = 0;
            uint fmt = (codec & 0x3) | ((chans & 0x7) << 2) | ((rate & 0x3FFFF) << 5)
                     | ((align & 0xFF) << 23) | ((bits & 0x1) << 31);

            w.Write(0u);                       // dwFlagsAndDuration
            w.Write(fmt);                      // Format
            w.Write((uint)audioOffsets[i]);    // PlayRegion.dwOffset
            w.Write((uint)t.Data.Length);      // PlayRegion.dwLength
            w.Write(0u);                       // LoopRegion.dwOffset
            w.Write(0u);                       // LoopRegion.dwLength
        }

        // Pad to wave data start
        while (f.Position < waveOff)
            w.Write((byte)0);

        foreach (var t in tracks)
            w.Write(t.Data);
    }
}
