using System.IO;

namespace XwbStudio.Core;

/// <summary>WAV header construction for the codecs found inside XWB banks.</summary>
public static class WavCodecs
{
    public const int AdpcmBlockAlignOffset = 22;

    /// <summary>
    /// Builds a RIFF/WAVE header for the given codec, or null when the codec
    /// has no WAV representation (WMA).
    /// </summary>
    public static byte[]? MakeWavHeader(XwbCodec codec, int channels, int rate, int bits, int align, uint dataSize)
    {
        if (channels <= 0)
            channels = 1;

        ushort fmtTag;
        ushort bitsPer;
        int blockAlign;
        int avgBytes;
        byte[] extra;

        switch (codec)
        {
            case XwbCodec.Pcm:
                fmtTag = 0x0001;
                bitsPer = (ushort)(8 << bits);
                blockAlign = (bitsPer / 8) * channels;
                avgBytes = rate * blockAlign;
                extra = [];
                break;

            case XwbCodec.Xma:
                fmtTag = 0x0069;
                bitsPer = 4;
                blockAlign = 36 * channels;
                avgBytes = (689 * blockAlign) + 4;
                extra = [0x02, 0x00, 0x40, 0x00];
                break;

            case XwbCodec.Adpcm:
                fmtTag = 0x0002;
                bitsPer = 4;
                blockAlign = (align + AdpcmBlockAlignOffset) * channels;
                avgBytes = 21 * blockAlign;
                int samplesPerBlock = ((blockAlign / channels - 7) * 2) + 2;
                using (var ems = new MemoryStream())
                using (var ew = new BinaryWriter(ems))
                {
                    ew.Write((ushort)2);
                    ew.Write((ushort)samplesPerBlock);
                    ew.Flush();
                    extra = ems.ToArray();
                }
                break;

            default:
                return null;
        }

        int fmtSize = 16 + extra.Length;
        uint riffSize = (uint)(4 + 8 + fmtSize + 8 + dataSize);

        using var ms = new MemoryStream();
        using var w = new BinaryWriter(ms);
        w.Write("RIFF"u8);
        w.Write(riffSize);
        w.Write("WAVE"u8);
        w.Write("fmt "u8);
        w.Write((uint)fmtSize);
        w.Write((short)fmtTag);
        w.Write((ushort)channels);
        w.Write((uint)rate);
        w.Write((uint)avgBytes);
        w.Write((ushort)blockAlign);
        w.Write(bitsPer);
        w.Write(extra);
        w.Write("data"u8);
        w.Write(dataSize);
        w.Flush();
        return ms.ToArray();
    }
}
