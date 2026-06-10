namespace XwbStudio.Core;

public enum XwbCodec
{
    Pcm = 0,
    Xma = 1,
    Adpcm = 2,
    Wma = 3,
}

/// <summary>One audio entry inside a wave bank.</summary>
public sealed class XwbTrack
{
    public required int Index { get; init; }

    /// <summary>Absolute byte offset of the audio data in the file.</summary>
    public required long Offset { get; init; }

    public required uint Size { get; init; }
    public required XwbCodec Codec { get; init; }
    public required int Channels { get; init; }
    public required int SampleRate { get; init; }

    /// <summary>0 = 8-bit, 1 = 16-bit (PCM only).</summary>
    public required int Bits { get; init; }

    public required int Align { get; init; }
    public required double DurationSeconds { get; init; }

    /// <summary>Absolute byte offset of this entry's metadata record (used by the injector).</summary>
    public required long EntryPosition { get; init; }

    public string HexName => Index.ToString("x8");

    public string CodecName => Codec switch
    {
        XwbCodec.Pcm => "PCM",
        XwbCodec.Xma => "XMA",
        XwbCodec.Adpcm => "ADPCM",
        XwbCodec.Wma => "WMA",
        _ => "???",
    };
}
