using System.IO;
using XwbStudio.Core;

// Round-trip test of the XWB core: build → parse → extract → inject → verify.

string dir = Path.Combine(Path.GetTempPath(), "xwb_core_test_" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(dir);
int failures = 0;

try
{
    // 1. Generate three small PCM WAVs with distinct payloads.
    byte[] MakePayload(byte seed, int length)
    {
        var data = new byte[length];
        for (int i = 0; i < length; i++)
            data[i] = (byte)(seed + i * 7);
        return data;
    }

    void WriteWav(string path, byte[] payload, int channels, int rate, int bits)
    {
        using var f = File.Create(path);
        using var w = new BinaryWriter(f);
        int blockAlign = channels * bits / 8;
        w.Write("RIFF"u8);
        w.Write((uint)(36 + payload.Length));
        w.Write("WAVE"u8);
        w.Write("fmt "u8);
        w.Write(16u);
        w.Write((ushort)1);
        w.Write((ushort)channels);
        w.Write((uint)rate);
        w.Write((uint)(rate * blockAlign));
        w.Write((ushort)blockAlign);
        w.Write((ushort)bits);
        w.Write("data"u8);
        w.Write((uint)payload.Length);
        w.Write(payload);
    }

    void Check(bool condition, string what)
    {
        if (condition)
        {
            Console.WriteLine($"  OK   {what}");
        }
        else
        {
            Console.WriteLine($"  FAIL {what}");
            failures++;
        }
    }

    var payloadA = MakePayload(11, 4000);
    var payloadB = MakePayload(37, 6002);
    var payloadC = MakePayload(73, 2500);

    string wavA = Path.Combine(dir, "a.wav");
    string wavB = Path.Combine(dir, "b.wav");
    string wavC = Path.Combine(dir, "c.wav");
    WriteWav(wavA, payloadA, 2, 44100, 16);
    WriteWav(wavB, payloadB, 1, 22050, 16);
    WriteWav(wavC, payloadC, 1, 44100, 8);

    // 2. Build an XWB.
    Console.WriteLine("Build:");
    string xwb = Path.Combine(dir, "test.xwb");
    XwbBuilder.Create([wavA, wavB, wavC], xwb, "TestBank");
    Check(File.Exists(xwb), "bank file created");

    // 3. Parse it back.
    Console.WriteLine("Parse:");
    var bank = XwbBank.Open(xwb);
    Check(bank.Version == 43, $"version 43 (got {bank.Version})");
    Check(bank.Tracks.Count == 3, $"3 tracks (got {bank.Tracks.Count})");
    Check(bank.Tracks[0].Codec == XwbCodec.Pcm, "track 0 is PCM");
    Check(bank.Tracks[0].Channels == 2, $"track 0 stereo (got {bank.Tracks[0].Channels})");
    Check(bank.Tracks[0].SampleRate == 44100, $"track 0 44100 Hz (got {bank.Tracks[0].SampleRate})");
    Check(bank.Tracks[0].Bits == 1, "track 0 16-bit");
    Check(bank.Tracks[1].SampleRate == 22050, $"track 1 22050 Hz (got {bank.Tracks[1].SampleRate})");
    Check(bank.Tracks[2].Bits == 0, "track 2 8-bit");
    Check(bank.Tracks[0].Size == (uint)payloadA.Length, "track 0 size matches");
    Check(bank.Tracks[1].Size == (uint)payloadB.Length, "track 1 size matches");

    double expectedDur = payloadA.Length / (44100.0 * 2 * 2);
    Check(Math.Abs(bank.Tracks[0].DurationSeconds - expectedDur) < 0.001, "track 0 duration matches");

    // 4. Extract all and compare payloads.
    Console.WriteLine("Extract:");
    string outDir = Path.Combine(dir, "out");
    var extracted = bank.ExtractAll(outDir);
    Check(extracted.Count == 3, $"3 files extracted (got {extracted.Count})");

    var reparsedA = WavFile.Parse(extracted[0]);
    Check(reparsedA.Data.AsSpan().SequenceEqual(payloadA), "track 0 payload round-trips");
    Check(reparsedA.Channels == 2 && reparsedA.SampleRate == 44100 && reparsedA.BitsPerSample == 16,
        "track 0 format round-trips");
    var reparsedB = WavFile.Parse(extracted[1]);
    Check(reparsedB.Data.AsSpan().SequenceEqual(payloadB), "track 1 payload round-trips");

    // 5. Custom names.
    Console.WriteLine("Naming:");
    var names = new Dictionary<string, string> { ["00000001"] = "friendly_name" };
    var named = bank.ExtractAll(Path.Combine(dir, "named"), names);
    Check(named.Any(p => Path.GetFileName(p) == "friendly_name.wav"), "custom track name applied");

    // 6. Inject: replace track 1 with payloadC's wav, save to a new file.
    Console.WriteLine("Inject:");
    string injected = Path.Combine(dir, "injected.xwb");
    XwbInjector.Rebuild(xwb, 1, wavC, injected);

    var bank2 = XwbBank.Open(injected);
    Check(bank2.Tracks.Count == 3, "injected bank still has 3 tracks");
    Check(bank2.Tracks[1].Size == (uint)payloadC.Length, $"track 1 resized (got {bank2.Tracks[1].Size}, want {payloadC.Length})");

    var extracted2 = bank2.ExtractAll(Path.Combine(dir, "out2"));
    var newB = WavFile.Parse(extracted2[1]);
    Check(newB.Data.AsSpan().SequenceEqual(payloadC), "replaced track carries new payload");
    var stillA = WavFile.Parse(extracted2[0]);
    Check(stillA.Data.AsSpan().SequenceEqual(payloadA), "untouched track 0 still intact");
    var stillC = WavFile.Parse(extracted2[2]);
    Check(stillC.Data.AsSpan().SequenceEqual(payloadC), "untouched track 2 still intact");

    Console.WriteLine();
    Console.WriteLine(failures == 0 ? "ALL TESTS PASSED" : $"{failures} TEST(S) FAILED");
}
finally
{
    try { Directory.Delete(dir, recursive: true); } catch { }
}

return failures == 0 ? 0 : 1;
