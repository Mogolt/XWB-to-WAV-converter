using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Threading;
using XwbStudio.Core;

namespace XwbStudio.Services;

/// <summary>
/// Plays one track at a time by extracting it to a temp WAV and handing it
/// to winmm. Auto-stops after the track's duration and cleans up the temp file.
/// </summary>
public sealed class PreviewPlayer
{
    [DllImport("winmm.dll", CharSet = CharSet.Unicode)]
    private static extern bool PlaySound(string? pszSound, IntPtr hmod, uint fdwSound);

    private const uint SND_ASYNC = 0x0001;
    private const uint SND_NODEFAULT = 0x0002;
    private const uint SND_PURGE = 0x0040;
    private const uint SND_FILENAME = 0x00020000;

    private DispatcherTimer? _timer;
    private string? _tempFile;

    public bool IsPlaying { get; private set; }

    /// <summary>Raised on the UI thread whenever playback starts or stops.</summary>
    public event Action? StateChanged;

    public async Task PlayAsync(XwbBank bank, XwbTrack track)
    {
        Stop();

        string temp = Path.Combine(Path.GetTempPath(), $"xwb_preview_{Guid.NewGuid():N}.wav");
        await Task.Run(() => bank.ExtractTrack(track, temp));

        _tempFile = temp;
        PlaySound(temp, IntPtr.Zero, SND_FILENAME | SND_ASYNC | SND_NODEFAULT);
        IsPlaying = true;
        StateChanged?.Invoke();

        _timer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(track.DurationSeconds * 1000 + 1000),
        };
        _timer.Tick += (_, _) => Stop();
        _timer.Start();
    }

    public void Stop()
    {
        _timer?.Stop();
        _timer = null;

        if (IsPlaying)
        {
            PlaySound(null, IntPtr.Zero, SND_PURGE);
            IsPlaying = false;
            StateChanged?.Invoke();
        }

        var temp = _tempFile;
        _tempFile = null;
        if (temp is not null)
        {
            // Give winmm a moment to release the file before deleting it.
            _ = Task.Delay(500).ContinueWith(_ =>
            {
                try { File.Delete(temp); } catch { /* best effort */ }
            });
        }
    }
}
