using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using XwbStudio.Core;
using XwbStudio.Services;

namespace XwbStudio.ViewModels;

public sealed class InjectViewModel : ObservableObject
{
    private readonly Action<string> _setStatus;
    private readonly PreviewPlayer _preview = new();

    private XwbBank? _bank;
    private TrackItem? _selectedTrack;
    private string _xwbPath = "";
    private string _wavPath = "";
    private string _outputFolder = "";
    private bool _saveToSeparateFolder;
    private string _selectedInfo = "No track selected";
    private string _statusText = "";
    private LogKind _statusKind = LogKind.Info;
    private bool _isWorking;
    private string _loadHint = "";

    public InjectViewModel(Action<string> setStatus)
    {
        _setStatus = setStatus;
        _preview.StateChanged += () => OnPropertyChanged(nameof(IsPreviewPlaying));

        BrowseXwbCommand = new RelayCommand(BrowseXwb);
        BrowseWavCommand = new RelayCommand(BrowseWav);
        BrowseOutputFolderCommand = new RelayCommand(BrowseOutputFolder);
        OpenOutputFolderCommand = new RelayCommand(OpenOutputFolder);
        PreviewCommand = new RelayCommand(async () => await TogglePreviewAsync());
        ReplaceCommand = new RelayCommand(Replace);
    }

    public ObservableCollection<TrackItem> Tracks { get; } = [];

    public RelayCommand BrowseXwbCommand { get; }
    public RelayCommand BrowseWavCommand { get; }
    public RelayCommand BrowseOutputFolderCommand { get; }
    public RelayCommand OpenOutputFolderCommand { get; }
    public RelayCommand PreviewCommand { get; }
    public RelayCommand ReplaceCommand { get; }

    public string XwbPath { get => _xwbPath; set => Set(ref _xwbPath, value); }

    public string WavPath
    {
        get => _wavPath;
        set
        {
            if (Set(ref _wavPath, value))
                OnPropertyChanged(nameof(IsArmed));
        }
    }

    public string OutputFolder { get => _outputFolder; set => Set(ref _outputFolder, value); }

    public bool SaveToSeparateFolder
    {
        get => _saveToSeparateFolder;
        set => Set(ref _saveToSeparateFolder, value);
    }

    public TrackItem? SelectedTrack
    {
        get => _selectedTrack;
        set
        {
            if (!Set(ref _selectedTrack, value))
                return;

            StopPreview();
            OnPropertyChanged(nameof(HasSelection));
            OnPropertyChanged(nameof(IsArmed));

            if (value is null)
            {
                SelectedInfo = "No track selected";
            }
            else
            {
                var t = value.Track;
                SelectedInfo = $"Track {t.Index:000}\n"
                             + $"Duration: {value.DurationText}\n"
                             + $"Size: {t.Size / 1024} KB\n"
                             + $"Codec: {t.CodecName}";
            }
        }
    }

    public bool HasSelection => _selectedTrack is not null;

    /// <summary>True when both a track and a replacement WAV are chosen — drives the pulse animation.</summary>
    public bool IsArmed => _selectedTrack is not null && File.Exists(_wavPath.Trim()) && !_isWorking;

    public string SelectedInfo { get => _selectedInfo; private set => Set(ref _selectedInfo, value); }
    public string StatusText { get => _statusText; private set => Set(ref _statusText, value); }
    public LogKind StatusKind { get => _statusKind; private set => Set(ref _statusKind, value); }
    public string LoadHint { get => _loadHint; private set => Set(ref _loadHint, value); }
    public bool IsPreviewPlaying => _preview.IsPlaying;

    public bool IsWorking
    {
        get => _isWorking;
        private set
        {
            if (Set(ref _isWorking, value))
            {
                OnPropertyChanged(nameof(IsIdle));
                OnPropertyChanged(nameof(IsArmed));
            }
        }
    }

    public bool IsIdle => !_isWorking;

    // ── Loading ──────────────────────────────────────────────────────────────

    private void BrowseXwb()
    {
        var dlg = new OpenFileDialog
        {
            Title = "Select XWB file to modify",
            Filter = "XWB files (*.xwb)|*.xwb|All files (*.*)|*.*",
        };
        if (dlg.ShowDialog() != true)
            return;

        XwbPath = dlg.FileName;
        LoadTracks(dlg.FileName);
    }

    public async void LoadTracks(string path)
    {
        StopPreview();
        Tracks.Clear();
        SelectedTrack = null;
        LoadHint = "Loading...";

        try
        {
            var bank = await Task.Run(() => XwbBank.Open(path));
            _bank = bank;
            foreach (var track in bank.Tracks)
                Tracks.Add(new TrackItem(track));
            LoadHint = $"{Tracks.Count} tracks loaded — select one to replace";
        }
        catch (Exception ex)
        {
            _bank = null;
            LoadHint = $"Error: {ex.Message}";
        }
    }

    private void BrowseWav()
    {
        var dlg = new OpenFileDialog
        {
            Title = "Select replacement WAV file",
            Filter = "WAV files (*.wav)|*.wav|All files (*.*)|*.*",
        };
        if (dlg.ShowDialog() == true)
            WavPath = dlg.FileName;
    }

    private void BrowseOutputFolder()
    {
        var dlg = new OpenFolderDialog { Title = "Select output folder" };
        if (dlg.ShowDialog() == true)
            OutputFolder = dlg.FolderName;
    }

    private void OpenOutputFolder()
    {
        string path = OutputFolder.Trim();
        if (path.Length > 0 && Directory.Exists(path))
            Process.Start("explorer.exe", path);
        else
            MessageBox.Show("No valid output folder selected.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    // ── Preview ──────────────────────────────────────────────────────────────

    private async Task TogglePreviewAsync()
    {
        if (_preview.IsPlaying)
        {
            StopPreview();
            return;
        }

        if (_selectedTrack is null || _bank is null)
            return;

        try
        {
            await _preview.PlayAsync(_bank, _selectedTrack.Track);
        }
        catch
        {
            StopPreview();
        }
    }

    public void StopPreview() => _preview.Stop();

    // ── Replace ──────────────────────────────────────────────────────────────

    private async void Replace()
    {
        if (_selectedTrack is null || _bank is null)
        {
            SetStatus("No track selected.", LogKind.Skip);
            return;
        }

        string wavPath = WavPath.Trim();
        if (wavPath.Length == 0 || !File.Exists(wavPath))
        {
            SetStatus("No valid WAV file selected.", LogKind.Skip);
            return;
        }

        string srcPath = _bank.FilePath;
        string xwbName = Path.GetFileName(srcPath);
        string outPath;
        if (SaveToSeparateFolder)
        {
            string outFolder = OutputFolder.Trim();
            if (outFolder.Length == 0 || !Directory.Exists(outFolder))
            {
                SetStatus("No valid output folder selected.", LogKind.Skip);
                return;
            }
            outPath = Path.Combine(outFolder, xwbName);
        }
        else
        {
            outPath = srcPath;
        }

        StopPreview();
        IsWorking = true;
        SetStatus("Rebuilding XWB...", LogKind.Info);

        int trackIndex = _selectedTrack.Track.Index;

        try
        {
            await Task.Run(() => XwbInjector.Rebuild(srcPath, trackIndex, wavPath, outPath));
            SetStatus(
                SaveToSeparateFolder ? $"Saved to:\n{Path.GetFileName(outPath)}" : "Original overwritten!",
                LogKind.Ok);
            _setStatus($"Track {trackIndex:000} replaced in {xwbName}.");

            // The bank on disk changed — reload so offsets stay correct.
            if (!SaveToSeparateFolder)
                LoadTracks(srcPath);
        }
        catch (Exception ex)
        {
            SetStatus($"Error: {ex.Message}", LogKind.Error);
        }
        finally
        {
            IsWorking = false;
        }
    }

    private void SetStatus(string text, LogKind kind)
    {
        StatusText = text;
        StatusKind = kind;
    }
}
