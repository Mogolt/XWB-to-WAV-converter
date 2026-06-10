using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using XwbStudio.Core;
using XwbStudio.Services;

namespace XwbStudio.ViewModels;

public sealed class ExtractViewModel : ObservableObject
{
    private readonly Action<string> _setStatus;
    private readonly RecentFoldersService _recent = new();
    private readonly PreviewPlayer _preview = new();

    private CancellationTokenSource? _cts;
    private Dictionary<string, Dictionary<string, string>> _trackNames = new();
    private XwbBank? _browserBank;

    private string _inputFolder = "";
    private string _outputFolder = "";
    private string _configPath = "";
    private bool _isRunning;
    private double _progress;
    private double _progressMax = 1;
    private string _progressText = "";
    private bool _isBrowserOpen;
    private string _browserXwbPath = "";
    private string _browserHint = "Load an XWB to browse";
    private LogKind _browserHintKind = LogKind.Info;
    private string _renameTo = "";
    private int _selectedCount;

    public ExtractViewModel(Action<string> setStatus)
    {
        _setStatus = setStatus;
        _preview.StateChanged += () => OnPropertyChanged(nameof(IsPreviewPlaying));

        RecentFolders = new ObservableCollection<string>(_recent.Folders);

        BrowseInputCommand = new RelayCommand(BrowseInput);
        BrowseOutputCommand = new RelayCommand(BrowseOutput);
        BrowseConfigCommand = new RelayCommand(BrowseConfig);
        ExtractAllCommand = new RelayCommand(ExtractAll);
        StopCommand = new RelayCommand(() => { _cts?.Cancel(); _setStatus("Stopping after current file..."); });
        OpenOutputCommand = new RelayCommand(OpenOutput);
        CreateConfigCommand = new RelayCommand(CreateConfigTemplate);
        BrowseSingleXwbCommand = new RelayCommand(BrowseSingleXwb);
        PreviewCommand = new RelayCommand(async () => await TogglePreviewAsync());
        ExtractSelectedCommand = new RelayCommand(ExtractSelected);
        UseRecentCommand = new RelayCommand(() => { });

        TryLoadConfigFromAppDir();
    }

    public ObservableCollection<LogEntry> Log { get; } = [];
    public ObservableCollection<string> RecentFolders { get; }
    public ObservableCollection<TrackItem> Tracks { get; } = [];

    public RelayCommand BrowseInputCommand { get; }
    public RelayCommand BrowseOutputCommand { get; }
    public RelayCommand BrowseConfigCommand { get; }
    public RelayCommand ExtractAllCommand { get; }
    public RelayCommand StopCommand { get; }
    public RelayCommand OpenOutputCommand { get; }
    public RelayCommand CreateConfigCommand { get; }
    public RelayCommand BrowseSingleXwbCommand { get; }
    public RelayCommand PreviewCommand { get; }
    public RelayCommand ExtractSelectedCommand { get; }
    public RelayCommand UseRecentCommand { get; }

    public string InputFolder { get => _inputFolder; set => Set(ref _inputFolder, value); }
    public string OutputFolder { get => _outputFolder; set => Set(ref _outputFolder, value); }
    public string ConfigPath { get => _configPath; set => Set(ref _configPath, value); }

    public bool IsRunning
    {
        get => _isRunning;
        private set
        {
            if (Set(ref _isRunning, value))
                OnPropertyChanged(nameof(IsIdle));
        }
    }

    public bool IsIdle => !_isRunning;

    public double Progress { get => _progress; private set => Set(ref _progress, value); }
    public double ProgressMax { get => _progressMax; private set => Set(ref _progressMax, value); }
    public string ProgressText { get => _progressText; private set => Set(ref _progressText, value); }

    public bool IsBrowserOpen
    {
        get => _isBrowserOpen;
        set
        {
            if (Set(ref _isBrowserOpen, value) && !value)
                StopPreview();
        }
    }

    public string BrowserXwbPath { get => _browserXwbPath; set => Set(ref _browserXwbPath, value); }
    public string BrowserHint { get => _browserHint; private set => Set(ref _browserHint, value); }
    public LogKind BrowserHintKind { get => _browserHintKind; private set => Set(ref _browserHintKind, value); }
    public string RenameTo { get => _renameTo; set => Set(ref _renameTo, value); }

    public bool IsPreviewPlaying => _preview.IsPlaying;

    public int SelectedCount
    {
        get => _selectedCount;
        private set
        {
            if (Set(ref _selectedCount, value))
            {
                OnPropertyChanged(nameof(HasSelection));
                OnPropertyChanged(nameof(IsSingleSelection));
            }
        }
    }

    public bool HasSelection => _selectedCount > 0;
    public bool IsSingleSelection => _selectedCount == 1;

    // ── Browsing ─────────────────────────────────────────────────────────────

    private void BrowseInput()
    {
        var dlg = new OpenFolderDialog { Title = "Select folder containing .xwb files" };
        if (dlg.ShowDialog() == true)
            InputFolder = dlg.FolderName;
    }

    private void BrowseOutput()
    {
        var dlg = new OpenFolderDialog { Title = "Select output folder" };
        if (dlg.ShowDialog() == true)
            OutputFolder = dlg.FolderName;
    }

    private void BrowseConfig()
    {
        var dlg = new OpenFileDialog
        {
            Title = "Select config.json",
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
        };
        if (dlg.ShowDialog() == true)
        {
            ConfigPath = dlg.FileName;
            LoadConfig(dlg.FileName);
        }
    }

    public void UseRecentFolder(string path)
    {
        InputFolder = path;
        AddLog($"Recent folder selected: {path}", LogKind.Info);
    }

    private void OpenOutput()
    {
        string path = OutputFolder.Trim();
        if (path.Length > 0 && Directory.Exists(path))
            Process.Start("explorer.exe", path);
        else
            MessageBox.Show("No valid output folder selected.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    // ── Config ───────────────────────────────────────────────────────────────

    private void TryLoadConfigFromAppDir()
    {
        string local = Path.Combine(AppContext.BaseDirectory, TrackNameConfig.FileName);
        if (File.Exists(local))
        {
            ConfigPath = local;
            LoadConfig(local);
        }
    }

    private void LoadConfig(string path)
    {
        try
        {
            _trackNames = TrackNameConfig.Load(path);
            AddLog($"Config loaded: {Path.GetFileName(path)}", LogKind.Info);
            int total = _trackNames.Values.Sum(v => v.Count);
            AddLog($"  {total} custom track name(s) found across {_trackNames.Count} bank(s)", LogKind.Info);
        }
        catch (Exception ex)
        {
            AddLog($"Could not load config: {ex.Message}", LogKind.Skip);
            _trackNames = new Dictionary<string, Dictionary<string, string>>();
        }
    }

    private void CreateConfigTemplate()
    {
        var dlg = new SaveFileDialog
        {
            Title = "Save config template as...",
            DefaultExt = ".json",
            Filter = "JSON files (*.json)|*.json",
            FileName = TrackNameConfig.FileName,
        };
        if (dlg.ShowDialog() != true)
            return;

        try
        {
            TrackNameConfig.WriteTemplate(dlg.FileName);
            ConfigPath = dlg.FileName;
            AddLog($"Config template created: {dlg.FileName}", LogKind.Ok);
            AddLog("  Open it in any text editor to add custom track names.", LogKind.Info);
            AddLog("  Under \"track_names\", add entries like:", LogKind.Info);
            AddLog("    \"bio4bgm\": { \"00000000\": \"main_menu\", \"00000001\": \"village\" }", LogKind.Info);
        }
        catch (Exception ex)
        {
            AddLog($"Failed to create config: {ex.Message}", LogKind.Error);
        }
    }

    // ── Batch extraction ─────────────────────────────────────────────────────

    private async void ExtractAll()
    {
        string input = InputFolder.Trim();
        string output = OutputFolder.Trim();

        if (input.Length == 0 || !Directory.Exists(input))
        {
            AddLog("No valid XWB input folder selected.", LogKind.Skip);
            return;
        }
        if (output.Length == 0)
        {
            AddLog("No output folder set.", LogKind.Skip);
            return;
        }

        var xwbFiles = Directory.EnumerateFiles(input, "*.xwb")
            .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (xwbFiles.Count == 0)
        {
            AddLog($"No .xwb files found in: {input}", LogKind.Skip);
            return;
        }

        IsRunning = true;
        _cts = new CancellationTokenSource();
        Progress = 0;
        ProgressMax = xwbFiles.Count;
        ProgressText = $"0 / {xwbFiles.Count}";
        Log.Clear();

        string configPath = ConfigPath.Trim();
        if (configPath.Length > 0 && File.Exists(configPath))
            LoadConfig(configPath);

        _recent.Add(input);
        RecentFolders.Clear();
        foreach (var folder in _recent.Folders)
            RecentFolders.Add(folder);

        var token = _cts.Token;
        await Task.Run(() => RunExtraction(input, output, xwbFiles, token));

        IsRunning = false;
    }

    private void RunExtraction(string inputFolder, string outputFolder, List<string> xwbFiles, CancellationToken ct)
    {
        int total = xwbFiles.Count;
        int okCount = 0;
        int failCount = 0;
        var okLines = new List<string>();
        var failLines = new List<string>();

        AddLog($"Starting extraction of {total} XWB files...\n", LogKind.Heading);

        for (int i = 0; i < total; i++)
        {
            if (ct.IsCancellationRequested)
            {
                AddLog("\nStopped by user.", LogKind.Skip);
                break;
            }

            string xwbPath = xwbFiles[i];
            string fname = Path.GetFileName(xwbPath);
            string baseName = Path.GetFileNameWithoutExtension(fname);
            string outDir = Path.Combine(outputFolder, baseName);
            double sizeMb = new FileInfo(xwbPath).Length / (1024.0 * 1024.0);

            _setStatus($"Processing {i + 1}/{total}: {fname}");
            var line = AddLog($"[{i + 1}/{total}]  {fname}  ({sizeMb:F1} MB) ... ", LogKind.Info);

            _trackNames.TryGetValue(baseName, out var namesForBank);

            try
            {
                var bank = XwbBank.Open(xwbPath);
                var extracted = bank.ExtractAll(outDir, namesForBank, ct);
                if (extracted.Count > 0)
                {
                    UpdateLog(line, line.Text + $"OK  ({extracted.Count} tracks)", LogKind.Ok);
                    okCount++;
                    okLines.Add(fname);
                }
                else
                {
                    UpdateLog(line, line.Text + "SKIPPED  (no tracks found)", LogKind.Skip);
                    failCount++;
                    failLines.Add($"{fname} - no tracks found");
                }
            }
            catch (Exception ex)
            {
                UpdateLog(line, line.Text + $"SKIPPED  ({ex.Message})", LogKind.Skip);
                failCount++;
                failLines.Add($"{fname} - {ex.Message}");
            }

            int done = i + 1;
            RunOnUi(() =>
            {
                Progress = done;
                ProgressText = $"{done} / {total}";
            });
        }

        try
        {
            File.WriteAllLines(Path.Combine(outputFolder, "converted_ok.txt"), okLines);
            File.WriteAllLines(Path.Combine(outputFolder, "failed.txt"), failLines);
        }
        catch
        {
            // Report files are best effort.
        }

        AddLog($"\n{new string('─', 50)}", LogKind.Info);
        AddLog($"Done!  Extracted: {okCount}   Skipped: {failCount}", LogKind.Heading);
        AddLog($"Output folder: {outputFolder}", LogKind.Info);
        if (failCount > 0)
            AddLog("Check failed.txt for details on skipped files.", LogKind.Skip);

        _setStatus($"Done!  {okCount} extracted, {failCount} skipped.");
    }

    // ── Track browser ────────────────────────────────────────────────────────

    private void BrowseSingleXwb()
    {
        var dlg = new OpenFileDialog
        {
            Title = "Select XWB file to browse",
            Filter = "XWB files (*.xwb)|*.xwb|All files (*.*)|*.*",
        };
        if (dlg.ShowDialog() != true)
            return;

        BrowserXwbPath = dlg.FileName;
        LoadBrowserTracks(dlg.FileName);
    }

    private async void LoadBrowserTracks(string path)
    {
        StopPreview();
        Tracks.Clear();
        SelectedCount = 0;
        SetBrowserHint("Loading tracks...", LogKind.Info);

        try
        {
            var bank = await Task.Run(() => XwbBank.Open(path));
            _browserBank = bank;
            foreach (var track in bank.Tracks)
            {
                var item = new TrackItem(track);
                item.PropertyChanged += OnTrackItemChanged;
                Tracks.Add(item);
            }
            SetBrowserHint($"{Tracks.Count} tracks  —  Ctrl+click or drag to multi-select", LogKind.Info);
        }
        catch (Exception ex)
        {
            _browserBank = null;
            SetBrowserHint($"Error: {ex.Message}", LogKind.Error);
        }
    }

    private void OnTrackItemChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName != nameof(TrackItem.IsSelected))
            return;

        // Selection change always stops a running preview, like the original.
        StopPreview();

        int count = Tracks.Count(t => t.IsSelected);
        SelectedCount = count;
        if (count == 1)
            RenameTo = "";
        if (count > 0)
            SetBrowserHint($"{count} track{(count > 1 ? "s" : "")} selected", LogKind.Ok);
    }

    private async Task TogglePreviewAsync()
    {
        if (_preview.IsPlaying)
        {
            StopPreview();
            return;
        }

        var selected = Tracks.FirstOrDefault(t => t.IsSelected);
        if (selected is null || _browserBank is null)
            return;

        try
        {
            await _preview.PlayAsync(_browserBank, selected.Track);
        }
        catch
        {
            StopPreview();
        }
    }

    public void StopPreview() => _preview.Stop();

    private async void ExtractSelected()
    {
        var selected = Tracks.Where(t => t.IsSelected).Select(t => t.Track).ToList();
        if (selected.Count == 0 || _browserBank is null)
        {
            AddLog("No tracks selected — load an XWB and select tracks first.", LogKind.Skip);
            return;
        }

        string outputFolder = OutputFolder.Trim();
        if (outputFolder.Length == 0)
        {
            AddLog("No output folder set.", LogKind.Skip);
            return;
        }

        var bank = _browserBank;
        string xwbBase = Path.GetFileNameWithoutExtension(bank.FilePath);
        string outDir = Path.Combine(outputFolder, xwbBase);
        Directory.CreateDirectory(outDir);

        _trackNames.TryGetValue(xwbBase, out var namesForBank);
        string customName = selected.Count == 1 ? RenameTo.Trim() : "";

        IsRunning = true;
        Log.Clear();
        AddLog($"Extracting {selected.Count} selected track(s) from {xwbBase}.xwb...\n", LogKind.Heading);

        await Task.Run(() =>
        {
            int ok = 0, fail = 0;
            foreach (var track in selected)
            {
                string baseName = customName.Length > 0
                    ? customName
                    : namesForBank?.GetValueOrDefault(track.HexName) ?? track.HexName;
                string ext = track.Codec == XwbCodec.Wma ? ".wma" : ".wav";
                string outPath = Path.Combine(outDir, baseName + ext);

                try
                {
                    bank.ExtractTrack(track, outPath);
                    AddLog($"  {baseName}{ext}  OK", LogKind.Ok);
                    ok++;
                }
                catch (Exception ex)
                {
                    AddLog($"  {baseName}  FAILED  ({ex.Message})", LogKind.Error);
                    fail++;
                }
            }

            AddLog($"\nDone!  {ok} extracted, {fail} failed.\nOutput: {outDir}", LogKind.Heading);
            _setStatus($"Done!  {ok} extracted, {fail} failed.");
        });

        IsRunning = false;
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private void SetBrowserHint(string text, LogKind kind)
    {
        BrowserHint = text;
        BrowserHintKind = kind;
    }

    private LogEntry AddLog(string text, LogKind kind)
    {
        LogEntry? entry = null;
        RunOnUi(() =>
        {
            entry = new LogEntry(text, kind);
            Log.Add(entry);
        });
        return entry!;
    }

    private void UpdateLog(LogEntry entry, string text, LogKind kind)
        => RunOnUi(() =>
        {
            entry.Text = text;
            entry.Kind = kind;
        });

    private static void RunOnUi(Action action)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is null || dispatcher.CheckAccess())
            action();
        else
            dispatcher.Invoke(action);
    }
}
