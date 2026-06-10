using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using XwbStudio.Core;

namespace XwbStudio.ViewModels;

public sealed class WavEntry(string path) : ObservableObject
{
    private bool _isSelected;

    public string Path { get; } = path;
    public string Name => System.IO.Path.GetFileName(Path);
    public bool IsSelected { get => _isSelected; set => Set(ref _isSelected, value); }
}

public sealed class ConvertViewModel : ObservableObject
{
    private readonly Action<string> _setStatus;

    private string _outputPath = "";
    private string _bankName = "CustomBank";
    private string _statusText = "";
    private LogKind _statusKind = LogKind.Info;
    private bool _isConverting;

    public ConvertViewModel(Action<string> setStatus)
    {
        _setStatus = setStatus;

        AddFilesCommand = new RelayCommand(AddFiles);
        AddFolderCommand = new RelayCommand(AddFolder);
        RemoveSelectedCommand = new RelayCommand(RemoveSelected);
        ClearAllCommand = new RelayCommand(() => Files.Clear());
        BrowseOutputCommand = new RelayCommand(BrowseOutput);
        ConvertCommand = new RelayCommand(Convert);
        OpenFolderCommand = new RelayCommand(OpenFolder);

        Files.CollectionChanged += (_, _) => OnPropertyChanged(nameof(FileCountText));
    }

    public ObservableCollection<WavEntry> Files { get; } = [];

    public RelayCommand AddFilesCommand { get; }
    public RelayCommand AddFolderCommand { get; }
    public RelayCommand RemoveSelectedCommand { get; }
    public RelayCommand ClearAllCommand { get; }
    public RelayCommand BrowseOutputCommand { get; }
    public RelayCommand ConvertCommand { get; }
    public RelayCommand OpenFolderCommand { get; }

    public string OutputPath { get => _outputPath; set => Set(ref _outputPath, value); }
    public string BankName { get => _bankName; set => Set(ref _bankName, value); }
    public string StatusText { get => _statusText; private set => Set(ref _statusText, value); }
    public LogKind StatusKind { get => _statusKind; private set => Set(ref _statusKind, value); }

    public bool IsConverting
    {
        get => _isConverting;
        private set
        {
            if (Set(ref _isConverting, value))
                OnPropertyChanged(nameof(IsIdle));
        }
    }

    public bool IsIdle => !_isConverting;

    public string FileCountText => Files.Count switch
    {
        0 => "No files added yet — add WAVs or drop them here",
        1 => "1 file",
        var n => $"{n} files",
    };

    public void AddPaths(IEnumerable<string> paths)
    {
        foreach (var path in paths)
        {
            if (Directory.Exists(path))
            {
                foreach (var wav in Directory.EnumerateFiles(path, "*.wav").OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
                    AddOne(wav);
            }
            else if (path.EndsWith(".wav", StringComparison.OrdinalIgnoreCase))
            {
                AddOne(path);
            }
        }
    }

    private void AddOne(string path)
    {
        if (!Files.Any(f => string.Equals(f.Path, path, StringComparison.OrdinalIgnoreCase)))
            Files.Add(new WavEntry(path));
    }

    private void AddFiles()
    {
        var dlg = new OpenFileDialog
        {
            Title = "Select WAV files",
            Filter = "WAV files (*.wav)|*.wav|All files (*.*)|*.*",
            Multiselect = true,
        };
        if (dlg.ShowDialog() == true)
            AddPaths(dlg.FileNames);
    }

    private void AddFolder()
    {
        var dlg = new OpenFolderDialog { Title = "Select folder containing WAV files" };
        if (dlg.ShowDialog() == true)
            AddPaths([dlg.FolderName]);
    }

    private void RemoveSelected()
    {
        foreach (var entry in Files.Where(f => f.IsSelected).ToList())
            Files.Remove(entry);
    }

    private void BrowseOutput()
    {
        var dlg = new SaveFileDialog
        {
            Title = "Save XWB file as",
            DefaultExt = ".xwb",
            Filter = "XWB files (*.xwb)|*.xwb",
        };
        if (dlg.ShowDialog() == true)
            OutputPath = dlg.FileName;
    }

    private void OpenFolder()
    {
        string output = OutputPath.Trim();
        string? folder = output.Length > 0 ? Path.GetDirectoryName(output) : null;
        if (!string.IsNullOrEmpty(folder) && Directory.Exists(folder))
            Process.Start("explorer.exe", folder);
        else
            MessageBox.Show("No valid output path set.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private async void Convert()
    {
        if (Files.Count == 0)
        {
            MessageBox.Show("Add at least one WAV file first.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }
        string outPath = OutputPath.Trim();
        if (outPath.Length == 0)
        {
            MessageBox.Show("Please set an output XWB path.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        string bankName = BankName.Trim();
        if (bankName.Length == 0)
            bankName = "CustomBank";

        IsConverting = true;
        StatusText = "";

        var files = Files.Select(f => f.Path).ToList();

        try
        {
            await Task.Run(() => XwbBuilder.Create(files, outPath, bankName));
            StatusText = $"Done! Saved: {Path.GetFileName(outPath)}";
            StatusKind = LogKind.Ok;
            _setStatus($"Converted {files.Count} WAV file(s) to {Path.GetFileName(outPath)}.");
        }
        catch (Exception ex)
        {
            StatusText = $"Error: {ex.Message}";
            StatusKind = LogKind.Error;
        }
        finally
        {
            IsConverting = false;
        }
    }
}
