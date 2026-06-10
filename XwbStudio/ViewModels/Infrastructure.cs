using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using System.Windows.Media;
using XwbStudio.Core;

namespace XwbStudio.ViewModels;

public abstract class ObservableObject : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    protected bool Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
            return false;
        field = value;
        OnPropertyChanged(name);
        return true;
    }
}

public sealed class RelayCommand(Action execute, Func<bool>? canExecute = null) : ICommand
{
    public event EventHandler? CanExecuteChanged;

    public bool CanExecute(object? parameter) => canExecute?.Invoke() ?? true;

    public void Execute(object? parameter) => execute();

    public void RaiseCanExecuteChanged()
        => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
}

public enum LogKind
{
    Info,
    Ok,
    Skip,
    Error,
    Heading,
}

/// <summary>A single colored line in the activity log. Mutable so a line can be completed in place.</summary>
public sealed class LogEntry(string text, LogKind kind) : ObservableObject
{
    private string _text = text;
    private LogKind _kind = kind;

    public string Text { get => _text; set => Set(ref _text, value); }
    public LogKind Kind { get => _kind; set => Set(ref _kind, value); }
}

/// <summary>List row wrapper around an <see cref="XwbTrack"/> with selection state.</summary>
public sealed class TrackItem(XwbTrack track) : ObservableObject
{
    private static readonly Brush PcmBrush = new SolidColorBrush(Color.FromRgb(0x4E, 0xCC, 0xA3));
    private static readonly Brush AdpcmBrush = new SolidColorBrush(Color.FromRgb(0xF5, 0xA6, 0x23));
    private static readonly Brush XmaBrush = new SolidColorBrush(Color.FromRgb(0x7A, 0x9A, 0xDA));
    private static readonly Brush WmaBrush = new SolidColorBrush(Color.FromRgb(0xE9, 0x45, 0x60));

    static TrackItem()
    {
        PcmBrush.Freeze();
        AdpcmBrush.Freeze();
        XmaBrush.Freeze();
        WmaBrush.Freeze();
    }

    private bool _isSelected;

    public XwbTrack Track { get; } = track;

    public bool IsSelected { get => _isSelected; set => Set(ref _isSelected, value); }

    public string IndexText => Track.Index.ToString("000");

    public string DurationText => Track.DurationSeconds > 0
        ? $"{(int)(Track.DurationSeconds / 60)}:{(int)(Track.DurationSeconds % 60):00}"
        : "?";

    public string SizeText => $"{Track.Size / 1024:N0} KB";

    public string CodecName => Track.CodecName;

    public Brush CodecBrush => Track.Codec switch
    {
        XwbCodec.Pcm => PcmBrush,
        XwbCodec.Adpcm => AdpcmBrush,
        XwbCodec.Xma => XmaBrush,
        _ => WmaBrush,
    };
}
