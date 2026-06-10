using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;
using XwbStudio.ViewModels;

namespace XwbStudio;

/// <summary>Maps a bool to the string "pulse" so ControlTemplate Tag-triggers can start pulse storyboards.</summary>
public sealed class PulseWhenTrueConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is true ? "pulse" : "";

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

/// <summary>Visible when a count is zero (empty-state hints).</summary>
public sealed class ZeroToVisibilityConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is int n && n == 0 ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

/// <summary>Maps a LogKind to its theme brush (for status labels).</summary>
public sealed class LogKindToBrushConverter : IValueConverter
{
    private static readonly Brush Info = Frozen(0x7A, 0x7A, 0x9A);
    private static readonly Brush Ok = Frozen(0x4E, 0xCC, 0xA3);
    private static readonly Brush Skip = Frozen(0xF5, 0xA6, 0x23);
    private static readonly Brush Error = Frozen(0xE9, 0x45, 0x60);
    private static readonly Brush Heading = Frozen(0xEA, 0xEA, 0xEA);

    private static Brush Frozen(byte r, byte g, byte b)
    {
        var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
        brush.Freeze();
        return brush;
    }

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value switch
        {
            LogKind.Ok => Ok,
            LogKind.Skip => Skip,
            LogKind.Error => Error,
            LogKind.Heading => Heading,
            _ => Info,
        };

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
