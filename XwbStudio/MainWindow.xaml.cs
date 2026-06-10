using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using XwbStudio.ViewModels;

namespace XwbStudio;

public partial class MainWindow : Window
{
    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

    private const int DwmwaUseImmersiveDarkMode = 20;

    private readonly MainViewModel _vm = new();
    private UserControl[] _views = [];
    private RadioButton[] _tabs = [];
    private int _currentTab;
    private bool _ready;

    public MainWindow()
    {
        InitializeComponent();
        DataContext = _vm;

        _views = [ExtractTab, ConvertTab, InjectTab];
        _tabs = [TabExtract, TabConvert, TabInject];

        SourceInitialized += (_, _) => EnableDarkTitleBar();
        ContentRendered += (_, _) => OnFirstRender();
        SizeChanged += (_, _) => { if (_ready) MoveIndicator(_currentTab, animate: false); };
    }

    private void EnableDarkTitleBar()
    {
        int on = 1;
        var hwnd = new WindowInteropHelper(this).Handle;
        _ = DwmSetWindowAttribute(hwnd, DwmwaUseImmersiveDarkMode, ref on, sizeof(int));
    }

    private void OnFirstRender()
    {
        if (_ready)
            return;
        _ready = true;

        MoveIndicator(0, animate: false);
        BeginAnimation(OpacityProperty, new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(320)));
    }

    // ── Tabs ─────────────────────────────────────────────────────────────────

    private void TabChecked(object sender, RoutedEventArgs e)
    {
        if (_views.Length == 0)
            return;

        int index = Array.IndexOf(_tabs, (RadioButton)sender);
        if (index < 0 || index == _currentTab && _ready)
            return;

        _vm.OnTabSwitched();
        _currentTab = index;

        for (int i = 0; i < _views.Length; i++)
            _views[i].Visibility = i == index ? Visibility.Visible : Visibility.Collapsed;

        if (!_ready)
            return;

        MoveIndicator(index, animate: true);

        // Fade + slide the incoming view.
        var view = _views[index];
        view.Opacity = 0;
        var transform = (TranslateTransform)view.RenderTransform;
        transform.BeginAnimation(TranslateTransform.YProperty, null);
        transform.Y = 14;

        view.BeginAnimation(OpacityProperty,
            new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(200)));
        transform.BeginAnimation(TranslateTransform.YProperty,
            new DoubleAnimation(14, 0, TimeSpan.FromMilliseconds(300))
            {
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut },
            });
    }

    private void MoveIndicator(int index, bool animate)
    {
        var tab = _tabs[index];
        if (tab.ActualWidth <= 0)
            return;

        Point origin = tab.TranslatePoint(new Point(0, 0), IndicatorCanvas);
        double inset = 14;
        double left = origin.X + inset;
        double width = Math.Max(0, tab.ActualWidth - inset * 2);

        if (!animate)
        {
            TabIndicator.BeginAnimation(Canvas.LeftProperty, null);
            TabIndicator.BeginAnimation(WidthProperty, null);
            Canvas.SetLeft(TabIndicator, left);
            TabIndicator.Width = width;
            return;
        }

        var ease = new CubicEase { EasingMode = EasingMode.EaseOut };
        TabIndicator.BeginAnimation(Canvas.LeftProperty,
            new DoubleAnimation(left, TimeSpan.FromMilliseconds(280)) { EasingFunction = ease });
        TabIndicator.BeginAnimation(WidthProperty,
            new DoubleAnimation(width, TimeSpan.FromMilliseconds(280)) { EasingFunction = ease });
    }

    // ── Help overlay ─────────────────────────────────────────────────────────

    private void HelpClick(object sender, RoutedEventArgs e) => ShowHelp();

    private void HelpScrimClick(object sender, MouseButtonEventArgs e) => HideHelp();

    private void HelpCloseClick(object sender, RoutedEventArgs e) => HideHelp();

    private void ShowHelp()
    {
        HelpOverlay.Visibility = Visibility.Visible;
        HelpOverlay.BeginAnimation(OpacityProperty,
            new DoubleAnimation(1, TimeSpan.FromMilliseconds(180)));

        var transform = (TranslateTransform)HelpCard.RenderTransform;
        transform.BeginAnimation(TranslateTransform.YProperty,
            new DoubleAnimation(26, 0, TimeSpan.FromMilliseconds(280))
            {
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut },
            });
    }

    private void HideHelp()
    {
        var fade = new DoubleAnimation(0, TimeSpan.FromMilliseconds(160));
        fade.Completed += (_, _) => HelpOverlay.Visibility = Visibility.Collapsed;
        HelpOverlay.BeginAnimation(OpacityProperty, fade);
    }
}
