using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;
using XwbStudio.ViewModels;

namespace XwbStudio.Views;

public partial class ExtractView : UserControl
{
    private ExtractViewModel? _vm;

    public ExtractView()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
    }

    private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (_vm is not null)
        {
            ((INotifyCollectionChanged)_vm.Log).CollectionChanged -= OnLogChanged;
            _vm.PropertyChanged -= OnVmPropertyChanged;
        }

        _vm = DataContext as ExtractViewModel;
        if (_vm is not null)
        {
            ((INotifyCollectionChanged)_vm.Log).CollectionChanged += OnLogChanged;
            _vm.PropertyChanged += OnVmPropertyChanged;
            SetPanelOpen(_vm.IsBrowserOpen, animate: false);
        }
    }

    private void OnLogChanged(object? sender, NotifyCollectionChangedEventArgs e)
        => LogScroll.ScrollToEnd();

    private void OnVmPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ExtractViewModel.IsBrowserOpen) && _vm is not null)
            SetPanelOpen(_vm.IsBrowserOpen, animate: true);
    }

    private void SetPanelOpen(bool open, bool animate)
    {
        double targetWidth = open ? 314 : 0;
        double targetOpacity = open ? 1 : 0;

        if (!animate)
        {
            BrowserPanel.Width = targetWidth;
            BrowserPanel.Opacity = targetOpacity;
            return;
        }

        var ease = new CubicEase { EasingMode = EasingMode.EaseInOut };
        BrowserPanel.BeginAnimation(WidthProperty,
            new DoubleAnimation(targetWidth, TimeSpan.FromMilliseconds(300)) { EasingFunction = ease });
        BrowserPanel.BeginAnimation(OpacityProperty,
            new DoubleAnimation(targetOpacity, TimeSpan.FromMilliseconds(open ? 360 : 200)));
    }

    private void RecentClick(object sender, RoutedEventArgs e)
    {
        if (_vm is not null && sender is Button { DataContext: string path })
            _vm.UseRecentFolder(path);
    }
}
