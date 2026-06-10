using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media.Animation;

namespace XwbStudio;

/// <summary>
/// Attached property that eases a ProgressBar to its new value instead of jumping,
/// so batch progress feels fluid.
/// </summary>
public static class Smooth
{
    public static readonly DependencyProperty TargetProperty = DependencyProperty.RegisterAttached(
        "Target", typeof(double), typeof(Smooth), new PropertyMetadata(0.0, OnTargetChanged));

    public static double GetTarget(DependencyObject obj) => (double)obj.GetValue(TargetProperty);

    public static void SetTarget(DependencyObject obj, double value) => obj.SetValue(TargetProperty, value);

    private static void OnTargetChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not ProgressBar bar)
            return;

        double newValue = (double)e.NewValue;
        if (newValue <= 0)
        {
            // Reset (new run starting) — snap instead of animating backwards.
            bar.BeginAnimation(RangeBase.ValueProperty, null);
            bar.Value = 0;
            return;
        }

        var animation = new DoubleAnimation(newValue, TimeSpan.FromMilliseconds(400))
        {
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut },
        };
        bar.BeginAnimation(RangeBase.ValueProperty, animation);
    }
}
