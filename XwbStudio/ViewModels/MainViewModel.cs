using System.Windows;

namespace XwbStudio.ViewModels;

public sealed class MainViewModel : ObservableObject
{
    private string _status = "Ready.";

    public MainViewModel()
    {
        Action<string> setStatus = s =>
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher is null || dispatcher.CheckAccess())
                Status = s;
            else
                dispatcher.Invoke(() => Status = s);
        };

        Extract = new ExtractViewModel(setStatus);
        Convert = new ConvertViewModel(setStatus);
        Inject = new InjectViewModel(setStatus);
    }

    public ExtractViewModel Extract { get; }
    public ConvertViewModel Convert { get; }
    public InjectViewModel Inject { get; }

    public string Status { get => _status; set => Set(ref _status, value); }

    /// <summary>Called when switching tabs — stops any audio preview, like the original.</summary>
    public void OnTabSwitched()
    {
        Extract.StopPreview();
        Inject.StopPreview();
    }
}
