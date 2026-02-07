using System.Windows;
using H.NotifyIcon;
using SonyBTConnect.Services;
using SonyBTConnect.ViewModels;

namespace SonyBTConnect;

public partial class App : Application
{
    private TaskbarIcon? _trayIcon;
    private IBluetoothConnectionService? _bluetoothService;
    private IStartupService? _startupService;
    private TrayIconViewModel? _viewModel;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Utworz serwisy
        _bluetoothService = new BluetoothConnectionService();
        _startupService = new StartupService();

        // Utworz ViewModel
        _viewModel = new TrayIconViewModel(_bluetoothService, _startupService);

        // Pobierz TaskbarIcon z zasobow i ustaw DataContext
        _trayIcon = (TaskbarIcon)FindResource("TrayIcon");
        _trayIcon.DataContext = _viewModel;

        // Wymus utworzenie ikony
        _trayIcon.ForceCreate();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _trayIcon?.Dispose();
        _bluetoothService?.Dispose();
        base.OnExit(e);
    }
}
