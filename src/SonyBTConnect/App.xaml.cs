using System.Diagnostics;
using System.Windows;
using System.Windows.Threading;
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

        // Catch all unhandled exceptions - don't let the app crash
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;

        _bluetoothService = new BluetoothConnectionService();
        _startupService = new StartupService();
        _viewModel = new TrayIconViewModel(_bluetoothService, _startupService);

        _trayIcon = (TaskbarIcon)FindResource("TrayIcon");
        _trayIcon.DataContext = _viewModel;
        _trayIcon.ForceCreate();
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        Debug.WriteLine($"UI Exception: {e.Exception}");
        e.Handled = true;
    }

    private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        Debug.WriteLine($"Unhandled Exception: {e.ExceptionObject}");
    }

    private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        Debug.WriteLine($"Task Exception: {e.Exception}");
        e.SetObserved();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _trayIcon?.Dispose();
        _bluetoothService?.Dispose();
        base.OnExit(e);
    }
}
