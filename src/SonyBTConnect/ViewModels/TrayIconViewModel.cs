using System.Windows;
using System.Windows.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SonyBTConnect.Services;

namespace SonyBTConnect.ViewModels;

public partial class TrayIconViewModel : ObservableObject
{
    private readonly IBluetoothConnectionService _bluetoothService;
    private readonly IStartupService _startupService;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(TooltipText))]
    [NotifyPropertyChangedFor(nameof(IconBackground))]
    [NotifyPropertyChangedFor(nameof(StatusText))]
    private bool _isConnected;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ConnectCommand))]
    private bool _isConnecting;

    [ObservableProperty]
    private bool _isAutoStartEnabled;

    public string TooltipText => IsConnected
        ? "Sony WH-1000XM5: Polaczony"
        : "Sony WH-1000XM5: Rozlaczony (kliknij aby polaczyc)";

    public string StatusText => IsConnected
        ? "Polaczony"
        : IsConnecting ? "Laczenie..." : "Rozlaczony";

    // Kolor tla ikony: zielony = polaczony, czerwony = niepodlaczony
    public Brush IconBackground => IsConnected
        ? new SolidColorBrush(Color.FromRgb(76, 175, 80))   // Zielony (#4CAF50)
        : new SolidColorBrush(Color.FromRgb(244, 67, 54));  // Czerwony (#F44336)

    public Brush IconForeground => Brushes.White;

    public TrayIconViewModel(
        IBluetoothConnectionService bluetoothService,
        IStartupService startupService)
    {
        _bluetoothService = bluetoothService;
        _startupService = startupService;

        _isAutoStartEnabled = _startupService.IsAutoStartEnabled();
        _isConnected = _bluetoothService.IsConnected;

        _bluetoothService.ConnectionStatusChanged += OnConnectionStatusChanged;
        _bluetoothService.StartMonitoring();
    }

    private void OnConnectionStatusChanged(object? sender, bool connected)
    {
        Application.Current?.Dispatcher.Invoke(() =>
        {
            IsConnected = connected;
            IsConnecting = false;
        });
    }

    [RelayCommand(CanExecute = nameof(CanConnect))]
    private async Task ConnectAsync()
    {
        if (IsConnected || IsConnecting) return;

        IsConnecting = true;
        try
        {
            var result = await _bluetoothService.ConnectAsync();

            switch (result)
            {
                case ConnectionResult.Success:
                case ConnectionResult.AlreadyConnected:
                    // Status zaktualizowany przez event
                    break;

                case ConnectionResult.DeviceNotFound:
                    ShowBalloonTip("Sluchawki nie znalezione",
                        "Upewnij sie, ze sluchawki sa wlaczone i sparowane z tym komputerem.");
                    break;

                case ConnectionResult.ConnectionFailed:
                    ShowBalloonTip("Blad polaczenia",
                        "Nie udalo sie polaczyc ze sluchawkami. Sprobuj ponownie.");
                    break;

                case ConnectionResult.BluetoothError:
                    ShowBalloonTip("Blad Bluetooth",
                        "Sprawdz czy Bluetooth jest wlaczony na tym komputerze.");
                    break;
            }
        }
        finally
        {
            IsConnecting = false;
        }
    }

    private bool CanConnect() => !IsConnecting && !IsConnected;

    partial void OnIsAutoStartEnabledChanged(bool value)
    {
        if (value)
            _startupService.EnableAutoStart();
        else
            _startupService.DisableAutoStart();
    }

    [RelayCommand]
    private void Exit()
    {
        _bluetoothService.StopMonitoring();
        _bluetoothService.Dispose();
        Application.Current?.Shutdown();
    }

    private void ShowBalloonTip(string title, string message)
    {
        // Powiadomienie będzie obsługiwane przez TaskbarIcon w App.xaml
        // Na razie logujemy do debuggera
        System.Diagnostics.Debug.WriteLine($"Notification: {title} - {message}");
    }
}
