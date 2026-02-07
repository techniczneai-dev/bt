using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using SonyBTConnect.Services;

namespace SonyBTConnect.ViewModels;

public partial class TrayIconViewModel : ObservableObject
{
    private readonly IBluetoothConnectionService _bluetoothService;
    private readonly IStartupService _startupService;
    private readonly DispatcherTimer _blinkTimer;
    private bool _blinkState;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(TooltipText))]
    [NotifyPropertyChangedFor(nameof(IconBackground))]
    [NotifyPropertyChangedFor(nameof(StatusText))]
    private bool _isConnected;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IconBackground))]
    [NotifyPropertyChangedFor(nameof(StatusText))]
    [NotifyCanExecuteChangedFor(nameof(ConnectCommand))]
    private bool _isConnecting;

    [ObservableProperty]
    private bool _isAutoStartEnabled;

    public string TooltipText => IsConnected
        ? "Sony WH-1000XM5: Polaczony"
        : IsConnecting
            ? "Sony WH-1000XM5: Laczenie..."
            : "Sony WH-1000XM5: Rozlaczony (kliknij aby polaczyc)";

    public string StatusText => IsConnected
        ? "Polaczony"
        : IsConnecting ? "Laczenie..." : "Rozlaczony";

    // Kolor tla ikony: zielony = polaczony, czerwony = niepodlaczony, migajacy = laczenie
    public Brush IconBackground
    {
        get
        {
            if (IsConnected)
                return new SolidColorBrush(Color.FromRgb(76, 175, 80));   // Zielony (#4CAF50)

            if (IsConnecting)
            {
                // Miganie: czerwony <-> pomaranczowy
                return _blinkState
                    ? new SolidColorBrush(Color.FromRgb(255, 152, 0))    // Pomaranczowy (#FF9800)
                    : new SolidColorBrush(Color.FromRgb(244, 67, 54));   // Czerwony (#F44336)
            }

            return new SolidColorBrush(Color.FromRgb(244, 67, 54));      // Czerwony (#F44336)
        }
    }

    public Brush IconForeground => Brushes.White;

    public TrayIconViewModel(
        IBluetoothConnectionService bluetoothService,
        IStartupService startupService)
    {
        _bluetoothService = bluetoothService;
        _startupService = startupService;

        // Timer do migania
        _blinkTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(400)
        };
        _blinkTimer.Tick += OnBlinkTimerTick;

        _isAutoStartEnabled = _startupService.IsAutoStartEnabled();
        _isConnected = _bluetoothService.IsConnected;

        _bluetoothService.ConnectionStatusChanged += OnConnectionStatusChanged;
        _bluetoothService.StartMonitoring();
    }

    private void OnBlinkTimerTick(object? sender, EventArgs e)
    {
        _blinkState = !_blinkState;
        OnPropertyChanged(nameof(IconBackground));
    }

    private void OnConnectionStatusChanged(object? sender, bool connected)
    {
        Application.Current?.Dispatcher.Invoke(() =>
        {
            IsConnected = connected;
            if (connected)
            {
                StopBlinking();
                IsConnecting = false;
            }
        });
    }

    private void StartBlinking()
    {
        _blinkState = false;
        _blinkTimer.Start();
    }

    private void StopBlinking()
    {
        _blinkTimer.Stop();
        _blinkState = false;
        OnPropertyChanged(nameof(IconBackground));
    }

    [RelayCommand(CanExecute = nameof(CanConnect))]
    private async Task ConnectAsync()
    {
        if (IsConnected || IsConnecting) return;

        IsConnecting = true;
        StartBlinking();

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
                    System.Diagnostics.Debug.WriteLine("Device not found");
                    break;

                case ConnectionResult.ConnectionFailed:
                    System.Diagnostics.Debug.WriteLine("Connection failed");
                    break;

                case ConnectionResult.BluetoothError:
                    System.Diagnostics.Debug.WriteLine("Bluetooth error");
                    break;
            }
        }
        finally
        {
            StopBlinking();
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
        _blinkTimer.Stop();
        _bluetoothService.StopMonitoring();
        _bluetoothService.Dispose();
        Application.Current?.Shutdown();
    }
}
