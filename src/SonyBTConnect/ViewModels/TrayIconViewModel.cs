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
        ? "Sony WH-1000XM5: Connected"
        : IsConnecting
            ? "Sony WH-1000XM5: Connecting..."
            : "Sony WH-1000XM5: Disconnected (click to connect)";

    public string StatusText => IsConnected
        ? "Connected"
        : IsConnecting ? "Connecting..." : "Disconnected";

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
        try
        {
            Application.Current?.Dispatcher.BeginInvoke(() =>
            {
                IsConnected = connected;
                if (connected)
                {
                    StopBlinking();
                    IsConnecting = false;
                }
            });
        }
        catch { }
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
            await _bluetoothService.ConnectAsync();
        }
        catch { }
        finally
        {
            _blinkTimer.Stop();
            _blinkState = false;
            IsConnecting = false;
            // Don't set IsConnected here - monitoring timer handles it every 1s
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
