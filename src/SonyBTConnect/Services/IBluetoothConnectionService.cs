namespace SonyBTConnect.Services;

public interface IBluetoothConnectionService : IDisposable
{
    bool IsConnected { get; }
    event EventHandler<bool>? ConnectionStatusChanged;
    Task<ConnectionResult> ConnectAsync();
    void StartMonitoring();
    void StopMonitoring();
}

public enum ConnectionResult
{
    Success,
    AlreadyConnected,
    DeviceNotFound,
    ConnectionFailed,
    BluetoothError
}
