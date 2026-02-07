using System.Runtime.InteropServices;
using SonyBTConnect.Interop;

namespace SonyBTConnect.Services;

public class BluetoothConnectionService : IBluetoothConnectionService
{
    private const string DEVICE_NAME = "WH-1000XM5";
    private readonly Timer _statusCheckTimer;
    private bool _isConnected;
    private bool _isDisposed;

    public bool IsConnected => _isConnected;
    public event EventHandler<bool>? ConnectionStatusChanged;

    public BluetoothConnectionService()
    {
        _statusCheckTimer = new Timer(CheckConnectionStatus, null, Timeout.Infinite, Timeout.Infinite);
        _isConnected = CheckIfConnected();
    }

    public async Task<ConnectionResult> ConnectAsync()
    {
        return await Task.Run(() =>
        {
            try
            {
                // Sprawdź czy już połączony
                if (CheckIfConnected())
                {
                    UpdateConnectionStatus(true);
                    return ConnectionResult.AlreadyConnected;
                }

                // Znajdź urządzenie i połącz
                var result = FindAndConnectDevice();

                if (result == ConnectionResult.Success)
                {
                    // Daj czas na połączenie
                    Thread.Sleep(1000);
                    UpdateConnectionStatus(CheckIfConnected());
                }

                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Connect error: {ex.Message}");
                return ConnectionResult.BluetoothError;
            }
        });
    }

    private ConnectionResult FindAndConnectDevice()
    {
        IMMDeviceEnumerator? enumerator = null;
        IMMDeviceCollection? devices = null;

        try
        {
            enumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorClass();

            // Znajdź wszystkie audio endpoints (włącznie z niepodłączonymi)
            int hr = enumerator.EnumAudioEndpoints(
                EDataFlow.eRender,
                DeviceState.DEVICE_STATEMASK_ALL,
                out devices);

            if (hr != 0 || devices == null)
            {
                return ConnectionResult.BluetoothError;
            }

            devices.GetCount(out uint count);
            System.Diagnostics.Debug.WriteLine($"Found {count} audio endpoints");

            for (uint i = 0; i < count; i++)
            {
                devices.Item(i, out var device);
                if (device == null) continue;

                try
                {
                    string? deviceName = GetDeviceFriendlyName(device);
                    System.Diagnostics.Debug.WriteLine($"Device {i}: {deviceName}");

                    if (deviceName != null && deviceName.Contains(DEVICE_NAME, StringComparison.OrdinalIgnoreCase))
                    {
                        // Sprawdź stan urządzenia
                        device.GetState(out uint state);

                        if (state == DeviceState.DEVICE_STATE_ACTIVE)
                        {
                            System.Diagnostics.Debug.WriteLine("Device already connected");
                            return ConnectionResult.AlreadyConnected;
                        }

                        // Próba połączenia przez IKsControl
                        bool connected = TryConnectViaKsControl(device);
                        if (connected)
                        {
                            return ConnectionResult.Success;
                        }

                        // Fallback: próba przez DeviceTopology
                        connected = TryConnectViaDeviceTopology(device);
                        if (connected)
                        {
                            return ConnectionResult.Success;
                        }
                    }
                }
                finally
                {
                    if (device != null)
                        Marshal.ReleaseComObject(device);
                }
            }

            return ConnectionResult.DeviceNotFound;
        }
        finally
        {
            if (devices != null)
                Marshal.ReleaseComObject(devices);
            if (enumerator != null)
                Marshal.ReleaseComObject(enumerator);
        }
    }

    private bool TryConnectViaKsControl(IMMDevice device)
    {
        try
        {
            // Próba bezpośredniej aktywacji IKsControl
            Guid iidKsControl = AudioGuids.IID_IKsControl;
            int hr = device.Activate(
                ref iidKsControl,
                ClsCtx.CLSCTX_ALL,
                IntPtr.Zero,
                out var ksControlObj);

            if (hr == 0 && ksControlObj != null)
            {
                var ksControl = (IKsControl)ksControlObj;
                bool result = SendReconnectCommand(ksControl);
                Marshal.ReleaseComObject(ksControlObj);
                return result;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"TryConnectViaKsControl error: {ex.Message}");
        }
        return false;
    }

    private bool TryConnectViaDeviceTopology(IMMDevice device)
    {
        try
        {
            // Aktywuj IDeviceTopology
            Guid iidTopology = AudioGuids.IID_IDeviceTopology;
            int hr = device.Activate(
                ref iidTopology,
                ClsCtx.CLSCTX_ALL,
                IntPtr.Zero,
                out var topologyObj);

            if (hr != 0 || topologyObj == null)
                return false;

            var topology = (IDeviceTopology)topologyObj;

            try
            {
                topology.GetConnectorCount(out uint connectorCount);
                System.Diagnostics.Debug.WriteLine($"Device has {connectorCount} connectors");

                for (uint j = 0; j < connectorCount; j++)
                {
                    hr = topology.GetConnector(j, out var connector);
                    if (hr != 0 || connector == null) continue;

                    try
                    {
                        // Sprawdź czy connector jest podłączony
                        connector.IsConnected(out bool isConnected);

                        if (!isConnected)
                        {
                            // Pobierz connected device
                            hr = connector.GetConnectedTo(out var connectedTo);
                            if (hr == 0 && connectedTo != null)
                            {
                                // Spróbuj uzyskać IKsControl z connected device
                                if (connectedTo is IPart part)
                                {
                                    Guid iidKsControl = AudioGuids.IID_IKsControl;
                                    hr = part.Activate(ClsCtx.CLSCTX_ALL, ref iidKsControl, out var ksObj);
                                    if (hr == 0 && ksObj != null)
                                    {
                                        var ksControl = (IKsControl)ksObj;
                                        bool result = SendReconnectCommand(ksControl);
                                        Marshal.ReleaseComObject(ksObj);
                                        if (result) return true;
                                    }
                                }
                                Marshal.ReleaseComObject(connectedTo);
                            }
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(connector);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(topologyObj);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"TryConnectViaDeviceTopology error: {ex.Message}");
        }
        return false;
    }

    private bool SendReconnectCommand(IKsControl ksControl)
    {
        try
        {
            var property = new KsProperty
            {
                Set = KsPropertyIds.KSPROPSETID_BtAudio,
                Id = KsPropertyIds.KSPROPERTY_ONESHOT_RECONNECT,
                Flags = KsPropertyIds.KSPROPERTY_TYPE_SET | KsPropertyIds.KSPROPERTY_TYPE_TOPOLOGY
            };

            int hr = ksControl.KsProperty(
                ref property,
                (uint)Marshal.SizeOf<KsProperty>(),
                IntPtr.Zero,
                0,
                out _);

            System.Diagnostics.Debug.WriteLine($"KsProperty reconnect result: 0x{hr:X8}");
            return hr >= 0;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"SendReconnectCommand error: {ex.Message}");
            return false;
        }
    }

    private string? GetDeviceFriendlyName(IMMDevice device)
    {
        try
        {
            device.OpenPropertyStore(StorageAccessMode.STGM_READ, out var props);
            if (props == null) return null;

            try
            {
                var propKey = PropertyKeys.PKEY_Device_FriendlyName;
                props.GetValue(ref propKey, out var propValue);
                return propValue.GetString();
            }
            finally
            {
                Marshal.ReleaseComObject(props);
            }
        }
        catch
        {
            return null;
        }
    }

    private bool CheckIfConnected()
    {
        IMMDeviceEnumerator? enumerator = null;
        IMMDeviceCollection? devices = null;

        try
        {
            enumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorClass();

            // Pobierz tylko aktywne urządzenia
            int hr = enumerator.EnumAudioEndpoints(
                EDataFlow.eRender,
                DeviceState.DEVICE_STATE_ACTIVE,
                out devices);

            if (hr != 0 || devices == null)
                return false;

            devices.GetCount(out uint count);

            for (uint i = 0; i < count; i++)
            {
                devices.Item(i, out var device);
                if (device == null) continue;

                try
                {
                    string? deviceName = GetDeviceFriendlyName(device);
                    if (deviceName != null && deviceName.Contains(DEVICE_NAME, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(device);
                }
            }

            return false;
        }
        catch
        {
            return false;
        }
        finally
        {
            if (devices != null)
                Marshal.ReleaseComObject(devices);
            if (enumerator != null)
                Marshal.ReleaseComObject(enumerator);
        }
    }

    private void CheckConnectionStatus(object? state)
    {
        try
        {
            bool nowConnected = CheckIfConnected();
            UpdateConnectionStatus(nowConnected);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"CheckConnectionStatus error: {ex.Message}");
        }
    }

    private void UpdateConnectionStatus(bool connected)
    {
        if (connected != _isConnected)
        {
            _isConnected = connected;
            ConnectionStatusChanged?.Invoke(this, _isConnected);
        }
    }

    public void StartMonitoring()
    {
        _statusCheckTimer.Change(0, 2000); // Sprawdzaj co 2 sekundy
    }

    public void StopMonitoring()
    {
        _statusCheckTimer.Change(Timeout.Infinite, Timeout.Infinite);
    }

    public void Dispose()
    {
        if (!_isDisposed)
        {
            StopMonitoring();
            _statusCheckTimer.Dispose();
            _isDisposed = true;
        }
    }
}
