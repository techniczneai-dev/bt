using System.Diagnostics;
using System.IO;
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
        try
        {
            Debug.WriteLine("Starting connection attempt...");

            if (CheckIfConnected())
            {
                Debug.WriteLine("Already connected");
                UpdateConnectionStatus(true);
                return ConnectionResult.AlreadyConnected;
            }

            // Retry up to 3 times
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                Debug.WriteLine($"Connection attempt {attempt}/3...");

                await ConnectViaSettingsAsync();

                // Wait for stable connection (max 12 seconds per attempt)
                for (int i = 0; i < 12; i++)
                {
                    await Task.Delay(1000);
                    if (CheckIfConnected())
                    {
                        // Verify connection is stable - check again after 2s
                        await Task.Delay(2000);
                        if (CheckIfConnected())
                        {
                            Debug.WriteLine($"Connected on attempt {attempt}");
                            UpdateConnectionStatus(true);
                            return ConnectionResult.Success;
                        }
                        Debug.WriteLine("Connection dropped, retrying...");
                        break;
                    }
                }
            }

            return ConnectionResult.ConnectionFailed;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Connect error: {ex.Message}");
            return ConnectionResult.BluetoothError;
        }
    }

    private async Task ConnectViaSettingsAsync()
    {
        string scriptPath = Path.Combine(Path.GetTempPath(), "bt_connect.ps1");

        string script = @"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$deviceName = 'WH-1000XM5'

# Open Bluetooth settings
Start-Process 'ms-settings:bluetooth'
Start-Sleep -Seconds 4

# Find Settings window by name (PL: Ustawienia, EN: Settings)
$root = [System.Windows.Automation.AutomationElement]::RootElement
$settingsWindow = $null

$allWindows = $root.FindAll(
    [System.Windows.Automation.TreeScope]::Children,
    [System.Windows.Automation.Condition]::TrueCondition
)
foreach ($win in $allWindows) {
    $name = $win.Current.Name
    if ($name -like '*stawienia*' -or $name -like '*etting*') {
        $settingsWindow = $win
        Write-Host ""Found Settings: '$name'""
        break
    }
}

if (-not $settingsWindow) {
    Write-Host 'Settings window not found'
    exit 1
}

Start-Sleep -Seconds 1

# Get ALL elements in the window
$allElements = $settingsWindow.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition
)

Write-Host ""Total elements: $($allElements.Count)""

# Find device and then the first Connect button after it
$foundDevice = $false
$clicked = $false

foreach ($el in $allElements) {
    $name = $el.Current.Name
    $type = $el.Current.ControlType.ProgrammaticName

    if (-not $name -or $name.Length -eq 0) { continue }

    # Look for our device (appears as Group or Text element)
    if ($name -like ""*$deviceName*"") {
        $foundDevice = $true
        Write-Host ""Found device: [$type] '$name'""
        continue
    }

    # After finding device, look for Connect/Polacz button
    if ($foundDevice -and $type -like '*Button*') {
        # Match: Connect (EN), Polacz (PL), Verbinden (DE)
        $autoId = $el.Current.AutomationId
        $isConnect = ($name -eq 'Connect') -or
                     ($name -match '(?i)^po\S*cz$') -or
                     ($name -match '(?i)^verbind') -or
                     ($autoId -match '_Button$' -and $name -notlike '*EntityItemButton*' -and $name -notlike '*Poka*')

        if ($isConnect) {
            Write-Host ""Found Connect button: '$name' AutoId='$autoId' - invoking""

            try {
                $invokePattern = $el.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                $invokePattern.Invoke()
                Write-Host ""Clicked!""
                $clicked = $true
            } catch {
                Write-Host ""InvokePattern failed: $($_.Exception.Message)""
            }
            break
        }
    }
}

if (-not $foundDevice) {
    Write-Host ""Device '$deviceName' not found""
}
if ($foundDevice -and -not $clicked) {
    Write-Host ""Device found but Connect button not clicked""
}

# Close Settings after connection initiated
Start-Sleep -Seconds 5
Stop-Process -Name 'SystemSettings' -ErrorAction SilentlyContinue
";

        await File.WriteAllTextAsync(scriptPath, script, System.Text.Encoding.UTF8);

        var psi = new ProcessStartInfo
        {
            FileName = "powershell",
            Arguments = $"-ExecutionPolicy Bypass -File \"{scriptPath}\"",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var process = Process.Start(psi);
        if (process != null)
        {
            string output = await process.StandardOutput.ReadToEndAsync();
            string error = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            Debug.WriteLine($"UI Automation output:\n{output}");
            if (!string.IsNullOrEmpty(error))
                Debug.WriteLine($"UI Automation error:\n{error}");
        }

        try { File.Delete(scriptPath); } catch { }
    }

    private bool CheckIfConnected()
    {
        return CheckIfConnectedViaAudioEndpoints();
    }

    private bool CheckIfConnectedViaAudioEndpoints()
    {
        IMMDeviceEnumerator? enumerator = null;
        IMMDeviceCollection? devices = null;

        try
        {
            enumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorClass();
            int hr = enumerator.EnumAudioEndpoints(EDataFlow.eRender, DeviceState.DEVICE_STATE_ACTIVE, out devices);

            if (hr != 0 || devices == null) return false;

            devices.GetCount(out uint count);

            for (uint i = 0; i < count; i++)
            {
                devices.Item(i, out var device);
                if (device == null) continue;

                try
                {
                    string? deviceName = GetDeviceFriendlyName(device);
                    if (deviceName != null && deviceName.Contains(DEVICE_NAME, StringComparison.OrdinalIgnoreCase))
                        return true;
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
            if (devices != null) Marshal.ReleaseComObject(devices);
            if (enumerator != null) Marshal.ReleaseComObject(enumerator);
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

    private void CheckConnectionStatus(object? state)
    {
        try
        {
            bool nowConnected = CheckIfConnected();
            UpdateConnectionStatus(nowConnected);
        }
        catch { }
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
        _statusCheckTimer.Change(0, 2000);
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
