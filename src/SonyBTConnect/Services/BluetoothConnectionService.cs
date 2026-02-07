using System.Diagnostics;
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

            // Użyj UI Automation przez PowerShell
            await ConnectViaUIAutomationAsync();

            // Czekaj na połączenie (max 10 sekund)
            for (int i = 0; i < 10; i++)
            {
                await Task.Delay(1000);
                if (CheckIfConnected())
                {
                    UpdateConnectionStatus(true);
                    return ConnectionResult.Success;
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

    private async Task ConnectViaUIAutomationAsync()
    {
        // Zapisz skrypt do pliku tymczasowego (unika problemów z kodowaniem)
        string scriptPath = Path.Combine(Path.GetTempPath(), "bt_connect.ps1");

        string script = @"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms

$deviceName = 'WH-1000XM5'

# Otwórz ustawienia Bluetooth
Start-Process 'ms-settings:bluetooth'
Start-Sleep -Seconds 3

# Znajdź okno ustawień
$root = [System.Windows.Automation.AutomationElement]::RootElement
$allWindows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)

$settingsWindow = $null
foreach ($win in $allWindows) {
    $name = $win.Current.Name
    if ($name -like '*stawienia*' -or $name -like '*etting*') {
        $settingsWindow = $win
        Write-Host ""Found window: $name""
        break
    }
}

if ($settingsWindow) {
    # Przygotuj mouse click
    Add-Type @'
    using System;
    using System.Runtime.InteropServices;
    public class MouseOps {
        [DllImport(""user32.dll"")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
        public static void Click() {
            mouse_event(0x0002, 0, 0, 0, 0);
            mouse_event(0x0004, 0, 0, 0, 0);
        }
    }
'@

    Start-Sleep -Seconds 1

    # Znajdź wszystkie przyciski
    $btnCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Button
    )
    $allButtons = $settingsWindow.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        $btnCondition
    )

    Write-Host ""Found $($allButtons.Count) buttons""

    # Szukaj przycisku Connect/Polacz przy urzadzeniu
    $foundDevice = $false
    foreach ($btn in $allButtons) {
        $btnName = $btn.Current.Name

        # Sprawdz czy to nasze urzadzenie
        if ($btnName -like ""*$deviceName*"") {
            $foundDevice = $true
            Write-Host ""Found device button: $btnName""
        }

        # Szukaj przycisku Connect/Połącz
        if ($btnName -eq 'Connect' -or $btnName -eq 'Połącz' -or $btnName -eq 'Polacz') {

            Write-Host ""Found Connect button: $btnName""

            $rect = $btn.Current.BoundingRectangle
            if ($rect.Width -gt 0 -and $rect.Height -gt 0) {
                $x = [int]($rect.X + $rect.Width / 2)
                $y = [int]($rect.Y + $rect.Height / 2)

                Write-Host ""Clicking at $x, $y""
                [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
                Start-Sleep -Milliseconds 200
                [MouseOps]::Click()
                Write-Host ""Clicked!""
                break
            }
        }
    }

    # Zamknij po chwili
    Start-Sleep -Seconds 3
    Stop-Process -Name 'SystemSettings' -ErrorAction SilentlyContinue
} else {
    Write-Host 'Settings window not found'
}
";

        // Zapisz skrypt z kodowaniem UTF-8
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

            Debug.WriteLine($"UI Automation output: {output}");
            if (!string.IsNullOrEmpty(error))
                Debug.WriteLine($"UI Automation error: {error}");
        }

        // Usuń plik tymczasowy
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
