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
        // Skrypt PowerShell z UI Automation
        string script = @"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms

$deviceName = 'WH-1000XM5'

# Otwórz ustawienia Bluetooth
Start-Process 'ms-settings:bluetooth'
Start-Sleep -Seconds 2

# Znajdź okno ustawień
$root = [System.Windows.Automation.AutomationElement]::RootElement
$condition = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::NameProperty,
    'Settings'
)
$settingsWindow = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)

if (-not $settingsWindow) {
    $condition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ClassNameProperty,
        'ApplicationFrameWindow'
    )
    $windows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $condition)
    foreach ($win in $windows) {
        if ($win.Current.Name -like '*Settings*' -or $win.Current.Name -like '*Ustawienia*') {
            $settingsWindow = $win
            break
        }
    }
}

if ($settingsWindow) {
    Write-Host 'Found Settings window'

    # Przygotuj mouse_event
    $signature = @'
[DllImport(""user32.dll"")]
public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
'@
    $mouse = Add-Type -MemberDefinition $signature -Name 'MouseClick' -Namespace 'Win32API' -PassThru -ErrorAction SilentlyContinue
    if (-not $mouse) { $mouse = [Win32API.MouseClick] }

    # Szukaj elementu z nazwą urządzenia
    Start-Sleep -Seconds 1

    $nameCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty,
        $deviceName
    )

    $deviceElement = $settingsWindow.FindFirst(
        [System.Windows.Automation.TreeScope]::Descendants,
        $nameCondition
    )

    if ($deviceElement) {
        Write-Host 'Found device element'

        # Kliknij na urządzenie
        $invokePattern = $deviceElement.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
        if ($invokePattern) {
            $invokePattern.Invoke()
            Start-Sleep -Milliseconds 500
        } else {
            # Spróbuj kliknąć przez pozycję
            $rect = $deviceElement.Current.BoundingRectangle
            $x = [int]($rect.X + $rect.Width / 2)
            $y = [int]($rect.Y + $rect.Height / 2)

            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
            Start-Sleep -Milliseconds 100
            $mouse::mouse_event(0x0002, 0, 0, 0, 0) # MOUSEEVENTF_LEFTDOWN
            $mouse::mouse_event(0x0004, 0, 0, 0, 0) # MOUSEEVENTF_LEFTUP
            Start-Sleep -Milliseconds 500
        }

        # Szukaj przycisku Connect/Połącz
        Start-Sleep -Milliseconds 500

        # Szukaj przycisku Connect - przez typ Button i nazwe zawierajaca connect/polacz
        $btnCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Button
        )
        $allButtons = $settingsWindow.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            $btnCondition
        )

        foreach ($btn in $allButtons) {
            $btnName = $btn.Current.Name
            if ($btnName -and ($btnName -like '*onnect*' -or $btnName -like '*łącz*' -or $btnName -like '*olacz*')) {
                # Pomijamy Disconnect
                if ($btnName -notlike '*isconnect*' -and $btnName -notlike '*ozłącz*' -and $btnName -notlike '*ozlacz*') {
                    Write-Host ""Found Connect button: $btnName""

                    # Sprobuj Invoke
                    try {
                        $invokePattern = $btn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                        if ($invokePattern) {
                            $invokePattern.Invoke()
                            Write-Host 'Clicked Connect via Invoke'
                            break
                        }
                    } catch {
                        # Kliknij myszka
                        $rect = $btn.Current.BoundingRectangle
                        if ($rect.Width -gt 0) {
                            $x = [int]($rect.X + $rect.Width / 2)
                            $y = [int]($rect.Y + $rect.Height / 2)
                            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
                            Start-Sleep -Milliseconds 100
                            $mouse::mouse_event(0x0002, 0, 0, 0, 0)
                            $mouse::mouse_event(0x0004, 0, 0, 0, 0)
                            Write-Host 'Clicked Connect via mouse'
                            break
                        }
                    }
                }
            }
        }
    } else {
        Write-Host 'Device element not found, trying alternative method'

        # Alternatywna metoda - szukaj wszystkich elementów i znajdź po tekście
        $allCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::IsEnabledProperty,
            $true
        )
        $allElements = $settingsWindow.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            $allCondition
        )

        foreach ($elem in $allElements) {
            if ($elem.Current.Name -like ""*$deviceName*"") {
                Write-Host ""Found element with device name: $($elem.Current.Name)""

                $rect = $elem.Current.BoundingRectangle
                if ($rect.Width -gt 0) {
                    $x = [int]($rect.X + $rect.Width / 2)
                    $y = [int]($rect.Y + $rect.Height / 2)

                    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
                    Start-Sleep -Milliseconds 100
                    $mouse::mouse_event(0x0002, 0, 0, 0, 0)
                    $mouse::mouse_event(0x0004, 0, 0, 0, 0)

                    Start-Sleep -Seconds 1

                    # Szukaj Connect - przez wszystkie przyciski
                    $btnCond = New-Object System.Windows.Automation.PropertyCondition(
                        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                        [System.Windows.Automation.ControlType]::Button
                    )
                    $btns = $settingsWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, $btnCond)

                    foreach ($b in $btns) {
                        $bName = $b.Current.Name
                        if ($bName -and ($bName -like '*onnect*' -or $bName -like '*łącz*' -or $bName -like '*olacz*')) {
                            if ($bName -notlike '*isconnect*' -and $bName -notlike '*ozłącz*') {
                                $rect2 = $b.Current.BoundingRectangle
                                if ($rect2.Width -gt 0) {
                                    $x2 = [int]($rect2.X + $rect2.Width / 2)
                                    $y2 = [int]($rect2.Y + $rect2.Height / 2)
                                    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x2, $y2)
                                    Start-Sleep -Milliseconds 100
                                    $mouse::mouse_event(0x0002, 0, 0, 0, 0)
                                    $mouse::mouse_event(0x0004, 0, 0, 0, 0)
                                    Write-Host ""Clicked Connect button: $bName""
                                    break
                                }
                            }
                        }
                    }
                    break
                }
            }
        }
    }

    # Zamknij ustawienia po chwili
    Start-Sleep -Seconds 3
    Stop-Process -Name 'SystemSettings' -ErrorAction SilentlyContinue
} else {
    Write-Host 'Settings window not found'
}
";

        var psi = new ProcessStartInfo
        {
            FileName = "powershell",
            Arguments = $"-ExecutionPolicy Bypass -Command \"{script.Replace("\"", "\\\"")}\"",
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
