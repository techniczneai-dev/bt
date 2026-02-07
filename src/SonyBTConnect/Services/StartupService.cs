using Microsoft.Win32;

namespace SonyBTConnect.Services;

public interface IStartupService
{
    bool IsAutoStartEnabled();
    void EnableAutoStart();
    void DisableAutoStart();
}

public class StartupService : IStartupService
{
    private const string AppName = "SonyBTConnect";
    private const string RegistryPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";

    public bool IsAutoStartEnabled()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RegistryPath, false);
            return key?.GetValue(AppName) != null;
        }
        catch
        {
            return false;
        }
    }

    public void EnableAutoStart()
    {
        try
        {
            string? appPath = Environment.ProcessPath;

            if (string.IsNullOrEmpty(appPath))
            {
                appPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            }

            // Dla single-file apps .NET
            if (appPath.EndsWith(".dll", StringComparison.OrdinalIgnoreCase))
            {
                appPath = appPath.Replace(".dll", ".exe", StringComparison.OrdinalIgnoreCase);
            }

            using var key = Registry.CurrentUser.OpenSubKey(RegistryPath, true);
            key?.SetValue(AppName, $"\"{appPath}\"", RegistryValueKind.String);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"EnableAutoStart error: {ex.Message}");
        }
    }

    public void DisableAutoStart()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RegistryPath, true);
            key?.DeleteValue(AppName, throwOnMissingValue: false);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"DisableAutoStart error: {ex.Message}");
        }
    }
}
