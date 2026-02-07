using System.Runtime.InteropServices;

namespace SonyBTConnect.Interop;

internal static class BluetoothApis
{
    // Bluetooth API functions
    [DllImport("bthprops.cpl", SetLastError = true)]
    public static extern bool BluetoothSetServiceState(
        IntPtr hRadio,
        ref BLUETOOTH_DEVICE_INFO pbtdi,
        ref Guid pGuidService,
        uint dwServiceFlags);

    [DllImport("bthprops.cpl", SetLastError = true)]
    public static extern IntPtr BluetoothFindFirstRadio(
        ref BLUETOOTH_FIND_RADIO_PARAMS pbtfrp,
        out IntPtr phRadio);

    [DllImport("bthprops.cpl", SetLastError = true)]
    public static extern bool BluetoothFindNextRadio(
        IntPtr hFind,
        out IntPtr phRadio);

    [DllImport("bthprops.cpl", SetLastError = true)]
    public static extern bool BluetoothFindRadioClose(IntPtr hFind);

    [DllImport("bthprops.cpl", SetLastError = true)]
    public static extern IntPtr BluetoothFindFirstDevice(
        ref BLUETOOTH_DEVICE_SEARCH_PARAMS pbtsp,
        ref BLUETOOTH_DEVICE_INFO pbtdi);

    [DllImport("bthprops.cpl", SetLastError = true)]
    public static extern bool BluetoothFindNextDevice(
        IntPtr hFind,
        ref BLUETOOTH_DEVICE_INFO pbtdi);

    [DllImport("bthprops.cpl", SetLastError = true)]
    public static extern bool BluetoothFindDeviceClose(IntPtr hFind);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool CloseHandle(IntPtr hObject);

    // Service GUIDs
    public static readonly Guid AudioSinkServiceClass_UUID = new("0000110B-0000-1000-8000-00805F9B34FB");
    public static readonly Guid AudioSourceServiceClass_UUID = new("0000110A-0000-1000-8000-00805F9B34FB");
    public static readonly Guid HandsfreeServiceClass_UUID = new("0000111E-0000-1000-8000-00805F9B34FB");
    public static readonly Guid HeadsetServiceClass_UUID = new("00001108-0000-1000-8000-00805F9B34FB");
    public static readonly Guid AVRemoteControlServiceClass_UUID = new("0000110E-0000-1000-8000-00805F9B34FB");

    // Service flags
    public const uint BLUETOOTH_SERVICE_DISABLE = 0x00;
    public const uint BLUETOOTH_SERVICE_ENABLE = 0x01;
}

[StructLayout(LayoutKind.Sequential)]
internal struct BLUETOOTH_FIND_RADIO_PARAMS
{
    public uint dwSize;

    public static BLUETOOTH_FIND_RADIO_PARAMS Create()
    {
        return new BLUETOOTH_FIND_RADIO_PARAMS
        {
            dwSize = (uint)Marshal.SizeOf<BLUETOOTH_FIND_RADIO_PARAMS>()
        };
    }
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
internal struct BLUETOOTH_DEVICE_INFO
{
    public uint dwSize;
    public ulong Address;
    public uint ulClassofDevice;
    [MarshalAs(UnmanagedType.Bool)]
    public bool fConnected;
    [MarshalAs(UnmanagedType.Bool)]
    public bool fRemembered;
    [MarshalAs(UnmanagedType.Bool)]
    public bool fAuthenticated;
    public SYSTEMTIME stLastSeen;
    public SYSTEMTIME stLastUsed;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 248)]
    public string szName;

    public static BLUETOOTH_DEVICE_INFO Create()
    {
        return new BLUETOOTH_DEVICE_INFO
        {
            dwSize = (uint)Marshal.SizeOf<BLUETOOTH_DEVICE_INFO>()
        };
    }
}

[StructLayout(LayoutKind.Sequential)]
internal struct BLUETOOTH_DEVICE_SEARCH_PARAMS
{
    public uint dwSize;
    [MarshalAs(UnmanagedType.Bool)]
    public bool fReturnAuthenticated;
    [MarshalAs(UnmanagedType.Bool)]
    public bool fReturnRemembered;
    [MarshalAs(UnmanagedType.Bool)]
    public bool fReturnUnknown;
    [MarshalAs(UnmanagedType.Bool)]
    public bool fReturnConnected;
    [MarshalAs(UnmanagedType.Bool)]
    public bool fIssueInquiry;
    public byte cTimeoutMultiplier;
    public IntPtr hRadio;

    public static BLUETOOTH_DEVICE_SEARCH_PARAMS Create(IntPtr hRadio)
    {
        return new BLUETOOTH_DEVICE_SEARCH_PARAMS
        {
            dwSize = (uint)Marshal.SizeOf<BLUETOOTH_DEVICE_SEARCH_PARAMS>(),
            fReturnAuthenticated = true,
            fReturnRemembered = true,
            fReturnUnknown = false,
            fReturnConnected = true,
            fIssueInquiry = false,
            cTimeoutMultiplier = 0,
            hRadio = hRadio
        };
    }
}

[StructLayout(LayoutKind.Sequential)]
internal struct SYSTEMTIME
{
    public ushort wYear;
    public ushort wMonth;
    public ushort wDayOfWeek;
    public ushort wDay;
    public ushort wHour;
    public ushort wMinute;
    public ushort wSecond;
    public ushort wMilliseconds;
}
