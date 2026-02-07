using System.Runtime.InteropServices;

namespace SonyBTConnect.Interop;

// MMDeviceEnumerator CLSID
[ComImport]
[Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
internal class MMDeviceEnumeratorClass { }

// IMMDeviceEnumerator Interface
[ComImport]
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IMMDeviceEnumerator
{
    [PreserveSig]
    int EnumAudioEndpoints(
        EDataFlow dataFlow,
        uint dwStateMask,
        out IMMDeviceCollection ppDevices);

    [PreserveSig]
    int GetDefaultAudioEndpoint(
        EDataFlow dataFlow,
        ERole role,
        out IMMDevice ppEndpoint);

    [PreserveSig]
    int GetDevice(
        [MarshalAs(UnmanagedType.LPWStr)] string pwstrId,
        out IMMDevice ppDevice);

    [PreserveSig]
    int RegisterEndpointNotificationCallback(IntPtr pClient);

    [PreserveSig]
    int UnregisterEndpointNotificationCallback(IntPtr pClient);
}

// IMMDeviceCollection Interface
[ComImport]
[Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IMMDeviceCollection
{
    [PreserveSig]
    int GetCount(out uint pcDevices);

    [PreserveSig]
    int Item(uint nDevice, out IMMDevice ppDevice);
}

// IMMDevice Interface
[ComImport]
[Guid("D666063F-1587-4E43-81F1-B948E807363F")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IMMDevice
{
    [PreserveSig]
    int Activate(
        ref Guid iid,
        uint dwClsCtx,
        IntPtr pActivationParams,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);

    [PreserveSig]
    int OpenPropertyStore(
        uint stgmAccess,
        out IPropertyStore ppProperties);

    [PreserveSig]
    int GetId(
        [MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);

    [PreserveSig]
    int GetState(out uint pdwState);
}

// IPropertyStore Interface
[ComImport]
[Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IPropertyStore
{
    [PreserveSig]
    int GetCount(out uint cProps);

    [PreserveSig]
    int GetAt(uint iProp, out PropertyKey pkey);

    [PreserveSig]
    int GetValue(ref PropertyKey key, out PropVariant pv);

    [PreserveSig]
    int SetValue(ref PropertyKey key, ref PropVariant propvar);

    [PreserveSig]
    int Commit();
}

// IDeviceTopology Interface
[ComImport]
[Guid("2A07407E-6497-4A18-9787-32F79BD0D98F")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IDeviceTopology
{
    [PreserveSig]
    int GetConnectorCount(out uint pCount);

    [PreserveSig]
    int GetConnector(uint nIndex, out IConnector ppConnector);

    [PreserveSig]
    int GetSubunitCount(out uint pCount);

    [PreserveSig]
    int GetSubunit(uint nIndex, [MarshalAs(UnmanagedType.IUnknown)] out object ppSubunit);

    [PreserveSig]
    int GetPartById(uint nId, [MarshalAs(UnmanagedType.IUnknown)] out object ppPart);

    [PreserveSig]
    int GetDeviceId([MarshalAs(UnmanagedType.LPWStr)] out string ppwstrDeviceId);

    [PreserveSig]
    int GetSignalPath(
        [MarshalAs(UnmanagedType.IUnknown)] object pIPartFrom,
        [MarshalAs(UnmanagedType.IUnknown)] object pIPartTo,
        bool bRejectMixedPaths,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppParts);
}

// IConnector Interface
[ComImport]
[Guid("9C2C4058-23F5-41DE-877A-DF3AF236A09E")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IConnector
{
    [PreserveSig]
    int Placeholder1(); // GetType

    [PreserveSig]
    int Placeholder2(); // GetDataFlow

    [PreserveSig]
    int ConnectTo([MarshalAs(UnmanagedType.IUnknown)] object pConnectTo);

    [PreserveSig]
    int Disconnect();

    [PreserveSig]
    int IsConnected(out bool pbConnected);

    [PreserveSig]
    int GetConnectedTo(out IConnector ppConTo);

    [PreserveSig]
    int GetConnectorIdConnectedTo([MarshalAs(UnmanagedType.LPWStr)] out string ppwstrConnectorId);

    [PreserveSig]
    int GetDeviceIdConnectedTo([MarshalAs(UnmanagedType.LPWStr)] out string ppwstrDeviceId);
}

// IPart Interface
[ComImport]
[Guid("AE2DE0E4-5BCA-4F2D-AA46-5D13F8FDB3A9")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IPart
{
    [PreserveSig]
    int GetName([MarshalAs(UnmanagedType.LPWStr)] out string ppwstrName);

    [PreserveSig]
    int GetLocalId(out uint pnId);

    [PreserveSig]
    int GetGlobalId([MarshalAs(UnmanagedType.LPWStr)] out string ppwstrGlobalId);

    [PreserveSig]
    int GetPartType(out uint pPartType);

    [PreserveSig]
    int GetSubType(out Guid pSubType);

    [PreserveSig]
    int GetControlInterfaceCount(out uint pCount);

    [PreserveSig]
    int GetControlInterface(uint nIndex, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterfaceDesc);

    [PreserveSig]
    int EnumPartsIncoming([MarshalAs(UnmanagedType.IUnknown)] out object ppParts);

    [PreserveSig]
    int EnumPartsOutgoing([MarshalAs(UnmanagedType.IUnknown)] out object ppParts);

    [PreserveSig]
    int GetTopologyObject(out IDeviceTopology ppTopology);

    [PreserveSig]
    int Activate(
        uint dwClsContext,
        ref Guid refiid,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject);

    [PreserveSig]
    int RegisterControlChangeCallback(ref Guid riid, IntPtr pNotify);

    [PreserveSig]
    int UnregisterControlChangeCallback(IntPtr pNotify);
}

// IKsControl Interface - kluczowy dla sterowania Bluetooth
[ComImport]
[Guid("28F54685-06FD-11D2-B27A-00A0C9223196")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IKsControl
{
    [PreserveSig]
    int KsProperty(
        ref KsProperty Property,
        uint PropertyLength,
        IntPtr PropertyData,
        uint DataLength,
        out uint BytesReturned);

    [PreserveSig]
    int KsMethod(
        IntPtr Method,
        uint MethodLength,
        IntPtr MethodData,
        uint DataLength,
        out uint BytesReturned);

    [PreserveSig]
    int KsEvent(
        IntPtr Event,
        uint EventLength,
        IntPtr EventData,
        uint DataLength,
        out uint BytesReturned);
}

// Enums
internal enum EDataFlow
{
    eRender = 0,
    eCapture = 1,
    eAll = 2
}

internal enum ERole
{
    eConsole = 0,
    eMultimedia = 1,
    eCommunications = 2
}

// Device States
internal static class DeviceState
{
    public const uint DEVICE_STATE_ACTIVE = 0x00000001;
    public const uint DEVICE_STATE_DISABLED = 0x00000002;
    public const uint DEVICE_STATE_NOTPRESENT = 0x00000004;
    public const uint DEVICE_STATE_UNPLUGGED = 0x00000008;
    public const uint DEVICE_STATEMASK_ALL = 0x0000000F;
}

// CLSCTX
internal static class ClsCtx
{
    public const uint CLSCTX_INPROC_SERVER = 0x1;
    public const uint CLSCTX_INPROC_HANDLER = 0x2;
    public const uint CLSCTX_LOCAL_SERVER = 0x4;
    public const uint CLSCTX_ALL = CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER | CLSCTX_LOCAL_SERVER;
}

// STGM
internal static class StorageAccessMode
{
    public const uint STGM_READ = 0x00000000;
}

// PropertyKey structure
[StructLayout(LayoutKind.Sequential)]
internal struct PropertyKey
{
    public Guid fmtid;
    public uint pid;

    public PropertyKey(Guid formatId, uint propertyId)
    {
        fmtid = formatId;
        pid = propertyId;
    }
}

// PKEY_Device_FriendlyName
internal static class PropertyKeys
{
    // {A45C254E-DF1C-4EFD-8020-67D146A850E0}, 14
    public static readonly PropertyKey PKEY_Device_FriendlyName = new(
        new Guid("A45C254E-DF1C-4EFD-8020-67D146A850E0"), 14);

    // {B3F8FA53-0004-438E-9003-51A46E139BFC}, 6
    public static readonly PropertyKey PKEY_Device_DeviceDesc = new(
        new Guid("B3F8FA53-0004-438E-9003-51A46E139BFC"), 6);
}

// PropVariant - simplified for string values
[StructLayout(LayoutKind.Explicit)]
internal struct PropVariant
{
    [FieldOffset(0)]
    public ushort vt;

    [FieldOffset(8)]
    public IntPtr pwszVal;

    public string? GetString()
    {
        if (vt == 31) // VT_LPWSTR
        {
            return Marshal.PtrToStringUni(pwszVal);
        }
        return null;
    }
}

// GUIDs
internal static class AudioGuids
{
    public static readonly Guid IID_IDeviceTopology = new("2A07407E-6497-4A18-9787-32F79BD0D98F");
    public static readonly Guid IID_IKsControl = new("28F54685-06FD-11D2-B27A-00A0C9223196");
}
