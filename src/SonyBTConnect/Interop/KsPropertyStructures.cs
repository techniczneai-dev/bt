using System.Runtime.InteropServices;

namespace SonyBTConnect.Interop;

// KSPROPERTY structure
[StructLayout(LayoutKind.Sequential)]
internal struct KsProperty
{
    public Guid Set;
    public uint Id;
    public uint Flags;
}

// KsProperty IDs and GUIDs for Bluetooth Audio
internal static class KsPropertyIds
{
    // KSPROPSETID_BtAudio {7FA06C40-B8F6-4C7E-8556-E8C33A12E54D}
    public static readonly Guid KSPROPSETID_BtAudio = new("7FA06C40-B8F6-4C7E-8556-E8C33A12E54D");

    // Bluetooth Audio Property IDs
    public const uint KSPROPERTY_ONESHOT_RECONNECT = 0;
    public const uint KSPROPERTY_ONESHOT_DISCONNECT = 1;

    // Property Flags
    public const uint KSPROPERTY_TYPE_GET = 0x00000001;
    public const uint KSPROPERTY_TYPE_SET = 0x00000002;
    public const uint KSPROPERTY_TYPE_TOPOLOGY = 0x10000000;
}
