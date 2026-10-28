using System.Runtime.InteropServices;

namespace BTAudioTray;

/// <summary>
/// Connects and disconnects Bluetooth audio devices through the AUDIO DRIVER, not the
/// Bluetooth stack. This is the approach used by ToothTray and is the only reliable way
/// on Windows: "it is the driver that connects the bluetooth device to Windows' audio
/// system." We reach the driver's Kernel Streaming filter from the audio endpoint and
/// send KSPROPERTY_ONESHOT_RECONNECT / _DISCONNECT (KSPROPSETID_BtAudio).
///
/// WinRT BluetoothDevice + RFCOMM nudges and BluetoothSetServiceState were both tried
/// and failed: the link dropped within ~1.5s and the A2DP audio profile never engaged.
/// The KS path connects and the endpoint goes ACTIVE within ~2s.
/// </summary>
internal static class BluetoothManager
{
    /// <summary>A Bluetooth audio device, surfaced via its render audio endpoint.</summary>
    public record BtDevice(string EndpointId, string Name, bool IsConnected);

    private static readonly Guid KSPROPSETID_BtAudio = new("7FA06C40-B8F6-4C7E-8556-E8C33A12E54D");
    private const uint KSPROPERTY_ONESHOT_RECONNECT = 0;
    private const uint KSPROPERTY_ONESHOT_DISCONNECT = 1;
    private const uint KSPROPERTY_TYPE_GET = 0x00000001;
    private const uint KSPROPERTY_TYPE_BASICSUPPORT = 0x00000200;
    private const int CLSCTX_ALL = 23;

    /// <summary>
    /// All Bluetooth audio render devices (active or unplugged). A device is "connected"
    /// when its endpoint is ACTIVE. We identify Bluetooth endpoints by their interface
    /// being a Bluetooth audio enumerator rather than a sound card.
    /// </summary>
    public static List<BtDevice> GetBluetoothAudioDevices()
    {
        var result = new List<BtDevice>();
        var enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
        try
        {
            enumerator.EnumAudioEndpoints(EDataFlow.eRender,
                DEVICE_STATE_ACTIVE | DEVICE_STATE_UNPLUGGED, out var collection);
            collection.GetCount(out int count);
            for (int i = 0; i < count; i++)
            {
                collection.Item(i, out var device);
                try
                {
                    if (!IsBluetoothEndpoint(device)) continue;
                    device.GetId(out string id);
                    device.GetState(out int state);
                    string name = GetFriendlyName(device);
                    result.Add(new BtDevice(id, name, state == DEVICE_STATE_ACTIVE));
                }
                finally { Marshal.ReleaseComObject(device); }
            }
            Marshal.ReleaseComObject(collection);
        }
        finally { Marshal.ReleaseComObject(enumerator); }

        // Collapse duplicates: AirPods expose a "Headphones" (A2DP) and a "Headset" (HFP)
        // endpoint with near-identical names. Prefer the connected one, then the first.
        return result
            .GroupBy(d => CleanName(d.Name), StringComparer.OrdinalIgnoreCase)
            .Select(g => g.OrderByDescending(d => d.IsConnected).First() with { Name = CleanName(g.First().Name) })
            .OrderBy(d => d.Name)
            .ToList();
    }

    /// <summary>
    /// True if the endpoint is a Bluetooth audio device. The authoritative test: its KS
    /// filter supports the KSPROPSETID_BtAudio property set. Sound cards and virtual
    /// devices (Voicemeeter, HD Audio, etc.) return ERROR_NOT_FOUND; only Bluetooth audio
    /// drivers respond. This is far more reliable than sniffing device-path strings.
    /// </summary>
    private static bool IsBluetoothEndpoint(IMMDevice device)
    {
        IKsControl? ks = null;
        try
        {
            ks = GetKsControl(device);
            if (ks == null) return false;
            var prop = new KSPROPERTY
            {
                Set = KSPROPSETID_BtAudio,
                Id = KSPROPERTY_ONESHOT_RECONNECT,
                Flags = KSPROPERTY_TYPE_BASICSUPPORT,
            };
            IntPtr buf = Marshal.AllocHGlobal(64);
            try
            {
                int hr = ks.KsProperty(ref prop, (uint)Marshal.SizeOf<KSPROPERTY>(), buf, 64, out _);
                return hr >= 0;
            }
            finally { Marshal.FreeHGlobal(buf); }
        }
        catch { return false; }
        finally { if (ks != null) Marshal.ReleaseComObject(ks); }
    }

    /// <summary>True if the endpoint is currently ACTIVE (i.e. connected). Cheap: no
    /// topology walk, just a state read on the one endpoint.</summary>
    public static bool IsEndpointActive(string endpointId)
    {
        var enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
        try
        {
            if (enumerator.GetDevice(endpointId, out var dev) < 0 || dev == null) return false;
            try { dev.GetState(out int state); return state == DEVICE_STATE_ACTIVE; }
            finally { Marshal.ReleaseComObject(dev); }
        }
        finally { Marshal.ReleaseComObject(enumerator); }
    }

    /// <summary>Connect the device behind this endpoint via the audio driver.</summary>
    public static bool Connect(string endpointId) =>
        SendOneShot(endpointId, KSPROPERTY_ONESHOT_RECONNECT);

    /// <summary>
    /// Fully disconnect the device behind this endpoint.
    ///
    /// The KS one-shot disconnect only stops the audio profiles. Some earphones keep
    /// the underlying Bluetooth link (ACL) open afterwards because other services
    /// (media controls / AVRCP, vendor channels) are still using it, so Windows keeps
    /// showing them as "Connected" and they can't pair with another device. So:
    ///   1. send the KS disconnect to every audio endpoint of this device
    ///      (render + capture, A2DP + Hands-Free), not just the one in the list;
    ///   2. then close the link itself with IOCTL_BTH_DISCONNECT_DEVICE, which is what
    ///      the Settings "Disconnect" button effectively does.
    /// </summary>
    public static bool Disconnect(string endpointId)
    {
        var (containerId, name) = GetEndpointIdentity(endpointId);

        // 1. Audio profiles, on every endpoint belonging to the same physical device.
        var endpointIds = containerId is Guid c ? GetActiveEndpointsInContainer(c) : new List<string>();
        if (!endpointIds.Contains(endpointId, StringComparer.OrdinalIgnoreCase))
            endpointIds.Add(endpointId);

        bool any = false;
        foreach (var id in endpointIds)
            any |= SendOneShot(id, KSPROPERTY_ONESHOT_DISCONNECT);

        // 2. The Bluetooth link itself.
        var addresses = FindBluetoothAddresses(containerId, name);
        if (addresses.Count > 0)
        {
            Thread.Sleep(500); // let the audio driver finish tearing down its channels
            foreach (var address in addresses)
                any |= DisconnectLink(address);
        }
        return any;
    }

    /// <summary>Container ID (shared by all endpoints/interfaces of one physical device)
    /// and friendly name of an audio endpoint.</summary>
    private static (Guid? ContainerId, string Name) GetEndpointIdentity(string endpointId)
    {
        var enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
        try
        {
            if (enumerator.GetDevice(endpointId, out var dev) < 0 || dev == null) return (null, "");
            try { return (GetContainerId(dev), GetFriendlyName(dev)); }
            finally { Marshal.ReleaseComObject(dev); }
        }
        catch { return (null, ""); }
        finally { Marshal.ReleaseComObject(enumerator); }
    }

    /// <summary>All ACTIVE render and capture endpoints with the given container ID.</summary>
    private static List<string> GetActiveEndpointsInContainer(Guid containerId)
    {
        var ids = new List<string>();
        var enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
        try
        {
            enumerator.EnumAudioEndpoints(EDataFlow.eAll, DEVICE_STATE_ACTIVE, out var collection);
            collection.GetCount(out int count);
            for (int i = 0; i < count; i++)
            {
                collection.Item(i, out var device);
                try
                {
                    if (GetContainerId(device) == containerId)
                    {
                        device.GetId(out string id);
                        ids.Add(id);
                    }
                }
                finally { Marshal.ReleaseComObject(device); }
            }
            Marshal.ReleaseComObject(collection);
        }
        catch { /* best effort: caller still disconnects the original endpoint */ }
        finally { Marshal.ReleaseComObject(enumerator); }
        return ids;
    }

    private static Guid? GetContainerId(IMMDevice device)
    {
        IPropertyStore? store = null;
        try
        {
            device.OpenPropertyStore(STGM_READ, out store);
            var key = PKEY_Device_ContainerId;
            store.GetValue(ref key, out var pv);
            try
            {
                // VT_CLSID: the PROPVARIANT holds a pointer to the GUID.
                if (pv.vt != VT_CLSID || pv.pwszVal == IntPtr.Zero) return null;
                return Marshal.PtrToStructure<Guid>(pv.pwszVal);
            }
            finally { PropVariantClear(ref pv); }
        }
        catch { return null; }
        finally { if (store != null) Marshal.ReleaseComObject(store); }
    }

    /// <summary>
    /// Bluetooth addresses of the paired device(s) behind this endpoint: matched by
    /// container ID, or by name if the endpoint has no container ID. Checks both
    /// Classic and LE entries, since dual-mode earbuds can appear as both.
    /// </summary>
    private static List<ulong> FindBluetoothAddresses(Guid? containerId, string endpointName)
    {
        const string PropContainer = "System.Devices.Aep.ContainerId";
        const string PropAddress = "System.Devices.Aep.DeviceAddress";
        var result = new List<ulong>();
        string cleanName = CleanName(endpointName);

        string[] selectors =
        {
            Windows.Devices.Bluetooth.BluetoothDevice.GetDeviceSelectorFromPairingState(true),
            Windows.Devices.Bluetooth.BluetoothLEDevice.GetDeviceSelectorFromPairingState(true),
        };

        foreach (var selector in selectors)
        {
            Windows.Devices.Enumeration.DeviceInformationCollection infos;
            try
            {
                infos = Windows.Devices.Enumeration.DeviceInformation
                    .FindAllAsync(selector, new[] { PropContainer, PropAddress },
                        Windows.Devices.Enumeration.DeviceInformationKind.AssociationEndpoint)
                    .AsTask().GetAwaiter().GetResult();
            }
            catch { continue; }

            foreach (var di in infos)
            {
                bool match = containerId is Guid c
                    ? di.Properties.TryGetValue(PropContainer, out var cv) && cv is Guid g && g == c
                    : cleanName.Length > 0 &&
                      string.Equals(CleanName(di.Name), cleanName, StringComparison.OrdinalIgnoreCase);
                if (!match) continue;

                if (di.Properties.TryGetValue(PropAddress, out var av) && av is string s &&
                    TryParseBtAddress(s, out ulong address) && !result.Contains(address))
                    result.Add(address);
            }
        }
        return result;
    }

    // "aa:bb:cc:dd:ee:ff" (or with '-') -> 0xAABBCCDDEEFF
    private static bool TryParseBtAddress(string s, out ulong address) =>
        ulong.TryParse(s.Replace(":", "").Replace("-", ""),
            System.Globalization.NumberStyles.HexNumber, null, out address) && address != 0;

    /// <summary>Close the Bluetooth link to this address on every local radio.</summary>
    private static bool DisconnectLink(ulong address)
    {
        bool ok = false;
        var p = new BLUETOOTH_FIND_RADIO_PARAMS { dwSize = (uint)Marshal.SizeOf<BLUETOOTH_FIND_RADIO_PARAMS>() };
        IntPtr find = BluetoothFindFirstRadio(ref p, out IntPtr radio);
        if (find == IntPtr.Zero) return false;
        try
        {
            do
            {
                try
                {
                    ok |= DeviceIoControl(radio, IOCTL_BTH_DISCONNECT_DEVICE,
                        ref address, sizeof(ulong), IntPtr.Zero, 0, out _, IntPtr.Zero);
                }
                finally { CloseHandle(radio); }
            }
            while (BluetoothFindNextRadio(find, out radio));
        }
        finally { BluetoothFindRadioClose(find); }
        return ok;
    }

    private static bool SendOneShot(string endpointId, uint propertyId)
    {
        var enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
        try
        {
            int hr = enumerator.GetDevice(endpointId, out var endpoint);
            if (hr < 0 || endpoint == null) return false;
            try
            {
                var ks = GetKsControl(endpoint);
                if (ks == null) return false;
                try
                {
                    var prop = new KSPROPERTY
                    {
                        Set = KSPROPSETID_BtAudio,
                        Id = propertyId,
                        Flags = KSPROPERTY_TYPE_GET,
                    };
                    int r = ks.KsProperty(ref prop, (uint)Marshal.SizeOf<KSPROPERTY>(),
                        IntPtr.Zero, 0, out _);
                    return r >= 0;
                }
                finally { Marshal.ReleaseComObject(ks); }
            }
            finally { Marshal.ReleaseComObject(endpoint); }
        }
        finally { Marshal.ReleaseComObject(enumerator); }
    }

    /// <summary>
    /// Walk from an audio endpoint through its device topology to the connected adapter
    /// part, then activate IKsControl on that part's topology object. That control talks
    /// to the Bluetooth audio driver filter.
    /// </summary>
    private static IKsControl? GetKsControl(IMMDevice endpoint)
    {
        Guid iidTopo = typeof(IDeviceTopology).GUID;
        if (endpoint.Activate(ref iidTopo, CLSCTX_ALL, IntPtr.Zero, out object topoObj) < 0)
            return null;
        var topology = (IDeviceTopology)topoObj;
        try
        {
            topology.GetConnectorCount(out uint cc);
            for (uint i = 0; i < cc; i++)
            {
                topology.GetConnector(i, out IConnector conn);
                try
                {
                    conn.IsConnected(out bool connected);
                    if (!connected) continue;
                    conn.GetConnectedTo(out IConnector other);
                    try
                    {
                        var otherPart = (IPart)other;
                        otherPart.GetTopologyObject(out IDeviceTopology otherTopo);
                        try
                        {
                            otherTopo.GetDeviceId(out string devId);
                            var ks = ActivateKsControl(devId);
                            if (ks != null) return ks;
                        }
                        finally { Marshal.ReleaseComObject(otherTopo); }
                    }
                    finally { Marshal.ReleaseComObject(other); }
                }
                finally { Marshal.ReleaseComObject(conn); }
            }
        }
        finally { Marshal.ReleaseComObject(topology); }
        return null;
    }

    private static IKsControl? ActivateKsControl(string deviceId)
    {
        var enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
        try
        {
            if (enumerator.GetDevice(deviceId, out var dev) < 0 || dev == null) return null;
            try
            {
                Guid iid = typeof(IKsControl).GUID;
                if (dev.Activate(ref iid, CLSCTX_ALL, IntPtr.Zero, out object ksObj) < 0 || ksObj == null)
                    return null;
                return (IKsControl)ksObj;
            }
            finally { Marshal.ReleaseComObject(dev); }
        }
        finally { Marshal.ReleaseComObject(enumerator); }
    }

    // Strip the role prefix Windows adds: "Headphones (X)" / "Headset (X)" -> "X".
    private static string CleanName(string friendlyName)
    {
        int open = friendlyName.IndexOf('(');
        int close = friendlyName.LastIndexOf(')');
        string inner = (open >= 0 && close > open)
            ? friendlyName.Substring(open + 1, close - open - 1)
            : friendlyName;
        // Drop trailing " - Find My" and similar Apple suffixes for display.
        int dash = inner.IndexOf(" - ", StringComparison.Ordinal);
        if (dash > 0) inner = inner.Substring(0, dash);
        return inner.Trim();
    }

    private static string GetFriendlyName(IMMDevice device)
    {
        device.OpenPropertyStore(STGM_READ, out var store);
        try
        {
            var key = PKEY_Device_FriendlyName;
            store.GetValue(ref key, out var pv);
            try { return Marshal.PtrToStringUni(pv.pwszVal) ?? "(unknown)"; }
            finally { PropVariantClear(ref pv); }
        }
        finally { Marshal.ReleaseComObject(store); }
    }

    // ===================================================================
    //  COM / KS interop
    // ===================================================================
    private const int DEVICE_STATE_ACTIVE = 0x1;
    private const int DEVICE_STATE_UNPLUGGED = 0x8;
    private const int STGM_READ = 0x0;

    private enum EDataFlow { eRender, eCapture, eAll }

    // PKEY_Device_FriendlyName {a45c254e-...},14 ; PKEY_Device_InstanceId is {78c34fc8-...},256
    private static PROPERTYKEY PKEY_Device_FriendlyName = new()
    { fmtid = new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"), pid = 14 };

    [StructLayout(LayoutKind.Sequential)]
    public struct KSPROPERTY { public Guid Set; public uint Id; public uint Flags; }

    [StructLayout(LayoutKind.Sequential)]
    private struct PROPERTYKEY { public Guid fmtid; public int pid; }

    [StructLayout(LayoutKind.Explicit)]
    private struct PROPVARIANT { [FieldOffset(0)] public short vt; [FieldOffset(8)] public IntPtr pwszVal; }

    [DllImport("ole32.dll")] private static extern int PropVariantClear(ref PROPVARIANT pvar);

    // PKEY_Device_ContainerId {8c7ed206-3f8a-4827-b3ab-ae9e1faefc6c},2 (VT_CLSID)
    private static PROPERTYKEY PKEY_Device_ContainerId = new()
    { fmtid = new Guid("8c7ed206-3f8a-4827-b3ab-ae9e1faefc6c"), pid = 2 };
    private const short VT_CLSID = 72;

    // ===================================================================
    // Bluetooth link interop (bthioctl.h / bluetoothapis.h)
    // ===================================================================

    // BTH_CTL(BTH_IOCTL_BASE + 0x03) = CTL_CODE(FILE_DEVICE_BLUETOOTH, 0x03, METHOD_BUFFERED, FILE_ANY_ACCESS)
    private const uint IOCTL_BTH_DISCONNECT_DEVICE = 0x0041000C;

    [StructLayout(LayoutKind.Sequential)]
    private struct BLUETOOTH_FIND_RADIO_PARAMS { public uint dwSize; }

    [DllImport("BluetoothApis.dll", SetLastError = true)]
    private static extern IntPtr BluetoothFindFirstRadio(ref BLUETOOTH_FIND_RADIO_PARAMS p, out IntPtr phRadio);

    [DllImport("BluetoothApis.dll", SetLastError = true)]
    private static extern bool BluetoothFindNextRadio(IntPtr hFind, out IntPtr phRadio);

    [DllImport("BluetoothApis.dll", SetLastError = true)]
    private static extern bool BluetoothFindRadioClose(IntPtr hFind);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool DeviceIoControl(IntPtr hDevice, uint ioControlCode,
        ref ulong inBuffer, uint inBufferSize, IntPtr outBuffer, uint outBufferSize,
        out uint bytesReturned, IntPtr overlapped);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr handle);

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] private class MMDeviceEnumerator { }

    [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IMMDeviceEnumerator
    {
        int EnumAudioEndpoints(EDataFlow flow, int mask, out IMMDeviceCollection col);
        int GetDefaultAudioEndpoint(EDataFlow flow, int role, out IMMDevice dev);
        [PreserveSig] int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string id, out IMMDevice dev);
        int RegisterEndpointNotificationCallback(IntPtr c);
        int UnregisterEndpointNotificationCallback(IntPtr c);
    }

    [ComImport, Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IMMDeviceCollection { int GetCount(out int c); int Item(int i, out IMMDevice d); }

    [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IMMDevice
    {
        [PreserveSig] int Activate(ref Guid iid, int clsCtx, IntPtr p, [MarshalAs(UnmanagedType.IUnknown)] out object o);
        int OpenPropertyStore(int access, out IPropertyStore store);
        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);
        int GetState(out int state);
    }

    [ComImport, Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPropertyStore
    {
        int GetCount(out int c); int GetAt(int i, out PROPERTYKEY k);
        int GetValue(ref PROPERTYKEY k, out PROPVARIANT v);
        int SetValue(ref PROPERTYKEY k, ref PROPVARIANT v); int Commit();
    }

    [ComImport, Guid("2A07407E-6497-4A18-9787-32F79BD0D98F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IDeviceTopology
    {
        int GetConnectorCount(out uint count);
        int GetConnector(uint index, out IConnector connector);
        int GetSubunitCount(out uint count);
        int GetSubunit(uint index, out IntPtr subunit);
        int GetPartById(uint id, out IntPtr part);
        int GetDeviceId([MarshalAs(UnmanagedType.LPWStr)] out string id);
        int GetSignalPath(IntPtr from, IntPtr to, bool rejectMixed, out IntPtr parts);
    }

    [ComImport, Guid("9c2c4058-23f5-41de-877a-df3af236a09e"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IConnector
    {
        int GetType(out int type);
        int GetDataFlow(out int flow);
        int ConnectTo(IConnector other);
        int Disconnect();
        int IsConnected(out bool connected);
        int GetConnectedTo(out IConnector other);
        int GetConnectorIdConnectedTo([MarshalAs(UnmanagedType.LPWStr)] out string id, out IntPtr x);
        int GetDeviceIdConnectedTo([MarshalAs(UnmanagedType.LPWStr)] out string id);
    }

    [ComImport, Guid("AE2DE0E4-5BCA-4F2D-AA46-5D13F8FDB3A9"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPart
    {
        int GetName([MarshalAs(UnmanagedType.LPWStr)] out string name);
        int GetLocalId(out uint id);
        int GetGlobalId([MarshalAs(UnmanagedType.LPWStr)] out string id);
        int GetPartType(out int t);
        int GetSubType(out Guid g);
        int GetControlInterfaceCount(out uint c);
        int GetControlInterface(uint i, out IntPtr ci);
        int EnumPartsIncoming(out IntPtr parts);
        int EnumPartsOutgoing(out IntPtr parts);
        int GetTopologyObject(out IDeviceTopology topo);
    }

    [ComImport, Guid("28F54685-06FD-11D2-B27A-00A0C9223196"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IKsControl
    {
        [PreserveSig] int KsProperty(ref KSPROPERTY prop, uint propLen, IntPtr data, uint dataLen, out uint bytesReturned);
        [PreserveSig] int KsMethod(IntPtr m, uint mLen, IntPtr data, uint dataLen, out uint bytesReturned);
        [PreserveSig] int KsEvent(IntPtr e, uint eLen, IntPtr data, uint dataLen, out uint bytesReturned);
    }
}
