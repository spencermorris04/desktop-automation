using System;
using System.Runtime.InteropServices;

// Helper classes and interfaces for setting the default audio device
// Based on undocumented COM interfaces used by Windows itself.

internal enum Role
{
    Console = 0,       // Games, system sounds, voice commands
    Multimedia = 1,    // Music, movies, narration
    Communications = 2 // Voice communications (VoIP)
}

[ComImport, Guid("F8679F50-850A-41CF-9C72-430F290290C8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IPolicyConfig
{
    [PreserveSig] int GetMixFormat(string pszDeviceName, IntPtr ppFormat);
    [PreserveSig] int GetDeviceFormat(string pszDeviceName, bool bDefault, IntPtr ppFormat);
    [PreserveSig] int ResetDeviceFormat(string pszDeviceName);
    [PreserveSig] int SetDeviceFormat(string pszDeviceName, IntPtr pEndpointFormat, IntPtr MixFormat);
    [PreserveSig] int GetProcessingPeriod(string pszDeviceName, bool bDefault, IntPtr pmftDefaultPeriod, IntPtr pmftMinimumPeriod);
    [PreserveSig] int SetProcessingPeriod(string pszDeviceName, IntPtr pmftPeriod);
    [PreserveSig] int GetShareMode(string pszDeviceName, IntPtr pMode);
    [PreserveSig] int SetShareMode(string pszDeviceName, IntPtr mode);
    [PreserveSig] int GetPropertyValue(string pszDeviceName, IntPtr key, IntPtr pv);
    [PreserveSig] int SetPropertyValue(string pszDeviceName, IntPtr key, IntPtr pv);
    [PreserveSig] int SetDefaultEndpoint(string pszDeviceName, Role role); // The method we need
    [PreserveSig] int SetEndpointVisibility(string pszDeviceName, bool bVisible);
}

[ComImport, Guid("568b9108-44bf-40b4-9006-86afe5b5a680"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)] // Vista/Win7 version
internal interface IPolicyConfigVista
{
    [PreserveSig] int GetMixFormat(string pszDeviceName, IntPtr ppFormat);
    [PreserveSig] int GetDeviceFormat(string pszDeviceName, bool bDefault, IntPtr ppFormat);
    [PreserveSig] int ResetDeviceFormat(string pszDeviceName);
    [PreserveSig] int SetDeviceFormat(string pszDeviceName, IntPtr pEndpointFormat, IntPtr MixFormat);
    [PreserveSig] int GetProcessingPeriod(string pszDeviceName, bool bDefault, IntPtr pmftDefaultPeriod, IntPtr pmftMinimumPeriod);
    [PreserveSig] int SetProcessingPeriod(string pszDeviceName, IntPtr pmftPeriod);
    [PreserveSig] int GetShareMode(string pszDeviceName, IntPtr pMode);
    [PreserveSig] int SetShareMode(string pszDeviceName, IntPtr mode);
    [PreserveSig] int GetPropertyValue(string pszDeviceName, IntPtr key, IntPtr pv);
    [PreserveSig] int SetPropertyValue(string pszDeviceName, IntPtr key, IntPtr pv);
    [PreserveSig] int SetDefaultEndpoint(string pszDeviceName, Role role); // The method we need
    [PreserveSig] int SetEndpointVisibility(string pszDeviceName, bool bVisible);
}


// Helper class to instantiate the PolicyConfig client
internal class PolicyConfigClient
{
    private readonly IPolicyConfig _policyConfig;

    public PolicyConfigClient()
    {
        // Instantiate the PolicyConfigClient COM object
        // This might require different approaches based on OS version,
        // but this often works for Vista+.
        var policyConfigClientType = Type.GetTypeFromCLSID(new Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9"));
        var policyConfigClientInstance = Activator.CreateInstance(policyConfigClientType!);

        // Try casting to the appropriate interface (IPolicyConfig or IPolicyConfigVista)
        // Modern systems likely use IPolicyConfig
        _policyConfig = (IPolicyConfig)policyConfigClientInstance!;

        // Fallback or specific check for Vista/Win7 if needed:
        // if (Environment.OSVersion.Version.Major == 6 && Environment.OSVersion.Version.Minor <= 1) {
        //     _policyConfig = (IPolicyConfigVista)policyConfigClientInstance;
        // } else {
        //     _policyConfig = (IPolicyConfig)policyConfigClientInstance;
        // }
    }

    public void SetDefaultEndpoint(string deviceId, Role role)
    {
        Marshal.ThrowExceptionForHR(_policyConfig.SetDefaultEndpoint(deviceId, role));
    }
}
