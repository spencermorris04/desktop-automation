using System;
using System.Runtime.InteropServices;

public static class ComHelper
{
    // P/Invoke declarations for COM functions
    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(
        ref Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    [DllImport("ole32.dll")]
    private static extern int CLSIDFromProgID(
        [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
        out Guid pclsid);

    // Tries to get a running COM object. Returns null if not found or error.
    public static object? TryGetActiveObject(string progId)
    {
        Guid clsid;
        try
        {
            // Get CLSID from ProgID
            int hr = CLSIDFromProgID(progId, out clsid);
            if (hr < 0) // Check for HRESULT error
            {
                // Log specific error if needed using Marshal.ThrowExceptionForHR(hr);
                Console.WriteLine($"Warning: Could not get CLSID for ProgID '{progId}'. HR={hr:X}");
                return null;
            }

            // Get the active object using the CLSID
            GetActiveObject(ref clsid, IntPtr.Zero, out object comObject);
            return comObject;
        }
        catch (COMException ex) when ((uint)ex.ErrorCode == 0x800401E3) // MK_E_UNAVAILABLE
        {
            // Object not running, this is expected if the app isn't open
            // Console.WriteLine($"Debug: COM object with ProgID '{progId}' is not running (MK_E_UNAVAILABLE).");
            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Error calling GetActiveObject for '{progId}': {ex.Message}");
            return null;
        }
    }
}
