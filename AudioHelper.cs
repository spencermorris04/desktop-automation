using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NAudio.CoreAudioApi; // Requires NAudio NuGet package

public static class AudioHelper
{
    // Structure to hold device info
    public record AudioDeviceInfo(string Id, string Name, string State, bool IsDefault);

    // --- Device Enumeration and Default Setting ---

    /// <summary>
    /// Gets a list of active audio playback devices.
    /// </summary>
    public static List<AudioDeviceInfo> GetPlaybackDevices()
    {
        var devices = new List<AudioDeviceInfo>();
        MMDevice? defaultDevice = null; // Declare outside try
        try
        {
            using (var enumerator = new MMDeviceEnumerator())
            {
                // FIX: Qualify Role with NAudio namespace
                defaultDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, NAudio.CoreAudioApi.Role.Multimedia);
                foreach (var device in enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
                {
                    using (device) // Ensure device COM object is released
                    {
                        devices.Add(new AudioDeviceInfo(
                            device.ID,
                            device.FriendlyName,
                            device.State.ToString(),
                            device.ID == defaultDevice?.ID // Check if it's the default
                        ));
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error enumerating audio devices: {ex.Message}");
        }
        finally
        {
             // Release default device COM object if obtained
             defaultDevice?.Dispose();
        }
        return devices;
    }

    /// <summary>
    /// Sets the default audio playback device using undocumented COM interfaces.
    /// </summary>
    /// <param name="deviceId">The ID string of the device to set as default.</param>
    /// <returns>True if successful, false otherwise.</returns>
    public static bool SetDefaultPlaybackDevice(string deviceId)
    {
        try
        {
            // Use the PolicyConfigClient helper class
            var policyConfig = new PolicyConfigClient();
            // NOTE: These calls correctly use the *internal* Role enum defined with PolicyConfigClient
            policyConfig.SetDefaultEndpoint(deviceId, Role.Multimedia);
            policyConfig.SetDefaultEndpoint(deviceId, Role.Console);
            policyConfig.SetDefaultEndpoint(deviceId, Role.Communications);
            Console.WriteLine($"Successfully set default device ID: {deviceId}");
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error setting default audio device (ID: {deviceId}): {ex.Message}");
            Console.WriteLine("Note: This operation might require administrator privileges or may fail on some systems.");
            return false;
        }
    }

    // --- Volume Control ---

    /// <summary>
    /// Gets the current master volume level scalar (0.0 to 1.0).
    /// </summary>
    public static float? GetMasterVolume()
    {
        MMDevice? defaultDevice = null;
        try
        {
            using (var enumerator = new MMDeviceEnumerator())
            {
                // FIX: Qualify Role with NAudio namespace
                defaultDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, NAudio.CoreAudioApi.Role.Multimedia);
                if (defaultDevice == null)
                {
                    Console.WriteLine("Error: No default audio playback device found.");
                    return null;
                }
                // Access volume property safely
                float? volume = defaultDevice.AudioEndpointVolume?.MasterVolumeLevelScalar;
                return volume;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting master volume: {ex.Message}");
            return null;
        }
        finally
        {
            defaultDevice?.Dispose(); // Ensure disposal
        }
    }

    /// <summary>
    /// Sets the master volume level scalar (0.0 to 1.0).
    /// </summary>
    /// <param name="level">Volume level between 0.0f and 1.0f.</param>
    public static bool SetMasterVolume(float level)
    {
        MMDevice? defaultDevice = null;
        try
        {
            using (var enumerator = new MMDeviceEnumerator())
            {
                // FIX: Qualify Role with NAudio namespace
                defaultDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, NAudio.CoreAudioApi.Role.Multimedia);
                if (defaultDevice?.AudioEndpointVolume == null)
                {
                    Console.WriteLine("Error: Could not get volume control for default device.");
                    return false;
                }

                level = Math.Clamp(level, 0.0f, 1.0f);
                defaultDevice.AudioEndpointVolume.MasterVolumeLevelScalar = level;
                Console.WriteLine($"Master volume set to: {level:P0}");
                return true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error setting master volume: {ex.Message}");
            return false;
        }
         finally
        {
            defaultDevice?.Dispose(); // Ensure disposal
        }
    }

    // --- Media Key Simulation --- (No changes needed here)

    private const ushort VK_MEDIA_PLAY_PAUSE = 0xB3;
    private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
    private const uint KEYEVENTF_KEYUP = 0x0002;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll")]
    private static extern IntPtr GetMessageExtraInfo();

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT { /* ... */ public uint type; public InputUnion U; public static int Size => Marshal.SizeOf(typeof(INPUT)); }
    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion { /* ... */ [FieldOffset(0)] public MOUSEINPUT mi; [FieldOffset(0)] public KEYBDINPUT ki; [FieldOffset(0)] public HARDWAREINPUT hi; }
    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT { /* ... */ public int dx; public int dy; public uint mouseData; public uint dwFlags; public uint time; public IntPtr dwExtraInfo; }
    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT { /* ... */ public ushort wVk; public ushort wScan; public uint dwFlags; public uint time; public IntPtr dwExtraInfo; }
    [StructLayout(LayoutKind.Sequential)]
    private struct HARDWAREINPUT { /* ... */ public uint uMsg; public ushort wParamL; public ushort wParamH; }

    public static void TogglePlayPause()
    {
        Console.WriteLine("Sending Play/Pause media key...");
        INPUT[] inputs = new INPUT[2];
        inputs[0] = new INPUT { type = 1, U = new InputUnion { ki = new KEYBDINPUT { wVk = VK_MEDIA_PLAY_PAUSE, dwFlags = KEYEVENTF_EXTENDEDKEY, dwExtraInfo = GetMessageExtraInfo() } } };
        inputs[1] = new INPUT { type = 1, U = new InputUnion { ki = new KEYBDINPUT { wVk = VK_MEDIA_PLAY_PAUSE, dwFlags = KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, dwExtraInfo = GetMessageExtraInfo() } } };
        unsafe { uint result = SendInput((uint)inputs.Length, inputs, INPUT.Size); if (result != inputs.Length) { Console.WriteLine($"Warning: SendInput failed. Code: {Marshal.GetLastWin32Error()}"); } else { Console.WriteLine("Play/Pause key sent successfully."); } }
    }
}
