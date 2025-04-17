using System;
using System.Collections.Generic;
using System.Diagnostics;
using DiagnosticsProcess = System.Diagnostics.Process;
using System.Runtime.InteropServices;
using System.Text;
using System.Linq;
using FuzzySharp;

public static class WindowHelper
{
    // ─────────────────────────────────────────────────────────────────────────
    //  Public record & cache
    // ─────────────────────────────────────────────────────────────────────────
    public record WindowInfo(IntPtr Handle, string Title, string ProcessName, int ProcessId);

    private static List<WindowInfo> _cache = new();
    private static DateTime _lastScan      = DateTime.MinValue;
    private static readonly TimeSpan _ttl  = TimeSpan.FromSeconds(3);

    // ───────────────────────── enumeration ──────────────────────────────────
    public static List<WindowInfo> GetOpenWindows(bool includeInvisible = false, bool force = false)
    {
        if (!force && DateTime.Now - _lastScan < _ttl) return _cache;

        var list  = new List<WindowInfo>();
        IntPtr sh = Native.GetShellWindow();

        bool Enum(IntPtr hWnd, IntPtr _)
        {
            if (hWnd == sh) return true;
            if (!includeInvisible && !Native.IsWindowVisible(hWnd)) return true;

            int len = Native.GetWindowTextLength(hWnd);
            if (len == 0) return true;

            var sb = new StringBuilder(len + 1);
            Native.GetWindowText(hWnd, sb, sb.Capacity);

            Native.GetWindowThreadProcessId(hWnd, out uint pid);
            string proc = "unknown";
            try { proc = DiagnosticsProcess.GetProcessById((int)pid).ProcessName; } catch { }

            list.Add(new WindowInfo(hWnd, sb.ToString(), proc, (int)pid));
            return true;
        }

        Native.EnumWindows(Enum, IntPtr.Zero);
        _cache   = list;
        _lastScan = DateTime.Now;
        return list;
    }

    // ───────────────────────── fuzzy search ──────────────────────────────────
    public static List<WindowInfo> SearchWindows(string q, int top = 5, int cutoff = 60) =>
        string.IsNullOrWhiteSpace(q) ? new() :
        GetOpenWindows().Select(w => (w, score: Fuzz.WeightedRatio($"{w.ProcessName} {w.Title}", q)))
                        .Where(t => t.score >= cutoff)
                        .OrderByDescending(t => t.score)
                        .Take(top)
                        .Select(t => t.w)
                        .ToList();

    // ───────────────────────── bring to foreground (generic) ────────────────
    public static bool FocusWindow(WindowInfo w)
    {
        if (w.Handle == IntPtr.Zero) return false;

        if (Native.IsIconic(w.Handle))
            Native.ShowWindowAsync(w.Handle, Native.SW_RESTORE);
        else
            Native.ShowWindowAsync(w.Handle, Native.SW_SHOW);

        if (TryAttachSetForeground(w.Handle)) return true;

        Native.SwitchToThisWindow(w.Handle, true);
        if (Native.GetForegroundWindow() == w.Handle) return true;

        SimulateAltKeystroke();
        Native.SetForegroundWindow(w.Handle);
        return Native.GetForegroundWindow() == w.Handle;
    }

    // ───────────────────────── bring a **Chrome** window forward by windowId
    // ───────────────────────── (used by Program.cs) ─────────────────────────
    public static void FocusChromeWindow(int chromeWindowId)
    {
        IntPtr target   = IntPtr.Zero;
        IntPtr fallback = IntPtr.Zero;

        Native.EnumWindows((hWnd, _) =>
        {
            if (!Native.IsWindowVisible(hWnd)) return true;

            Native.GetWindowThreadProcessId(hWnd, out uint pid);
            try
            {
                var proc = DiagnosticsProcess.GetProcessById((int)pid);
                if (!proc.ProcessName.Equals("chrome", StringComparison.OrdinalIgnoreCase))
                    return true;

                if (fallback == IntPtr.Zero) fallback = hWnd; // first visible chrome

                // Chrome attaches an integer property "ChromeWindowId"
                IntPtr prop = Native.GetProp(hWnd, "ChromeWindowId");
                if (prop != IntPtr.Zero && prop.ToInt32() == chromeWindowId)
                {
                    target = hWnd;
                    return false; // stop enumeration
                }
            }
            catch { /* process may have exited */ }

            return true;             // continue enumeration
        }, IntPtr.Zero);

        if (target == IntPtr.Zero)   // use any Chrome window if exact id not found
            target = fallback;

        if (target != IntPtr.Zero)
        {
            var ww = new WindowInfo(target, string.Empty, "chrome", 0);
            FocusWindow(ww);
        }
    }

    // ───────────────────────── helpers ──────────────────────────────────────
    private static bool TryAttachSetForeground(IntPtr hWnd)
    {
        uint thisTid = Native.GetCurrentThreadId();
        Native.GetWindowThreadProcessId(hWnd, out uint targetTid);

        bool attached = false;
        if (thisTid != targetTid)
            attached = Native.AttachThreadInput(thisTid, targetTid, true);

        Native.AllowSetForegroundWindow(Native.ASFW_ANY);
        bool ok = Native.SetForegroundWindow(hWnd);

        if (attached)
            Native.AttachThreadInput(thisTid, targetTid, false);

        return ok && Native.GetForegroundWindow() == hWnd;
    }

    private static void SimulateAltKeystroke()
    {
        const uint KEYEVENTF_KEYUP = 0x0002;
        Native.INPUT[] input =
        {
            new() { type = 1, U = new Native.InputUnion
                    { ki = new Native.KEYBDINPUT { wVk = 0x12 } } },               // Alt down
            new() { type = 1, U = new Native.InputUnion
                    { ki = new Native.KEYBDINPUT { wVk = 0x12, dwFlags = KEYEVENTF_KEYUP } } }
        };
        Native.SendInput((uint)input.Length, input, Native.INPUT.Size);
    }

    // ───────────────────────── low‑level P/Invoke ───────────────────────────
    private static class Native
    {
        public const int  SW_SHOW    = 5;
        public const int  SW_RESTORE = 9;
        public const uint ASFW_ANY   = 0xFFFFFFFF;

        public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")] public static extern bool   EnumWindows(EnumProc f, IntPtr p);
        [DllImport("user32.dll")] public static extern IntPtr GetShellWindow();
        [DllImport("user32.dll")] public static extern bool   IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] public static extern bool   IsIconic(IntPtr h);
        [DllImport("user32.dll")] public static extern bool   ShowWindowAsync(IntPtr h, int n);
        [DllImport("user32.dll")] public static extern bool   SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern int    GetWindowTextLength(IntPtr h);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int    GetWindowText(IntPtr h, StringBuilder sb, int max);
        [DllImport("user32.dll")] public static extern uint   GetWindowThreadProcessId(IntPtr h, out uint pid);
        [DllImport("user32.dll")] public static extern bool   AttachThreadInput(uint id1, uint id2, bool attach);
        [DllImport("user32.dll")] public static extern bool   AllowSetForegroundWindow(uint pid);
        [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] public static extern void   SwitchToThisWindow(IntPtr hWnd, bool fAltTab);

        //  GetPropW to read ChromeWindowId
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr GetProp(IntPtr hWnd, string lpString);

        //  SendInput & related structs
        [DllImport("user32.dll", SetLastError = true)] public static extern uint SendInput(uint n, INPUT[] p, int cb);
        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT { public uint type; public InputUnion U; public static int Size => Marshal.SizeOf<INPUT>(); }
        [StructLayout(LayoutKind.Explicit)]
        public struct InputUnion
        {
            [FieldOffset(0)] public MOUSEINPUT    mi;
            [FieldOffset(0)] public KEYBDINPUT    ki;
            [FieldOffset(0)] public HARDWAREINPUT hi;
        }
        [StructLayout(LayoutKind.Sequential)] public struct MOUSEINPUT  { public int dx, dy; public uint mouseData, dwFlags, time; public IntPtr dwExtraInfo; }
        [StructLayout(LayoutKind.Sequential)] public struct HARDWAREINPUT { public uint uMsg; public ushort wParamL, wParamH; }
        [StructLayout(LayoutKind.Sequential)] public struct KEYBDINPUT
        {
            public ushort wVk, wScan;
            public uint   dwFlags, time;
            public IntPtr dwExtraInfo;
        }
    }
}
