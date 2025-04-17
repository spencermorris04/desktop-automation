using System;
using System.Collections.Generic;
using System.Diagnostics;
using DiagnosticsProcess = System.Diagnostics.Process;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

public static class WindowLayoutManager
{
    // ──────────────────── metadata ───────────────────────────────────────
    private enum InstancePolicy { UseExisting, AlwaysNew }

    private record Area(Rectangle Rect,
                        string Query,
                        InstancePolicy Policy,
                        bool PromptIfMultiple = false);

    private sealed class Layout
    {
        public string Name { get; }
        public IReadOnlyList<Area> Areas { get; }
        public Layout(string name, IEnumerable<Area> areas) { Name = name; Areas = areas.ToList(); }
    }

    // ──────────────────── layouts ────────────────────────────────────────
    private static readonly List<Layout> _layouts = new()
    {
        BuildCodingLayout(),
        BuildResearchLayout()
    };

    private static Layout BuildCodingLayout()
    {
        Rectangle w = Working;
        int halfW = w.Width / 2, halfH = w.Height / 2;
        return new("coding", new[]
        {
            new Area(new Rectangle(w.Left,       w.Top, halfW, w.Height),
                     "code", InstancePolicy.UseExisting, true),
            new Area(new Rectangle(w.Left+halfW, w.Top, halfW, halfH),
                     "notepad", InstancePolicy.AlwaysNew),
            new Area(new Rectangle(w.Left+halfW, w.Top+halfH, halfW, halfH),
                     "windows terminal", InstancePolicy.AlwaysNew)
        });
    }

    private static Layout BuildResearchLayout()
    {
        Rectangle w = Working;
        int halfW = w.Width / 2;
        return new("research", new[]
        {
            new Area(new Rectangle(w.Left,       w.Top, halfW, w.Height),
                     "chrome", InstancePolicy.UseExisting, true),
            new Area(new Rectangle(w.Left+halfW, w.Top, halfW, w.Height),
                     "obsidian", InstancePolicy.UseExisting, true)
        });
    }

    // ──────────────────── public API ─────────────────────────────────────
    public static IEnumerable<string> GetLayoutNames() => _layouts.Select(l => l.Name);

    public static bool ApplyLayout(string name)
    {
        var layout = _layouts.FirstOrDefault(l => l.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        if (layout == null) { Console.WriteLine($"Layout '{name}' not found."); return false; }

        // PASS 1: choose or launch windows (no moves yet)
        var selections = new List<(Area area, WindowHelper.WindowInfo? win)>();
        foreach (var area in layout.Areas)
            selections.Add((area, ResolveWindow(area)));

        // PASS 2: position after all prompts
        foreach (var (area, win) in selections)
        {
            if (win == null) { Console.WriteLine($" • Could not locate '{area.Query}'."); continue; }
            Position(win.Handle, area.Rect);
            WindowHelper.FocusWindow(win);
        }
        return true;
    }

    // ──────────────────── window resolution ──────────────────────────────
    private static WindowHelper.WindowInfo? ResolveWindow(Area area)
    {
        var existing = GetExistingWindows(area.Query);

        if (area.Policy == InstancePolicy.UseExisting && existing.Any())
        {
            if (existing.Count == 1)                       // auto‑select single
                return existing[0];

            if (area.PromptIfMultiple)                     // prompt when >1
            {
                Console.WriteLine($"\nSelect '{area.Query}' window:");
                for (int i = 0; i < existing.Count; i++)
                    Console.WriteLine($"{i + 1}. [{existing[i].ProcessName}] {existing[i].Title}");
                Console.WriteLine($"{existing.Count + 1}. Open NEW '{area.Query}' window");
                Console.Write("Choice #: ");

                if (int.TryParse(Console.ReadLine(), out int sel) &&
                    sel >= 1 && sel <= existing.Count + 1)
                {
                    WindowHelper.WindowInfo? chosen;
                    if (sel <= existing.Count)
                        chosen = existing[sel - 1];
                    else
                    {
                        if (!LaunchCommand(area.Query)) return null;
                        chosen = WaitForNewWindow(area.Query, existing);
                        if (chosen == null) return null;
                    }
                    return chosen;
                }
                Console.WriteLine("Invalid selection, skipping area.");
                return null;
            }

            return existing[0];                            // first match
        }

        // always‑new policy or none found
        if (!LaunchCommand(area.Query)) return null;
        return WaitForNewWindow(area.Query, existing);
    }

    private static List<WindowHelper.WindowInfo> GetExistingWindows(string query)
    {
        string procName =
            query.Equals("code",      StringComparison.OrdinalIgnoreCase) ? "Code" :
            query.Equals("obsidian",  StringComparison.OrdinalIgnoreCase) ? "Obsidian" : null;

        if (procName != null)
            return WindowHelper.GetOpenWindows()
                                .Where(w => w.ProcessName.Equals(procName, StringComparison.OrdinalIgnoreCase))
                                .ToList();

        return WindowHelper.SearchWindows(query, 10, 60);
    }

    private static WindowHelper.WindowInfo? WaitForNewWindow(string query, List<WindowHelper.WindowInfo> oldSet)
    {
        for (int i = 0; i < 20; i++)
        {
            Thread.Sleep(500);
            var win = GetExistingWindows(query).Except(oldSet).FirstOrDefault();
            if (win != null) return win;
        }
        Console.WriteLine($"   Timeout waiting for '{query}' window.");
        return null;
    }

    // ──────────────────── launcher map ───────────────────────────────────
    private static readonly Dictionary<string, string[]> LaunchMap = new()
    {
        { "code", new[] { "code",
                          @"%LOCALAPPDATA%\Programs\Microsoft VS Code\Code.exe",
                          @"C:\Program Files\Microsoft VS Code\Code.exe" } },

        { "obsidian", new[] { "obsidian",
                              @"%LOCALAPPDATA%\Programs\Obsidian\Obsidian.exe" } },

        { "windows terminal", new[] { "wt",  @"C:\Windows\System32\wt.exe" } },
        { "notepad", new[] { "notepad" } },
        { "chrome", new[] { "chrome" } }
    };

    private static bool LaunchCommand(string query)
    {
        if (!LaunchMap.TryGetValue(query.ToLowerInvariant(), out var cmds))
            cmds = new[] { query };

        foreach (var cmd in cmds)
        {
            string exe = Environment.ExpandEnvironmentVariables(cmd);
            try
            {
                var psi = new ProcessStartInfo { FileName = exe, UseShellExecute = true };
                if (DiagnosticsProcess.Start(psi) != null) return true;
            }
            catch { /* try next */ }
        }
        Console.WriteLine($"   Launch‑error: could not run '{query}'.");
        return false;
    }

    // ──────────────────── placement ──────────────────────────────────────
    private const int BORDER = 6;

    private static void Position(IntPtr hWnd, Rectangle target)
    {
        Native.ShowWindowAsync(hWnd, Native.SW_RESTORE);

        var r = new Rectangle(target.X - BORDER,
                              target.Y - BORDER,
                              target.Width  + 2 * BORDER,
                              target.Height + BORDER + BORDER/2);

        Native.SetWindowPos(hWnd, IntPtr.Zero, r.Left, r.Top, r.Width, r.Height,
                            Native.SWP_NOZORDER | Native.SWP_SHOWWINDOW | Native.SWP_FRAMECHANGED);
    }

    // ──────────────────── helpers / interop ──────────────────────────────
    private static Rectangle Working =>
        System.Windows.Forms.Screen.PrimaryScreen?.WorkingArea ??
        System.Windows.Forms.Screen.PrimaryScreen?.Bounds ?? new Rectangle(0, 0, 1920, 1080);

    private static class Native
    {
        public const int  SW_RESTORE       = 9;
        public const uint SWP_NOZORDER     = 0x0004;
        public const uint SWP_SHOWWINDOW   = 0x0040;
        public const uint SWP_FRAMECHANGED = 0x0020;

        [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int cmd);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr after,
                                               int X, int Y, int cx, int cy, uint flags);
    }
}
