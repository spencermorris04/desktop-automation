// -----------------------------------------------------------------------------
//  Program.cs – DesktopAutomator demo with Chrome‑tab bridge
// -----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using NAudio.CoreAudioApi;
using System.Net.Http;

class Program
{
    // reserved for future use
    private static readonly HttpClient _http = new();

    static void Main(string[] args)
    {
        Console.WriteLine("Desktop Automation Script");
        Console.WriteLine("=========================");

        ChromeBridgeServer.Start();                 // start local HTTP listener
        TaskManagerHelper.StartBackgroundSampler(); // CPU/RAM sampler

        try
        {
            while (true)
            {
                ShowMenu();
                Console.Write("\nSelect an option (number or name): ");
                var choice = Console.ReadLine()?.Trim().ToLowerInvariant() ?? "";

                if (choice == "0" || choice == "exit")
                    return;

                if (choice == "1" || choice == "notepad")
                    HandleNotepad();
                else if (choice == "2" || choice == "excel")
                    HandleExcel();
                else if (choice == "3" || choice == "tabs" || choice == "chrome")
                    HandleChromeTabs().GetAwaiter().GetResult();
                else if (choice == "4" || choice == "explorer")
                    HandleFileExplorer();
                else if (choice == "5" || choice == "search")
                    HandleFolderSearch();
                else if (choice == "6" || choice == "switch" || choice == "focus")
                    HandleWindowSwitcher();
                else if (choice == "7" || choice == "layout")
                    HandleWindowLayouts();
                else if (choice == "8" || choice == "tasks" || choice == "taskmanager")
                    HandleTaskManager();
                else if (choice == "9" || choice == "audio")
                    HandleAudioManager();
                else
                    Console.WriteLine("Invalid choice. Try again.");

                Console.WriteLine("\nPress Enter to continue…");
                Console.ReadLine();
                Console.Clear();
            }
        }
        finally
        {
            ChromeBridgeServer.Stop();
            TaskManagerHelper.StopBackgroundSampler();
            Console.WriteLine("Cleanup complete. Bye.");
        }
    }

    static void ShowMenu()
    {
        Console.WriteLine("\n=== Main Menu ===");
        Console.WriteLine(" 1. Notepad Automation      (notepad)");
        Console.WriteLine(" 2. Excel Helper            (excel)");
        Console.WriteLine(" 3. Chrome Tabs             (tabs)");
        Console.WriteLine(" 4. Open Downloads Folder   (explorer)");
        Console.WriteLine(" 5. Search & Open Folder    (search)");
        Console.WriteLine(" 6. Switch / Focus Window   (switch)");
        Console.WriteLine(" 7. Apply Window Layout     (layout)");
        Console.WriteLine(" 8. Task Manager            (tasks)");
        Console.WriteLine(" 9. Audio Mute/Unmute       (audio)");
        Console.WriteLine(" 0. Exit                    (exit)");
    }

    // =========================================================================
    //  OPTION 1 – Notepad Automation
    // =========================================================================
    static void HandleNotepad()
    {
        Console.WriteLine("\n--- Notepad Automation ---");
        NotepadAutomator.OpenNotepad();
        Console.WriteLine("1) Append text");
        Console.WriteLine("2) Replace text");
        Console.WriteLine("3) Read text");
        Console.Write("Choose action: ");

        var action = Console.ReadLine();
        switch (action)
        {
            case "1":
                Console.Write("Enter text to append: ");
                var toAppend = Console.ReadLine();
                if (!string.IsNullOrEmpty(toAppend) && 
                    NotepadAutomator.WriteText(toAppend, true))
                    Console.WriteLine("Appended.");
                else
                    Console.WriteLine("Failed or no text entered.");
                break;

            case "2":
                Console.Write("Enter text to replace: ");
                var toReplace = Console.ReadLine();
                if (!string.IsNullOrEmpty(toReplace) &&
                    NotepadAutomator.WriteText(toReplace, false))
                    Console.WriteLine("Replaced.");
                else
                    Console.WriteLine("Failed or no text entered.");
                break;

            case "3":
                var read = NotepadAutomator.ReadText();
                Console.WriteLine(read is not null
                    ? $"\n--- Content ---\n{read}\n----------------"
                    : "Read failed or Notepad not found.");
                break;

            default:
                Console.WriteLine("Invalid selection.");
                break;
        }
    }

    // =========================================================================
    //  OPTION 2 – Excel Helper (stub)
    // =========================================================================
    static void HandleExcel()
    {
        Console.WriteLine("\nExcel automation not implemented in this demo.");
    }

    // =========================================================================
    //  OPTION 3 – Chrome tab workflow
    // =========================================================================
    static async Task HandleChromeTabs()
    {
        while (true)
        {
            var tabsElem = await ChromeBridgeServer.RequestAsync("getTabs");
            var tabs = JsonSerializer
                       .Deserialize<List<TabInfo>>(tabsElem.GetRawText()) 
                       ?? new List<TabInfo>();

            Console.WriteLine("\n--- Chrome Tabs ---");
            for (int i = 0; i < tabs.Count; i++)
                Console.WriteLine($"{i+1}. {(tabs[i].active ? "*" : " ")} {tabs[i].title}");

            Console.Write("\nEnter #[activate], o <url>, c <#>, or q to quit → ");
            var cmd = Console.ReadLine()?.Trim() ?? "";
            if (cmd.Equals("q", StringComparison.OrdinalIgnoreCase))
                return;

            // Activate
            if (int.TryParse(cmd, out int idx))
            {
                if (idx < 1 || idx > tabs.Count)
                {
                    Console.WriteLine("Invalid tab number.");
                    continue;
                }
                var t = tabs[idx-1];
                var resp = await ChromeBridgeServer.RequestAsync("activate",
                             new { t.windowId, t.tabId });
                var title = resp.GetProperty("title").GetString() ?? t.title;
                var hits = WindowHelper.SearchWindows(title, 1, 60);
                if (hits.Count > 0)
                    WindowHelper.FocusWindow(hits[0]);
                else
                    Console.WriteLine("Activated but OS window not found.");
                continue;
            }

            // Open URL
            if (cmd.StartsWith("o ", StringComparison.OrdinalIgnoreCase))
            {
                var url = cmd[2..].Trim();
                if (!url.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                    url = "https://" + url;

                // create tab -> get windowId + tabId
                var meta   = await ChromeBridgeServer.RequestAsync("open", new { url });
                int winId  = meta.GetProperty("windowId").GetInt32();
                int tabId  = meta.GetProperty("tabId").GetInt32();

                // tell Chrome to activate that pair (even if opened in background)
                await ChromeBridgeServer.RequestAsync("activate", new { windowId = winId, tabId });

                // raise the native chrome window by windowId
                WindowHelper.FocusChromeWindow(winId);

                continue;
            }
            Console.WriteLine("Unknown command.");
        }
    }

    private record TabInfo(int windowId, int tabId, string title, string url, bool active);

    // =========================================================================
    //  OPTION 4 – File Explorer (Downloads)
    // =========================================================================
    static void HandleFileExplorer()
    {
        var downloads = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Downloads");
        Process.Start(new ProcessStartInfo("explorer.exe", $"\"{downloads}\"")
        {
            UseShellExecute = true
        });
        Console.WriteLine("Opened Downloads folder.");
    }

    // =========================================================================
    //  OPTION 5 – Search & Open Folder
    // =========================================================================
    static void HandleFolderSearch()
    {
        Console.Write("\nEnter folder name to open: ");
        var q = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(q))
        {
            Console.WriteLine("Query cannot be empty.");
            return;
        }
        var matches = FolderSearchHelper.SearchFolders(q, 8, 60);
        if (matches.Count == 0)
        {
            Console.WriteLine("No folders found.");
            return;
        }
        if (matches.Count == 1)
        {
            FolderSearchHelper.OpenFolder(matches[0]);
        }
        else
        {
            Console.WriteLine("Matches:");
            for (int i = 0; i < matches.Count; i++)
                Console.WriteLine($"{i+1}. {matches[i]}");
            Console.Write("Select #: ");
            if (int.TryParse(Console.ReadLine(), out int sel) &&
                sel >= 1 && sel <= matches.Count)
                FolderSearchHelper.OpenFolder(matches[sel-1]);
            else
                Console.WriteLine("Invalid selection.");
        }
    }

    // =========================================================================
    //  OPTION 6 – Switch / Focus Window
    // =========================================================================
    static void HandleWindowSwitcher()
    {
        Console.Write("\nEnter window or app name: ");
        var q = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(q))
        {
            Console.WriteLine("Query cannot be empty.");
            return;
        }
        var hits = WindowHelper.SearchWindows(q, 10, 50);
        if (hits.Count == 0)
        {
            Console.WriteLine("No matching windows found.");
        }
        else if (hits.Count == 1)
        {
            WindowHelper.FocusWindow(hits[0]);
            Console.WriteLine($"Focused: {hits[0].Title}");
        }
        else
        {
            Console.WriteLine("Multiple matches:");
            for (int i = 0; i < hits.Count; i++)
                Console.WriteLine($"{i+1}. {hits[i].Title}");
            Console.Write("Select #: ");
            if (int.TryParse(Console.ReadLine(), out int sel) &&
                sel >= 1 && sel <= hits.Count)
                WindowHelper.FocusWindow(hits[sel-1]);
            else
                Console.WriteLine("Invalid selection.");
        }
    }

    // =========================================================================
    //  OPTION 7 – Apply Window Layout
    // =========================================================================
    static void HandleWindowLayouts()
    {
        var layouts = WindowLayoutManager.GetLayoutNames().ToList();
        Console.WriteLine("\nAvailable layouts:");
        for (int i = 0; i < layouts.Count; i++)
            Console.WriteLine($"{i+1}. {layouts[i]}");
        Console.Write("Enter layout name or #: ");
        var input = Console.ReadLine();
        if (int.TryParse(input, out int idx) &&
            idx >= 1 && idx <= layouts.Count)
        {
            WindowLayoutManager.ApplyLayout(layouts[idx-1]);
        }
        else if (!string.IsNullOrWhiteSpace(input))
        {
            WindowLayoutManager.ApplyLayout(input);
        }
        else
        {
            Console.WriteLine("No layout selected.");
        }
    }

    // =========================================================================
    //  OPTION 8 – Task Manager
    // =========================================================================
    static void HandleTaskManager()
    {
        Console.WriteLine("\n--- Task Manager ---");
        Console.WriteLine("1) List top processes");
        Console.WriteLine("2) Terminate a process");
        Console.Write("Choose action: ");

        var act = Console.ReadLine();
        switch (act)
        {
            case "1":
                var procs = TaskManagerHelper.GetCachedProcesses()
                             .OrderByDescending(p => p.CpuUsage)
                             .ThenByDescending(p => p.RamUsageMb)
                             .Take(20);
                Console.WriteLine($"{"PID",-8} {"Name",-30} {"CPU",-6} {"RAM (MB)",-8}");
                foreach (var p in procs)
                    Console.WriteLine($"{p.Id,-8} {p.Name,-30} {p.CpuUsage,6:F1} {p.RamUsageMb,8}");
                break;

            case "2":
                Console.Write("Enter PID to terminate: ");
                if (int.TryParse(Console.ReadLine(), out int pid))
                    TaskManagerHelper.TerminateProcess(pid);
                else
                    Console.WriteLine("Invalid PID.");
                break;

            default:
                Console.WriteLine("Invalid action.");
                break;
        }
    }

    // =========================================================================
    //  OPTION 9 – Audio Mute/Unmute
    // =========================================================================
    static void HandleAudioManager()
    {
        try
        {
            using var devEnum = new MMDeviceEnumerator();
            var device = devEnum.GetDefaultAudioEndpoint(
                             DataFlow.Render,
                             NAudio.CoreAudioApi.Role.Multimedia);
            var muted = device.AudioEndpointVolume.Mute;
            device.AudioEndpointVolume.Mute = !muted;
            Console.WriteLine(muted ? "Audio un‑muted." : "Audio muted.");
        }
        catch (Exception e)
        {
            Console.WriteLine($"Audio control failed: {e.Message}");
        }
    }
}
