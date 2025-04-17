// -----------------------------------------------------------------------------
//  Program.cs – DesktopAutomator demo with Chrome‑tab bridge
// -----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using NAudio.CoreAudioApi;
using System.Net.Http;

class Program
{
    private static readonly HttpClient _http = new(); // for Chrome bridge

    static void Main(string[] args)
    {
        Console.WriteLine("Desktop Automation Script");
        Console.WriteLine("=========================");

        ChromeBridgeServer.Start();          // ← NEW bridge
        TaskManagerHelper.StartBackgroundSampler();

        try
        {
            while (true)
            {
                Console.WriteLine("\nChoose an application or utility to use:");
                Console.WriteLine("1. Notepad");
                Console.WriteLine("2. Task Manager");
                Console.WriteLine("3. Audio Manager");
                Console.WriteLine("4. File Explorer (Downloads)");
                Console.WriteLine("5. Excel");
                Console.WriteLine("6. Switch / Focus Window");
                Console.WriteLine("7. Apply Window Layout");
                Console.WriteLine("8. Search & Open Folder");
                Console.WriteLine("9. Chrome Tabs");          // ← NEW
                Console.WriteLine("0. Exit");
                Console.Write("Enter choice: ");
                string? choice = Console.ReadLine();

                switch (choice)
                {
                    case "1": HandleNotepad();                             break;
                    case "2": HandleTaskManager();                         break;
                    case "3": HandleAudioManager();                        break;
                    case "4": HandleFileExplorer();                        break;
                    case "5": HandleExcel();                                break;
                    case "6": HandleWindowSwitcher();                       break;
                    case "7": HandleWindowLayouts();                        break;
                    case "8": HandleFolderSearch();                         break;
                    case "9": HandleChromeTabs().GetAwaiter().GetResult();  break;
                    case "0": Console.WriteLine("Exiting…");  return;
                    default:  Console.WriteLine("Invalid choice.");         break;
                }

                Console.WriteLine("\nPress Enter to continue…");
                Console.ReadLine();
                Console.Clear();
            }
        }
        finally
        {
            ChromeBridgeServer.Stop();                 // ← NEW
            TaskManagerHelper.StopBackgroundSampler();
            Console.WriteLine("Cleanup complete. Bye.");
        }
    }

    // =========================================================================
    //  OPTION 9 – Chrome tab workflow
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
                Console.WriteLine($"{i + 1}. {(tabs[i].active ? "*" : " ")} {tabs[i].title}");

            Console.Write("\nSelect [#] to activate, o <url> to open, c <#> to close, or q to quit → ");
            var input = Console.ReadLine()?.Trim() ?? "";
            if (input.Equals("q", StringComparison.OrdinalIgnoreCase))
                return;

            // Activate existing tab
            if (int.TryParse(input, out int idx))
            {
                if (idx >= 1 && idx <= tabs.Count)
                {
                    var t = tabs[idx - 1];
                    await ChromeBridgeServer.RequestAsync("activate", new { t.windowId, t.tabId });
                    WindowHelper.FocusChromeWindow(t.windowId);
                }
                else
                {
                    Console.WriteLine("Invalid tab number.");
                }
                continue;
            }

            // Open new tab
            if (input.StartsWith("o ", StringComparison.OrdinalIgnoreCase))
            {
                var url = input[2..].Trim();
                if (!url.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                    url = "https://" + url;

                var meta  = await ChromeBridgeServer.RequestAsync("open", new { url });
                int winId = meta.GetProperty("windowId").GetInt32();
                int tabId = meta.GetProperty("tabId").GetInt32();

                // focus newly opened tab’s window
                await ChromeBridgeServer.RequestAsync("activate", new { windowId = winId, tabId });
                WindowHelper.FocusChromeWindow(winId);
                continue;
            }

            // Close a tab
            if (input.StartsWith("c ", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(input[2..].Trim(), out int cidx)
                    && cidx >= 1 && cidx <= tabs.Count)
                {
                    await ChromeBridgeServer.RequestAsync("close", new { tabId = tabs[cidx - 1].tabId });
                }
                else
                {
                    Console.WriteLine("Invalid tab number.");
                }
                continue;
            }

            Console.WriteLine("Unknown command.");
        }
    }

    private record TabInfo(int windowId, int tabId, string title, string url, bool active);

    // =========================================================================
    //  OPTION 6 – Window switcher
    // =========================================================================
    static void HandleWindowSwitcher()
    {
        Console.Write("\nEnter window / app name: ");
        string? query = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(query))
        {
            Console.WriteLine("Query cannot be empty.");
            return;
        }

        var hits = WindowHelper.SearchWindows(query, 10, 50);
        if (hits.Count == 0)
        {
            Console.WriteLine("No matching windows found.");
            return;
        }

        if (hits.Count == 1)
        {
            WindowHelper.FocusWindow(hits[0]);
            Console.WriteLine($"Focused: {hits[0].Title}");
            return;
        }

        Console.WriteLine("\nMultiple matches:");
        for (int i = 0; i < hits.Count; i++)
            Console.WriteLine($"{i + 1}. [{hits[i].ProcessName}] {hits[i].Title}");

        Console.Write("Select #: ");
        if (int.TryParse(Console.ReadLine(), out int sel) && sel >= 1 && sel <= hits.Count)
            WindowHelper.FocusWindow(hits[sel - 1]);
        else
            Console.WriteLine("Invalid selection.");
    }

    // =========================================================================
    //  OPTION 7 – Window layouts
    // =========================================================================
    static void HandleWindowLayouts()
    {
        Console.WriteLine("\nAvailable layouts:");
        var names = WindowLayoutManager.GetLayoutNames().ToList();
        for (int i = 0; i < names.Count; i++)
            Console.WriteLine($"{i + 1}. {names[i]}");

        Console.Write("Layout name or #: ");
        string? input = Console.ReadLine();
        string layout = int.TryParse(input, out int idx) && idx >= 1 && idx <= names.Count
                        ? names[idx - 1]
                        : (input ?? string.Empty);

        if (string.IsNullOrWhiteSpace(layout))
        {
            Console.WriteLine("No layout selected.");
            return;
        }

        WindowLayoutManager.ApplyLayout(layout);
    }

    // =========================================================================
    //  OPTION 8 – Folder search & open
    // =========================================================================
    static void HandleFolderSearch()
    {
        Console.Write("\nFolder name to open: ");
        string? query = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(query))
        {
            Console.WriteLine("Query cannot be empty.");
            return;
        }

        var matches = FolderSearchHelper.SearchFolders(query, 8, 60);
        if (!matches.Any())
        {
            Console.WriteLine("No folders found.");
            return;
        }

        if (matches.Count == 1)
        {
            FolderSearchHelper.OpenFolder(matches[0]);
            return;
        }

        Console.WriteLine("\nMatches:");
        for (int i = 0; i < matches.Count; i++)
            Console.WriteLine($"{i + 1}. {matches[i]}");

        Console.Write("Select #: ");
        if (int.TryParse(Console.ReadLine(), out int sel) && sel >= 1 && sel <= matches.Count)
            FolderSearchHelper.OpenFolder(matches[sel - 1]);
        else
            Console.WriteLine("Invalid selection.");
    }

    // =========================================================================
    //  OPTION 1 – Notepad Automation
    // =========================================================================
    static void HandleNotepad()
    {
        Console.WriteLine("\n--- Notepad Automation ---");
        NotepadAutomator.OpenNotepad();
        Console.WriteLine("Choose action:");
        Console.WriteLine("1. Append text to Notepad");
        Console.WriteLine("2. Replace text in Notepad");
        Console.WriteLine("3. Read text from Notepad");
        Console.Write("Enter choice: ");
        string? action = Console.ReadLine();
        switch (action)
        {
            case "1":
                Console.Write("Enter text to append: ");
                string? textToAppend = Console.ReadLine();
                if (!string.IsNullOrEmpty(textToAppend))
                {
                    bool success = NotepadAutomator.WriteText(textToAppend, true);
                    Console.WriteLine(success ? "Append attempted." : "Append failed.");
                }
                else Console.WriteLine("No text entered.");
                break;

            case "2":
                Console.Write("Enter text to replace: ");
                string? textToReplace = Console.ReadLine();
                if (!string.IsNullOrEmpty(textToReplace))
                {
                    bool success = NotepadAutomator.WriteText(textToReplace, false);
                    Console.WriteLine(success ? "Write attempted." : "Write failed.");
                }
                else Console.WriteLine("No text entered.");
                break;

            case "3":
                string? readText = NotepadAutomator.ReadText();
                if (readText != null)
                    Console.WriteLine("\n--- Text Read ---\n" + readText + "\n--- End Text ---");
                else
                    Console.WriteLine("Read failed or Notepad not found.");
                break;

            default:
                Console.WriteLine("Invalid action.");
                break;
        }
    }

    // =========================================================================
    //  OPTION 2 – Task Manager helper
    // =========================================================================
    static void HandleTaskManager()
    {
        Console.WriteLine("\n--- Task Manager Helper ---");
        Console.WriteLine("Choose action:");
        Console.WriteLine("1. List top resource‑consuming processes");
        Console.WriteLine("2. Terminate a process by ID");
        Console.Write("Enter choice: ");
        string? action = Console.ReadLine();

        switch (action)
        {
            case "1":
                var processes = TaskManagerHelper.GetCachedProcesses();
                Console.WriteLine("\n--- Top Processes (Cached Data) ---");
                Console.WriteLine($"{"ID",-8} {"Name",-30} {"CPU (%)",-10} {"RAM (MB)",-10}");
                Console.WriteLine(new string('-', 70));
                int count = 0;
                foreach (var p in processes
                        .OrderByDescending(p => p.CpuUsage)
                        .ThenByDescending(p => p.RamUsageMb)
                        .Take(20))
                {
                    string procName = p.Name.Length > 29 ? p.Name.Substring(0, 26) + "…" : p.Name;
                    Console.WriteLine($"{p.Id,-8} {procName,-30} {p.CpuUsage,-10:F2} {p.RamUsageMb,-10}");
                    count++;
                }
                if (count == 0)
                    Console.WriteLine("No process data available in cache.");
                break;

            case "2":
                Console.Write("Enter the Process ID (PID) to terminate: ");
                if (int.TryParse(Console.ReadLine(), out int pid))
                    TaskManagerHelper.TerminateProcess(pid);
                else
                    Console.WriteLine("Invalid Process ID.");
                break;

            default:
                Console.WriteLine("Invalid action.");
                break;
        }
    }

    // =========================================================================
    //  OPTION 3 – Audio manager
    // =========================================================================
    static void HandleAudioManager()
    {
        Console.WriteLine("\n--- Audio Manager ---");
        Console.WriteLine("Choose action:");
        Console.WriteLine("1. List playback devices");
        Console.WriteLine("2. Set default device");
        Console.WriteLine("3. Get master volume");
        Console.WriteLine("4. Set master volume");
        Console.WriteLine("5. Toggle Play/Pause media key");
        Console.Write("Enter choice: ");
        string? action = Console.ReadLine();

        switch (action)
        {
            case "1":
                var devices = AudioHelper.GetPlaybackDevices();
                if (devices.Any())
                {
                    Console.WriteLine("\n--- Active Playback Devices ---");
                    foreach (var d in devices)
                        Console.WriteLine($"  {d.Name} {(d.IsDefault?"(Default)":"")}  ID: {d.Id}");
                }
                else Console.WriteLine("No devices found.");
                break;

            case "2":
                var cur = AudioHelper.GetPlaybackDevices();
                if (!cur.Any())
                {
                    Console.WriteLine("No playback devices to set.");
                    break;
                }
                Console.WriteLine("\n--- Available Devices ---");
                int di = 1;
                foreach (var d in cur)
                    Console.WriteLine($"{di++}. {d.Name} {(d.IsDefault?"(Current Default)":"")}  ID: {d.Id}");
                Console.Write("Enter device ID to set default: ");
                var devId = Console.ReadLine();
                if (!string.IsNullOrWhiteSpace(devId) && cur.Any(d=>d.Id==devId))
                    Console.WriteLine(AudioHelper.SetDefaultPlaybackDevice(devId)
                        ? "Default set." : "Failed to set default.");
                else Console.WriteLine("Invalid ID.");
                break;

            case "3":
                var vol = AudioHelper.GetMasterVolume();
                Console.WriteLine(vol.HasValue
                    ? $"Current volume: {vol.Value:P0}"
                    : "Could not retrieve volume.");
                break;

            case "4":
                Console.Write("Enter volume (0‑100): ");
                if (int.TryParse(Console.ReadLine(), out int vp))
                {
                    float scalar = Math.Clamp(vp/100f, 0f,1f);
                    Console.WriteLine(AudioHelper.SetMasterVolume(scalar)
                        ? "Volume set." : "Failed to set volume.");
                }
                else Console.WriteLine("Invalid input.");
                break;

            case "5":
                AudioHelper.TogglePlayPause();
                Console.WriteLine("Toggled Play/Pause.");
                break;

            default:
                Console.WriteLine("Invalid action.");
                break;
        }
    }

    // =========================================================================
    //  OPTION 4 – File Explorer (Downloads)
    // =========================================================================
    static void HandleFileExplorer()
    {
        Console.WriteLine("\n--- File Explorer Helper (Downloads) ---");
        Console.WriteLine("1. List recent files");
        Console.WriteLine("2. Search files by name (fuzzy)");
        Console.Write("Enter choice: ");
        string? action = Console.ReadLine();

        List<FileInfo> filesToList = new List<FileInfo>();
        switch (action)
        {
            case "1":
                filesToList = FileExplorerHelper.GetRecentDownloads();
                if (!filesToList.Any())
                {
                    Console.WriteLine("No recent files found.");
                    return;
                }
                Console.WriteLine("\n--- Top 10 Recent Files ---");
                break;

            case "2":
                Console.Write("Enter search term: ");
                string? searchTerm = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(searchTerm))
                {
                    Console.WriteLine("Search term cannot be empty.");
                    return;
                }
                filesToList = FileExplorerHelper.SearchDownloads(searchTerm);
                if (!filesToList.Any())
                {
                    Console.WriteLine($"No files found matching '{searchTerm}'.");
                    return;
                }
                Console.WriteLine($"\n--- Fuzzy Matches for '{searchTerm}' ---");
                break;

            default:
                Console.WriteLine("Invalid action.");
                return;
        }

        for (int i = 0; i < filesToList.Count; i++)
            Console.WriteLine($"{i + 1}. {filesToList[i].Name} (Modified: {filesToList[i].LastWriteTime})");

        Console.Write("\nEnter the number of the file to open (or 0 to cancel): ");
        if (int.TryParse(Console.ReadLine(), out int fileIndex) 
            && fileIndex > 0 
            && fileIndex <= filesToList.Count)
        {
            FileExplorerHelper.OpenFile(filesToList[fileIndex - 1].FullName);
        }
        else if (fileIndex != 0)
        {
            Console.WriteLine("Invalid selection.");
        }
    }

    // =========================================================================
    //  OPTION 5 – Excel helper
    // =========================================================================
    static void HandleExcel()
    {
        Console.WriteLine("\n--- Excel Helper ---");
        Console.Write("Enter Excel file path (leave blank to create new): ");
        string? filePath = Console.ReadLine();

        Console.Write("Enter Worksheet name: ");
        string? sheetName = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(sheetName))
        {
            Console.WriteLine("Worksheet name cannot be empty.");
            return;
        }

        Console.WriteLine("Enter CSV data (type 'ENDCSV' on a new line to finish):");
        var csvBuilder = new StringBuilder();
        string? line;
        while ((line = Console.ReadLine()) != null 
            && !line.Equals("ENDCSV", StringComparison.OrdinalIgnoreCase))
        {
            csvBuilder.AppendLine(line);
        }

        string csvData = csvBuilder.ToString().Trim();
        if (string.IsNullOrWhiteSpace(csvData))
        {
            Console.WriteLine("No CSV data entered.");
            return;
        }

        // Parse CSV into object[,] for Excel
        object[,]? data = ParseCsvToObjectArray(csvData);
        if (data == null)
        {
            Console.WriteLine("Aborting Excel operation due to CSV parse failure.");
            return;
        }

        Console.Write("Save and close the file? (y/n, default y): ");
        bool saveAndClose = !(Console.ReadLine()?.Trim().Equals("n", StringComparison.OrdinalIgnoreCase) ?? false);

        bool openAfterSave = false;
        if (saveAndClose)
        {
            Console.Write("Open the file after saving? (y/n, default n): ");
            openAfterSave = Console.ReadLine()?.Trim().Equals("y", StringComparison.OrdinalIgnoreCase) ?? false;
        }

        Console.WriteLine("\nAttempting Excel operation…");
        bool success = ExcelHelper.WriteTableData(
            filePath, sheetName, data, saveAndClose, openAfterSave);
        Console.WriteLine(success 
            ? "Excel operation completed successfully." 
            : "Excel operation failed.");
    }

    // =========================================================================
    //  Shared CSV helper
    // =========================================================================
    private static object[,]? ParseCsvToObjectArray(string csvData)
    {
        try
        {
            var lines = csvData.Split(new[]{Environment.NewLine,"\n"}, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0) return new object[0,0];
            int cols = lines[0].Split(',').Length, rows = lines.Length;
            var arr = new object[rows,cols];
            for (int r=0; r<rows; r++)
            {
                var cells = lines[r].Split(',');
                for (int c=0; c<cols; c++)
                    arr[r,c] = c<cells.Length ? cells[c].Trim() : string.Empty;
            }
            return arr;
        }
        catch
        {
            return null;
        }
    }
}
