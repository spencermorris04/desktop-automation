using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;                 // ← NEW (for window‑layout rectangles)
using System.IO;
using System.Linq;
using System.Text;                    // For StringBuilder in Excel handler
using NAudio.CoreAudioApi;            // For Audio helper

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Desktop Automation Script");
        Console.WriteLine("=========================");

        // Start the background sampler when the application starts
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
                Console.WriteLine("6. Switch / Focus Window");   // ← NEW
                Console.WriteLine("7. Apply Window Layout");     // ← NEW
                Console.WriteLine("8. Search & Open Folder");    // ← NEW
                Console.WriteLine("0. Exit");
                Console.Write("Enter choice: ");
                string? choice = Console.ReadLine();

                switch (choice)
                {
                    case "1": HandleNotepad();         break;
                    case "2": HandleTaskManager();     break;
                    case "3": HandleAudioManager();    break;
                    case "4": HandleFileExplorer();    break;
                    case "5": HandleExcel();           break;
                    case "6": HandleWindowSwitcher();  break; // NEW
                    case "7": HandleWindowLayouts();   break; // NEW
                    case "8": HandleFolderSearch();    break; // NEW
                    case "0": Console.WriteLine("Exiting..."); return;
                    default:  Console.WriteLine("Invalid choice. Please try again."); break;
                }

                Console.WriteLine("\nPress Enter to continue...");
                Console.ReadLine();
                Console.Clear();
            }
        }
        finally
        {
            // Stop the background sampler before exiting
            Console.WriteLine("Shutting down background sampler...");
            TaskManagerHelper.StopBackgroundSampler();
            Console.WriteLine("Cleanup complete. Exiting application.");
            System.Threading.Thread.Sleep(500);
        }
    }

    // ======================================================================
    //  NEW FEATURES
    // ======================================================================

    // --- Window switcher ---------------------------------------------------
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

    // --- Layout manager ----------------------------------------------------
    static void HandleWindowLayouts()
    {
        Console.WriteLine("\nAvailable layouts:");
        var names = WindowLayoutManager.GetLayoutNames().ToList();
        for (int i = 0; i < names.Count; i++)
            Console.WriteLine($"{i + 1}. {names[i]}");

        Console.Write("Layout name or #: ");
        string? input = Console.ReadLine();
        string layout = int.TryParse(input, out int index) && index >= 1 && index <= names.Count
                        ? names[index - 1]
                        : (input ?? string.Empty);

        if (string.IsNullOrWhiteSpace(layout))
        {
            Console.WriteLine("No layout selected.");
            return;
        }

        WindowLayoutManager.ApplyLayout(layout);
    }

    // --- Folder search / open ---------------------------------------------
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

    // ======================================================================
    //  EXISTING HANDLERS (unchanged from your original file)
    // ======================================================================

    // ---------- Notepad Handler -------------------------------------------
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

    // ---------- Task Manager Handler --------------------------------------
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
                List<TaskManagerHelper.ProcessInfo> processes = TaskManagerHelper.GetCachedProcesses();
                Console.WriteLine("\n--- Top Processes (Cached Data) ---");
                Console.WriteLine($"{"ID",-8} {"Name",-30} {"CPU (%)",-10} {"RAM (MB)",-10}");
                Console.WriteLine(new string('-', 70));
                int count = 0;
                foreach (var p in processes
                        .OrderByDescending(p => p.CpuUsage)
                        .ThenByDescending(p => p.RamUsageMb)
                        .Take(20))
                {
                    string procName = p.Name.Length > 29 ? p.Name.Substring(0, 26) + "..." : p.Name;
                    Console.WriteLine($"{p.Id,-8} {procName,-30} {p.CpuUsage,-10:F2} {p.RamUsageMb,-10}");
                    count++;
                }
                if (count == 0)
                    Console.WriteLine("No process data available in cache (sampler might be starting or encountered errors).");
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

    // ---------- Audio Manager Handler -------------------------------------
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
            case "1": // List Devices
                List<AudioHelper.AudioDeviceInfo> devices = AudioHelper.GetPlaybackDevices();
                if (devices.Any())
                {
                    Console.WriteLine("\n--- Active Playback Devices ---");
                    foreach (var device in devices)
                    {
                        Console.WriteLine($"  Name: {device.Name} {(device.IsDefault ? "(Default)" : "")}");
                        Console.WriteLine($"  State: {device.State}");
                        Console.WriteLine($"  ID: {device.Id}\n");
                    }
                }
                else Console.WriteLine("No active playback devices found or error occurred.");
                break;

            case "2": // Set Default Device
                List<AudioHelper.AudioDeviceInfo> currentDevices = AudioHelper.GetPlaybackDevices();
                if (!currentDevices.Any())
                {
                    Console.WriteLine("No active playback devices found to set as default.");
                    break;
                }
                Console.WriteLine("\n--- Available Playback Devices ---");
                int devIndex = 1;
                foreach (var device in currentDevices)
                {
                    Console.WriteLine($"{devIndex++}. {device.Name} {(device.IsDefault ? "(Current Default)" : "")}");
                    Console.WriteLine($"   ID: {device.Id}");
                }
                Console.Write("Enter the ID of the device to set as default: ");
                string? deviceId = Console.ReadLine();
                if (!string.IsNullOrWhiteSpace(deviceId))
                {
                    if (currentDevices.Any(d => d.Id.Equals(deviceId, StringComparison.OrdinalIgnoreCase)))
                    {
                        bool setResult = AudioHelper.SetDefaultPlaybackDevice(deviceId);
                        Console.WriteLine(setResult ? "Default device set successfully." : "Failed to set default device.");
                    }
                    else Console.WriteLine("Invalid device ID entered.");
                }
                else Console.WriteLine("No device ID entered.");
                break;

            case "3": // Get Volume
                float? volume = AudioHelper.GetMasterVolume();
                if (volume.HasValue)
                    Console.WriteLine($"Current master volume: {volume.Value:P0}");
                else
                    Console.WriteLine("Could not retrieve master volume.");
                break;

            case "4": // Set Volume
                Console.Write("Enter desired volume level (0‑100): ");
                if (int.TryParse(Console.ReadLine(), out int volumePercent))
                {
                    float volumeScalar = Math.Clamp(volumePercent / 100.0f, 0.0f, 1.0f);
                    bool setVolResult = AudioHelper.SetMasterVolume(volumeScalar);
                    if (!setVolResult) Console.WriteLine("Failed to set master volume.");
                }
                else Console.WriteLine("Invalid volume level entered. Please enter a number between 0 and 100.");
                break;

            case "5": // Toggle Play/Pause
                AudioHelper.TogglePlayPause();
                break;

            default:
                Console.WriteLine("Invalid action.");
                break;
        }
    }

    // ---------- File Explorer (Downloads) Handler -------------------------
    static void HandleFileExplorer()
    {
        Console.WriteLine("\n--- File Explorer Helper (Downloads) ---");
        Console.WriteLine("Choose action:");
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
                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
                    filesToList = FileExplorerHelper.SearchDownloads(searchTerm);
                    if (!filesToList.Any())
                    {
                        Console.WriteLine($"No files found matching '{searchTerm}' with sufficient relevance.");
                        return;
                    }
                    Console.WriteLine($"\n--- Top {filesToList.Count} Fuzzy Matches for '{searchTerm}' (Sorted by Relevance) ---");
                }
                else
                {
                    Console.WriteLine("Search term cannot be empty.");
                    return;
                }
                break;

            default:
                Console.WriteLine("Invalid action.");
                return;
        }

        for (int i = 0; i < filesToList.Count; i++)
            Console.WriteLine($"{i + 1}. {filesToList[i].Name} (Modified: {filesToList[i].LastWriteTime})");

        Console.Write("\nEnter the number of the file to open (or 0 to cancel): ");
        if (int.TryParse(Console.ReadLine(), out int fileIndex) && fileIndex > 0 && fileIndex <= filesToList.Count)
            FileExplorerHelper.OpenFile(filesToList[fileIndex - 1].FullName);
        else if (fileIndex != 0)
            Console.WriteLine("Invalid selection.");
    }

    // ---------- Excel Handler --------------------------------------------
    static void HandleExcel()
    {
        Console.WriteLine("\n--- Excel Helper ---");
        Console.Write("Enter Excel file path (leave blank for a new file): ");
        string? filePath = Console.ReadLine(); // Can be null or empty

        Console.Write("Enter Worksheet name: ");
        string? sheetName = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(sheetName))
        {
            Console.WriteLine("Worksheet name cannot be empty.");
            return;
        }

        Console.WriteLine("Enter CSV data (Type 'ENDCSV' on a new line when done):");
        StringBuilder csvBuilder = new StringBuilder();
        string? line;
        while ((line = Console.ReadLine()) != null && !line.Equals("ENDCSV", StringComparison.OrdinalIgnoreCase))
            csvBuilder.AppendLine(line);

        string csvData = csvBuilder.ToString().Trim();
        if (string.IsNullOrWhiteSpace(csvData))
        {
            Console.WriteLine("No CSV data entered.");
            return;
        }

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

        Console.WriteLine("\nAttempting Excel operation...");
        bool success = ExcelHelper.WriteTableData(filePath, sheetName, data, saveAndClose, openAfterSave);
        Console.WriteLine(success ? "\nExcel operation completed successfully." : "\nExcel operation failed. Check previous messages for details.");
    }

    private static object[,]? ParseCsvToObjectArray(string csvData)
    {
        try
        {
            string[] lines = csvData.Split(new[] { Environment.NewLine, "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0) return new object[0, 0];

            int numCols = lines[0].Split(',').Length;
            int numRows = lines.Length;
            object[,] dataArray = new object[numRows, numCols];

            for (int r = 0; r < numRows; r++)
            {
                string[] cells = lines[r].Split(',');
                for (int c = 0; c < numCols; c++)
                    dataArray[r, c] = c < cells.Length ? cells[c].Trim() : string.Empty;
            }
            return dataArray;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error parsing CSV data: {ex.Message}");
            return null;
        }
    }
}
