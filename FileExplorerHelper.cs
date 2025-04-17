using System.Diagnostics;
using System.IO;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using FuzzySharp;

public static class FileExplorerHelper
{
    // Gets the path to the user's Downloads folder.
    private static string? GetDownloadsFolderPath()
    {
        string? downloadsPath = null;
        try
        {
            try
            {
                downloadsPath = KnownFolders.GetPath(KnownFolder.Downloads);
                if (!string.IsNullOrEmpty(downloadsPath) && Directory.Exists(downloadsPath))
                    return downloadsPath;
            }
            catch { /* Ignore KnownFolders errors, proceed to fallback */ }

            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (!string.IsNullOrEmpty(userProfile))
            {
                downloadsPath = Path.Combine(userProfile, "Downloads");
                if (Directory.Exists(downloadsPath))
                    return downloadsPath;
            }
            Console.WriteLine($"Error: Could not find Downloads folder via KnownFolders or UserProfile methods.");
            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting Downloads folder path: {ex.Message}");
            return null;
        }
    }

    // Lists the most recent files in the Downloads folder.
    public static List<FileInfo> GetRecentDownloads(int count = 10)
    {
        string? downloadsPath = GetDownloadsFolderPath();
        if (downloadsPath == null) return new List<FileInfo>();
        try
        {
            var directory = new DirectoryInfo(downloadsPath);
            var recentFiles = directory.GetFiles()
                                       .OrderByDescending(f => f.LastWriteTime)
                                       .Take(count)
                                       .ToList();
            return recentFiles;
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException || ex is DirectoryNotFoundException)
        {
            Console.WriteLine($"Error accessing Downloads folder '{downloadsPath}': {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while listing files: {ex.Message}");
        }
        return new List<FileInfo>();
    }

    // Fuzzy search for files in the Downloads folder, returns top results by relevance.
    // --- NEW IMPLEMENTATION USING LINQ and Fuzz.Ratio ---
    public static List<FileInfo> SearchDownloads(string searchTerm, int resultLimit = 5, int scoreCutoff = 60)
    {
        string? downloadsPath = GetDownloadsFolderPath();
        if (downloadsPath == null) return new List<FileInfo>();
        if (string.IsNullOrWhiteSpace(searchTerm))
        {
            Console.WriteLine("Search term cannot be empty.");
            return new List<FileInfo>();
        }

        Console.WriteLine($"Fuzzy searching for '{searchTerm}' in: {downloadsPath}");
        var foundFiles = new List<FileInfo>();
        try
        {
            var directory = new DirectoryInfo(downloadsPath);
            var allFiles = directory.GetFiles("*", SearchOption.TopDirectoryOnly); // No need for ToList() here yet

            // --- Manual Fuzzy Matching using LINQ ---
            var scoredFiles = allFiles
                .Select(file => new // Create an intermediate object with file and score
                {
                    File = file,
                    // Calculate score using a core FuzzySharp function
                    Score = Fuzz.WeightedRatio(file.Name, searchTerm)
                })
                .Where(scored => scored.Score >= scoreCutoff) // Filter by cutoff score
                .OrderByDescending(scored => scored.Score) // Sort by score (highest first)
                .Take(resultLimit) // Take the top N results
                .Select(scored => scored.File); // Select only the FileInfo object

            foundFiles = scoredFiles.ToList(); // Convert the final result to a List

            Console.WriteLine($"Found {foundFiles.Count} matching files (Score >= {scoreCutoff}).");
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException || ex is DirectoryNotFoundException)
        {
            Console.WriteLine($"Error accessing Downloads folder '{downloadsPath}' for search: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during search: {ex.Message}");
        }
        return foundFiles;
    }


    // Opens a file using its default application.
    public static bool OpenFile(string filePath)
    {
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found at '{filePath}'.");
            return false;
        }
        Console.WriteLine($"Attempting to open file: {filePath}");
        try
        {
            ProcessStartInfo psi = new ProcessStartInfo(filePath) { UseShellExecute = true };
            System.Diagnostics.Process.Start(psi);
            Console.WriteLine("File opened successfully (or command sent to OS).");
            return true;
        }
        catch (Win32Exception ex)
        {
            Console.WriteLine($"Error opening file: {ex.Message}. (Associated application?)");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred opening file: {ex.Message}");
        }
        return false;
    }

    // KnownFolders helper class
    private static class KnownFolders
    {
        private static readonly Guid DownloadsGuid = new Guid("374DE290-123F-4565-9164-39C4925E467B");
        public static string GetPath(KnownFolder knownFolder) { return SHGetKnownFolderPath(GetGuid(knownFolder), 0); }
        private static Guid GetGuid(KnownFolder knownFolder) => knownFolder switch { KnownFolder.Downloads => DownloadsGuid, _ => throw new NotSupportedException(), };
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = false)]
        private static extern string SHGetKnownFolderPath([MarshalAs(UnmanagedType.LPStruct)] Guid rfid, uint dwFlags, IntPtr hToken = default);
    }
    public enum KnownFolder { Downloads }
}
