using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using DiagnosticsProcess = System.Diagnostics.Process;
using System.IO;
using System.Linq;
using System.Threading;
using FuzzySharp;

public static class FolderSearchHelper
{
    // ─── roots to scan ────────────────────────────────────────────────────
    private static readonly List<string> Roots = new()
    {
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads")
    };

    // ─── cache & sync ─────────────────────────────────────────────────────
    private static readonly List<string> _dirIndex = new();
    private static DateTime _lastRefresh = DateTime.MinValue;
    private static readonly object _lock = new();
    private static readonly TimeSpan _ttl = TimeSpan.FromMinutes(60);

    // ─── public API ───────────────────────────────────────────────────────
    public static List<string> SearchFolders(string query, int max = 8, int cutoff = 60)
    {
        if (string.IsNullOrWhiteSpace(query)) return new();
        BuildIndexIfNeeded();

        string q = query.ToLowerInvariant();

        return _dirIndex
            .Select(p => (path: p, score: Fuzz.WeightedRatio(Path.GetFileName(p).ToLowerInvariant(), q)))
            .Where(t => t.score >= cutoff)
            .OrderByDescending(t => t.score)
            .Take(max)
            .Select(t => t.path)
            .ToList();
    }

    public static bool OpenFolder(string path)
    {
        try
        {
            if (!Directory.Exists(path))
            {
                Console.WriteLine($"Folder not found: {path}");
                return false;
            }
            DiagnosticsProcess.Start("explorer.exe", $"\"{path}\"");
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error opening folder: {ex.Message}");
            return false;
        }
    }

    // ─── index builder ────────────────────────────────────────────────────
    private static void BuildIndexIfNeeded()
    {
        lock (_lock)
        {
            if (_dirIndex.Any() && DateTime.Now - _lastRefresh < _ttl) return;

            _dirIndex.Clear();
            bool verbose = Environment.GetEnvironmentVariable("FOLDER_SEARCH_VERBOSE") == "1";
            Console.WriteLine("Building folder index… (this happens once an hour)");

            var swTotal = Stopwatch.StartNew();
            var options = new EnumerationOptions
            {
                RecurseSubdirectories = true,
                IgnoreInaccessible    = true,
                AttributesToSkip      = FileAttributes.ReparsePoint
            };

            foreach (var root in Roots.Where(Directory.Exists))
            {
                var swRoot = Stopwatch.StartNew();
                int before = _dirIndex.Count;

                try
                {
                    // fastest available recursive directory enumerator
                    _dirIndex.AddRange(Directory.EnumerateDirectories(root, "*", options));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"   ⚠︎  Skipped '{root}': {ex.Message}");
                }

                int added = _dirIndex.Count - before;
                Console.WriteLine($"   ✓  {added:N0} folders indexed under '{root}' in {swRoot.Elapsed.TotalSeconds:F1}s");

                if (verbose)
                {
                    // reveal the first few paths for that root
                    foreach (var p in _dirIndex.Skip(before).Take(3))
                        Console.WriteLine($"      • {p}");
                    if (added > 3) Console.WriteLine("      • …");
                }
            }

            _lastRefresh = DateTime.Now;
            Console.WriteLine($"Folder index ready ({_dirIndex.Count:N0} folders, {swTotal.Elapsed.TotalSeconds:F1}s).");
        }
    }
}
