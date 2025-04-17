using System;
using System.Collections.Concurrent; // For thread-safe dictionary
using System.Collections.Generic;
using System.ComponentModel; // For Win32Exception
using System.Diagnostics; // For Process class (TerminateProcess)
using System.Linq;
using System.Management; // For WMI
using System.Threading; // For Thread, CancellationTokenSource

public static class TaskManagerHelper
{
    // Data record remains the same
    public record ProcessInfo(int Id, string Name, double CpuUsage, long RamUsageMb, double GpuUsage);

    // Thread-safe cache for the latest process data
    private static ConcurrentDictionary<int, ProcessInfo> _latestProcessData = new ConcurrentDictionary<int, ProcessInfo>();

    // Background thread management
    private static Thread? _samplerThread;
    private static CancellationTokenSource? _cts;
    private static readonly TimeSpan _sampleInterval = TimeSpan.FromSeconds(1); // How often to poll

    /// <summary>
    /// Starts the background thread that continuously samples process info.
    /// Should be called once when the application starts.
    /// </summary>
    public static void StartBackgroundSampler()
    {
        if (_samplerThread != null && _samplerThread.IsAlive)
        {
            Console.WriteLine("DEBUG: Background sampler already running.");
            return; // Already running
        }

        _cts = new CancellationTokenSource();
        _samplerThread = new Thread(BackgroundSamplerLoop);
        _samplerThread.IsBackground = true; // Allows app to exit even if thread is running
        _samplerThread.Name = "ProcessInfoSampler";
        _samplerThread.Start(_cts.Token); // Pass token to the thread method
        Console.WriteLine("DEBUG: Started background process sampler thread.");
    }

    /// <summary>
    /// Signals the background sampler thread to stop.
    /// Should be called once when the application is exiting.
    /// </summary>
    public static void StopBackgroundSampler()
    {
        if (_cts != null && !_cts.IsCancellationRequested)
        {
            Console.WriteLine("DEBUG: Requesting background sampler stop...");
            _cts.Cancel();
        }

        if (_samplerThread != null && _samplerThread.IsAlive)
        {
            Console.WriteLine("DEBUG: Waiting for background sampler thread to finish...");
            if (!_samplerThread.Join(TimeSpan.FromSeconds(2)))
            {
                 Console.WriteLine("Warning: Background sampler thread did not stop gracefully within timeout.");
            } else {
                 Console.WriteLine("DEBUG: Background sampler thread stopped.");
            }
        }
        _cts?.Dispose();
        _cts = null;
        _samplerThread = null;
    }

    /// <summary>
    /// The main loop for the background sampler thread.
    /// </summary>
    private static void BackgroundSamplerLoop(object? cancellationTokenObj)
    {
        if (cancellationTokenObj == null || !(cancellationTokenObj is CancellationToken))
        {
             Console.WriteLine("Error: Invalid cancellation token passed to sampler thread.");
             return;
        }
        var cancellationToken = (CancellationToken)cancellationTokenObj;

        Console.WriteLine("DEBUG: BackgroundSamplerLoop started.");
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                // Query WMI and update the cache
                QueryWmiAndUpdateCache();

                // Wait for the next interval, checking for cancellation
                bool cancelled = cancellationToken.WaitHandle.WaitOne(_sampleInterval);
                if (cancelled) break; // Exit loop if cancellation requested during wait
            }
            // REMOVED: catch (ThreadAbortException) block - Obsolete and unnecessary with CancellationToken
            catch (OperationCanceledException)
            {
                 Console.WriteLine("DEBUG: Sampler thread cancelled via OperationCanceledException.");
                 break; // Expected when cancellation is requested
            }
            catch (Exception ex)
            {
                // Log unexpected errors but keep the loop running if possible
                Console.WriteLine($"Error in background sampler loop: {ex.Message}");
                // Optional: Add a shorter delay after an error to prevent spamming
                // Avoid sleeping if cancellation is requested immediately after error
                if (!cancellationToken.IsCancellationRequested)
                {
                    Thread.Sleep(5000);
                }
            }
        }
        Console.WriteLine("DEBUG: BackgroundSamplerLoop finished.");
    }

    /// <summary>
    /// Queries WMI for process data and updates the internal cache.
    /// </summary>
    private static void QueryWmiAndUpdateCache()
    {
        // Console.WriteLine("DEBUG: Querying WMI..."); // Can be too verbose
        var currentData = new Dictionary<int, ProcessInfo>(); // Temp dictionary for current snapshot
        ManagementObjectSearcher? searcher = null;
        try
        {
            searcher = new ManagementObjectSearcher(
                "SELECT IDProcess, Name, PercentProcessorTime, WorkingSet FROM Win32_PerfFormattedData_PerfProc_Process"
            );

            using (var results = searcher.Get())
            {
                foreach (ManagementObject obj in results)
                {
                    using (obj) // Ensure ManagementObject is disposed
                    {
                        try
                        {
                            int id = Convert.ToInt32(obj["IDProcess"]);
                            string name = obj["Name"]?.ToString() ?? "";
                            if (name == "Idle" || name == "_Total") continue;

                            double cpu = 0.0;
                            if (obj["PercentProcessorTime"] != null)
                                // WMI value is per core, divide by core count for total %
                                cpu = Convert.ToDouble(obj["PercentProcessorTime"]) / Environment.ProcessorCount;
                                cpu = Math.Max(0.0, Math.Round(cpu, 2)); // Ensure non-negative and round

                            long ram = 0;
                            if (obj["WorkingSet"] != null)
                                ram = Convert.ToInt64(obj["WorkingSet"]) / (1024 * 1024);

                            currentData.Add(id, new ProcessInfo(id, name, cpu, ram, 0.0));
                        }
                        catch
                        {
                            // Ignore individual process errors (e.g., parsing, process exited)
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"WMI query error: {ex.Message}");
            return; // Don't update cache if query failed
        }
        finally
        {
             // No explicit dispose needed for ManagementObjectSearcher typically
        }

        // --- Update the ConcurrentDictionary Cache ---
        // 1. Add or Update entries found in the current WMI query
        foreach (var kvp in currentData)
        {
            _latestProcessData.AddOrUpdate(kvp.Key, kvp.Value, (key, existingVal) => kvp.Value);
        }

        // 2. Remove entries from cache that are no longer in the current WMI data (process exited)
        var keysToRemove = _latestProcessData.Keys.Except(currentData.Keys).ToList();
        foreach (var key in keysToRemove)
        {
            _latestProcessData.TryRemove(key, out _);
        }
        // Console.WriteLine($"DEBUG: Cache updated. Count: {_latestProcessData.Count}"); // Can be too verbose
    }


    /// <summary>
    /// Gets the latest cached list of running processes.
    /// Does NOT trigger a new WMI query.
    /// </summary>
    public static List<ProcessInfo> GetCachedProcesses()
    {
        // Return a snapshot of the current cache values
        return _latestProcessData.Values.ToList();
    }

    /// <summary>
    /// Attempts to terminate a process by its ID.
    /// </summary>
    public static bool TerminateProcess(int processId)
    {
        Process? process = null; // Use System.Diagnostics.Process here
        try
        {
            process = Process.GetProcessById(processId);
            Console.WriteLine($"Attempting to terminate process '{process.ProcessName}' (ID: {processId})...");
            process.Kill();
            process.WaitForExit(5000); // Wait up to 5 seconds
            if (process.HasExited)
            {
                Console.WriteLine("Process terminated successfully.");
                _latestProcessData.TryRemove(processId, out _);
                return true;
            }
            else
            {
                Console.WriteLine("Warning: Process did not exit after Kill command.");
                return false;
            }
        }
        catch (ArgumentException) { Console.WriteLine($"Error: Process with ID {processId} not found."); }
        catch (Win32Exception ex) { Console.WriteLine($"Error: Access denied terminating PID {processId}. ({ex.Message})"); }
        catch (InvalidOperationException ex) { Console.WriteLine($"Error: Cannot terminate PID {processId} (system process or already exiting?). ({ex.Message})"); }
        catch (Exception ex) { Console.WriteLine($"Unexpected error terminating PID {processId}: {ex.Message}"); }
        finally { process?.Dispose(); }
        return false;
    }
}
