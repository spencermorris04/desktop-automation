using NetOffice; // Core NetOffice namespace
using Excel = NetOffice.ExcelApi; // Alias for Excel types
using NetOffice.ExcelApi.Enums; // For Enums like XlSaveAsAccessMode, XlFileFormat
using System.IO;
using System.Diagnostics;
using System; // For Exception, DateTime, Environment, Type
using System.Threading; // Required for Thread.Sleep

public static class ExcelHelper
{
    // --- GetExcelApplication using NetOffice ---
    private static Excel.Application? GetExcelApplication(bool makeVisible = false)
    {
        Excel.Application? excelApp = null;
        try
        {
            Console.WriteLine("Creating new Excel instance using NetOffice...");
            excelApp = new Excel.Application();
            excelApp.Visible = makeVisible;
            excelApp.DisplayAlerts = false;
            Console.WriteLine($"New NetOffice Excel instance created (Visible: {makeVisible}).");
            return excelApp;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"FATAL Error creating NetOffice Excel instance: {ex.Message}");
            Console.WriteLine("Ensure Excel Desktop is installed and properly registered.");
            excelApp?.Dispose();
            return null;
        }
    }

    // --- WriteTableData using NetOffice ---
    public static bool WriteTableData(string? filePath, string sheetName, object[,] data, bool saveAndClose, bool openAfterSave)
    {
        Excel.Application? excelApp = null;
        Excel.Workbook? workbook = null;
        Excel.Worksheet? worksheet = null;
        bool success = false;
        string finalFilePath = "";

        if (string.IsNullOrEmpty(filePath))
        {
            finalFilePath = Path.Combine(Environment.CurrentDirectory, $"NewWorkbook_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
            Console.WriteLine($"No file path provided, will save as: {finalFilePath}");
        }
        else
        {
            finalFilePath = Path.GetFullPath(filePath);
        }

        excelApp = GetExcelApplication(makeVisible: openAfterSave);
        if (excelApp == null)
        {
            Console.WriteLine("Failed to create Excel Application instance via NetOffice.");
            return false;
        }

        try
        {
            excelApp.DisplayAlerts = false;
            Excel.Workbooks workbooks = excelApp.Workbooks;

            if (File.Exists(finalFilePath))
            {
                Console.WriteLine($"Opening existing workbook: {finalFilePath}");
                workbook = workbooks.Open(finalFilePath);
            }
            else
            {
                Console.WriteLine("Creating new workbook.");
                workbook = workbooks.Add();
            }

            try
            {
                worksheet = workbook.Worksheets[sheetName] as Excel.Worksheet;
                if (worksheet == null) throw new NullReferenceException($"Worksheet '{sheetName}' not found or is not a Worksheet.");
                Console.WriteLine($"Using existing worksheet: {sheetName}");
            }
            catch
            {
                Console.WriteLine($"Worksheet '{sheetName}' not found, trying first sheet or adding new.");
                if (workbook.Worksheets.Count > 0)
                {
                    worksheet = workbook.Worksheets[1] as Excel.Worksheet;
                    if (worksheet != null)
                    {
                        try
                        {
                            worksheet.Name = sheetName;
                            Console.WriteLine($"Renamed first sheet to: {sheetName}");
                        }
                        catch (Exception renameEx)
                        {
                            Console.WriteLine($"Warning: Could not rename first sheet ('{renameEx.Message}'). Using existing name '{worksheet.Name}'.");
                        }
                    }
                }
                if (worksheet == null)
                {
                    worksheet = workbook.Worksheets.Add() as Excel.Worksheet;
                    if (worksheet == null) throw new Exception("Failed to add new worksheet.");
                    worksheet.Name = sheetName;
                    Console.WriteLine($"Added new worksheet: {sheetName}");
                }
            }

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);
            if (rows > 0 && cols > 0)
            {
                using (Excel.Range range = worksheet.Cells[1, 1].Resize(rows, cols))
                {
                    Console.WriteLine($"Writing data to range {range.Address(false, false, XlReferenceStyle.xlA1)} on sheet '{worksheet.Name}'...");
                    range.Value = data;
                    range.Columns.AutoFit();
                }
            }
            else
            {
                Console.WriteLine("Warning: No data provided to write.");
            }

            Console.WriteLine($"Saving workbook to: {finalFilePath}");
            string? directory = Path.GetDirectoryName(finalFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
            workbook.SaveAs(finalFilePath, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlUserResolution, true, Type.Missing, Type.Missing, Type.Missing);
            Console.WriteLine("Workbook saved.");

            // Open logic with delay
            if (openAfterSave)
            {
                Console.WriteLine("Opening workbook using default application...");
                try
                {
                    // --- ADDED DELAY ---
                    Console.WriteLine("Waiting briefly for file lock release...");
                    Thread.Sleep(500); // Wait for 500 milliseconds

                    Process.Start(new ProcessStartInfo(finalFilePath) { UseShellExecute = true });
                    if (excelApp != null) excelApp.Visible = true;
                }
                catch (Exception openEx)
                {
                    Console.WriteLine($"Error opening Excel file after saving: {openEx.Message}");
                }
            }

            // Close workbook logic
            if (!openAfterSave && saveAndClose)
            {
                workbook.Close(false);
                Console.WriteLine("Workbook closed.");
            }
            else if (!openAfterSave && !saveAndClose)
            {
                Console.WriteLine("Workbook left open in Excel instance (as requested).");
            }
            else // openAfterSave is true
            {
                Console.WriteLine("Workbook left open in Excel instance (opened for user).");
            }

            success = true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during NetOffice Excel operation: {ex.Message}");
            Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            if (workbook != null && !workbook.IsDisposed)
            {
                try { workbook.Close(false); } catch { /* Ignore closing errors */ }
            }
        }
        finally
        {
            worksheet?.Dispose();
            workbook?.Dispose();

            if (excelApp != null && !excelApp.IsDisposed)
            {
                bool shouldQuit = saveAndClose && !openAfterSave;
                if (shouldQuit)
                {
                    int wbCount = 0;
                    try { wbCount = excelApp.Workbooks.Count; } catch { /* Ignore */ }
                    if (wbCount == 0)
                    {
                        excelApp.Quit();
                        Console.WriteLine("NetOffice Excel application instance Quit.");
                    }
                    else
                    {
                        Console.WriteLine($"NetOffice Excel application instance left open ({wbCount} workbook(s) still open).");
                    }
                }
                else
                {
                    Console.WriteLine("NetOffice Excel application instance left open.");
                }
                excelApp.Dispose();
                Console.WriteLine("NetOffice Excel application object disposed.");
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        return success;
    }
}
