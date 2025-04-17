using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using AutomationCondition = System.Windows.Automation.Condition; // Alias
using System.Windows.Automation;
using System.Windows.Forms;
using System.Windows;

public static class NotepadAutomator
{
    private static (AutomationElement? window, AutomationElement? textArea) FindNotepadControls()
    {
        AutomationElement? notepadWindow = null;
        AutomationElement? textArea = null;
        try
        {
            AutomationCondition windowCondition = new PropertyCondition(AutomationElement.ClassNameProperty, "Notepad");
            notepadWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, windowCondition);

            if (notepadWindow == null) return (null, null);

            AutomationCondition textConditionDoc = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document);
            textArea = notepadWindow.FindFirst(TreeScope.Descendants, textConditionDoc);

            if (textArea == null)
            {
                AutomationCondition textConditionEdit = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                textArea = notepadWindow.FindFirst(TreeScope.Descendants, textConditionEdit);
            }
        }
        catch (Exception ex)
        {
             Console.WriteLine($"Error finding Notepad controls: {ex.Message}");
             return (null, null);
        }
        return (notepadWindow, textArea);
    }

    public static void OpenNotepad()
    {
        try
        {
            if (Process.GetProcessesByName("notepad").Length == 0)
            {
                Console.WriteLine("Opening Notepad...");
                Process.Start("notepad.exe");
                System.Threading.Thread.Sleep(1500);
            }
            else
            {
                Console.WriteLine("Notepad is already running.");
            }
        }
        catch (Exception ex)
        {
             Console.WriteLine($"Error opening Notepad: {ex.Message}");
        }
    }

    // Modified to append text
    public static bool WriteText(string textToAppend, bool append = true)
    {
        var (notepadWindow, textArea) = FindNotepadControls();

        if (notepadWindow == null || textArea == null)
        {
            Console.WriteLine("Error: Could not find Notepad window or text area for writing.");
            return false;
        }

        try
        {
            if (notepadWindow.TryGetCurrentPattern(WindowPattern.Pattern, out object windowPatternObj))
            {
                 ((WindowPattern)windowPatternObj).SetWindowVisualState(WindowVisualState.Normal);
            }
            notepadWindow.SetFocus();
            System.Threading.Thread.Sleep(200);
        }
        catch (Exception focusEx)
        {
             Console.WriteLine($"Warning: Could not set focus/state for Notepad window: {focusEx.Message}");
        }

        string currentText = "";
        if (append)
        {
            Console.WriteLine("Reading existing text before appending...");
            // Use the ReadText method (which includes fallbacks) to get current content
            currentText = ReadTextInternal(notepadWindow, textArea) ?? ""; // Read current text
            if (!string.IsNullOrEmpty(currentText) && !currentText.EndsWith(Environment.NewLine))
            {
                // Add a newline if current text exists and doesn't end with one
                currentText += Environment.NewLine;
            }
            Console.WriteLine("Appending new text.");
        }
        else
        {
             Console.WriteLine("Replacing existing text.");
        }

        string combinedText = currentText + textToAppend;

        // --- Try writing using UIA ValuePattern ---
        try
        {
            if (textArea.TryGetCurrentPattern(ValuePattern.Pattern, out object valuePatternObj))
            {
                ((ValuePattern)valuePatternObj).SetValue(combinedText);
                Console.WriteLine("Successfully wrote text using UI Automation.");
                return true;
            }
            else
            {
                 Console.WriteLine("Warning: UI Automation ValuePattern not available. Trying SendKeys fallback.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error using UI Automation ValuePattern to write: {ex.Message}. Trying SendKeys fallback.");
        }

        // --- Fallback Method: SendKeys ---
        Console.WriteLine("Attempting fallback using SendKeys...");
        try
        {
            notepadWindow.SetFocus(); // Ensure focus again
            System.Threading.Thread.Sleep(200);

            // If replacing, clear first. If appending, just send the new text.
            if (!append)
            {
                SendKeys.SendWait("^a");
                System.Threading.Thread.Sleep(100);
                SendKeys.SendWait("{DEL}");
                System.Threading.Thread.Sleep(100);
            }
            else
            {
                 // Move cursor to the end before appending with SendKeys
                 SendKeys.SendWait("^{END}"); // Ctrl+End
                 System.Threading.Thread.Sleep(100);
                 // Add newline if needed (might already be handled by currentText logic)
                 if (!currentText.EndsWith(Environment.NewLine) && currentText.Length > 0)
                 {
                     SendKeys.SendWait("{ENTER}");
                     System.Threading.Thread.Sleep(100);
                 }
            }


            string escapedText = textToAppend.Replace("+", "{+}") // Only escape the appended part
                                     .Replace("^", "{^}")
                                     .Replace("%", "{%}")
                                     .Replace("~", "{~}")
                                     .Replace("(", "{(}")
                                     .Replace(")", "{)}")
                                     .Replace("[", "{[}")
                                     .Replace("]", "{]}");
            SendKeys.SendWait(escapedText);
            Console.WriteLine("Successfully wrote text using SendKeys fallback.");
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error using SendKeys fallback: {ex.Message}");
            return false;
        }
    }

    // Public ReadText method - finds controls first
    public static string? ReadText()
    {
        var (notepadWindow, textArea) = FindNotepadControls();
        if (notepadWindow == null || textArea == null)
        {
            Console.WriteLine("Error: Could not find Notepad window or text area for reading.");
            return null;
        }
        return ReadTextInternal(notepadWindow, textArea); // Call internal method
    }

    // Internal ReadText method - assumes controls are already found
    // Used by WriteText for appending
    private static string? ReadTextInternal(AutomationElement notepadWindow, AutomationElement textArea)
    {
        // Bring window to foreground and focus (might be redundant if called after WriteText focus)
        try
        {
            if (notepadWindow.TryGetCurrentPattern(WindowPattern.Pattern, out object windowPatternObj))
            {
                 ((WindowPattern)windowPatternObj).SetWindowVisualState(WindowVisualState.Normal);
            }
            notepadWindow.SetFocus();
            System.Threading.Thread.Sleep(100); // Shorter wait ok here
        }
        catch (Exception focusEx)
        {
             Console.WriteLine($"Warning: Could not set focus/state for Notepad window during read: {focusEx.Message}");
        }

        // --- Try reading using UIA ValuePattern ---
        try
        {
            if (textArea.TryGetCurrentPattern(ValuePattern.Pattern, out object valuePatternObj))
            {
                // Console.WriteLine("Reading text using UI Automation."); // Less verbose internally
                return ((ValuePattern)valuePatternObj).Current.Value;
            }
             else
            {
                 Console.WriteLine("Warning: UI Automation ValuePattern not available for read. Trying Clipboard fallback.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error using UI Automation ValuePattern to read: {ex.Message}. Trying Clipboard fallback.");
        }

        // --- Fallback Method: Clipboard ---
        Console.WriteLine("Attempting read fallback using Clipboard...");
        object? originalClipboardData = null;
        bool clipboardContainsText = false;
        string? clipboardText = null;

        Thread staThreadGet = new Thread(() => { /* ... get clipboard data ... */
            try {
                clipboardContainsText = System.Windows.Clipboard.ContainsText();
                if (System.Windows.Clipboard.ContainsData(System.Windows.DataFormats.UnicodeText)) {
                    originalClipboardData = System.Windows.Clipboard.GetData(System.Windows.DataFormats.UnicodeText);
                }
            } catch (Exception ex) { Console.WriteLine($"Warning: Could not get original clipboard data: {ex.Message}"); }
        });
        staThreadGet.SetApartmentState(ApartmentState.STA);
        staThreadGet.Start();
        staThreadGet.Join();

        try
        {
            notepadWindow.SetFocus(); // Ensure focus
            System.Threading.Thread.Sleep(100);

            SendKeys.SendWait("^a");
            System.Threading.Thread.Sleep(100);
            SendKeys.SendWait("^c");
            System.Threading.Thread.Sleep(200);

            Thread staThreadRead = new Thread(() => { /* ... read clipboard data ... */
                 try {
                    if (System.Windows.Clipboard.ContainsText()) {
                        clipboardText = System.Windows.Clipboard.GetText();
                    }
                } catch (Exception ex) { Console.WriteLine($"Error accessing clipboard: {ex.Message}"); }
            });
            staThreadRead.SetApartmentState(ApartmentState.STA);
            staThreadRead.Start();
            staThreadRead.Join();

            if (clipboardText != null) {
                 Console.WriteLine("Successfully read text using Clipboard fallback.");
                 return clipboardText;
            } else {
                Console.WriteLine("Failed to read text using Clipboard fallback.");
                return null;
            }
        }
        catch (Exception ex) {
            Console.WriteLine($"Error during Clipboard fallback operation: {ex.Message}");
            return null;
        }
        finally {
             Thread staThreadRestore = new Thread(() => { /* ... restore clipboard data ... */
                 try {
                    System.Windows.Clipboard.Clear();
                    if (originalClipboardData != null) {
                         System.Windows.Clipboard.SetData(System.Windows.DataFormats.UnicodeText, originalClipboardData);
                    } else if (clipboardContainsText) {
                         System.Windows.Clipboard.Clear();
                    }
                } catch (Exception ex) { Console.WriteLine($"Warning: Could not restore original clipboard data: {ex.Message}"); }
             });
            staThreadRestore.SetApartmentState(ApartmentState.STA);
            staThreadRestore.Start();
            staThreadRestore.Join();
        }
    }
}
