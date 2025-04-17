
# 🖥️ Desktop Automator Demo

**Desktop Automator Demo** is a C# /.NET 9 (Windows‑only) console application
that turns your keyboard—or a future voice assistant—into a Swiss‑army knife
for everyday desktop tasks:

*   open & write to Notepad  
*   kill runaway processes, watch live CPU/RAM  
*   list and switch audio devices, change volume, play/pause media  
*   fuzzy‑search recent downloads and open files  
*   paste CSV → instantly create/open an Excel workbook  
*   focus any window by fuzzy name *(“open chrome”)* or tile groups of windows  
*   fuzzy‑search every folder under your profile and open it  
*   sample process metrics in the background while you work  
*   list, activate, open, and close Chrome tabs via a local Chrome extension bridge  

The repo is intentionally **single‑executable, no external services**—just
`dotnet run` and go.

---

## Table of Contents
1. [Quick Start](#quick-start)  
2. [Project Structure](#project-structure)  
3. [Features in Detail](#features-in-detail)  
4. [Chrome Tab Bridge](#chrome-tab-bridge)  
5. [Configuration & Environment Variables](#configuration--environment-variables)  
6. [Extending the CLI](#extending-the-cli)  
7. [Troubleshooting](#troubleshooting)  
8. [License](#license)  

---

## Quick Start

```powershell
git clone https://github.com/YourUser/DesktopAutomatorDemo.git
cd DesktopAutomatorDemo

# build & run (requires .NET 9 preview SDK)
dotnet run
```

> **Tip:** on first launch the folder index builds (few seconds), then searches are instant for 60 minutes.

---

## Project Structure

| Path                         | Description                                                     |
|------------------------------|-----------------------------------------------------------------|
| `DesktopAutomatorDemo.csproj` | SDK‑style project file; targets `net9.0-windows`.               |
| `Program.cs`                 | CLI entry point and dispatch logic.                             |
| `WindowHelper.cs`            | Enumerates windows, fuzzy‑search, and reliable foreground focus. |
| `WindowLayoutManager.cs`     | Predefined layouts for coding / research environments.          |
| `FolderSearchHelper.cs`      | Recursive folder index + fuzzy search + open.                  |
| `TaskManagerHelper.cs`       | Background CPU/RAM sampler and process kill.                   |
| `AudioHelper.cs`             | Enumerate/set audio devices, volume, media key simulation.      |
| `NotepadAutomator.cs`        | UI Automation / SendKeys for Notepad.                          |
| `FileExplorerHelper.cs`      | Open Downloads folder and recent files.                        |
| `ExcelHelper.cs`             | CSV to Excel importer using NetOffice.                         |
| `ChromeExtension/`           | Chrome extension (manifest, background.js, poll.js).           |

---

## Chrome Tab Bridge

A lightweight Chrome extension listens on **http://127.0.0.1:9234** to relay
tab commands from the console app.

### Installation

1. In Chrome, navigate to `chrome://extensions`.  
2. Enable **Developer mode**.  
3. Click **Load unpacked** and select the `ChromeExtension/` folder.  
4. Ensure the extension is enabled and see logs via **Service worker → Inspect**.

### Usage

1. Run the console app (`dotnet run`).  
2. Select **Chrome Tabs** from the main menu (option 3 or `tabs`).  
3. Commands:
   - **`[number]`**: activate the tab with that index  
   - **`o <url>`**: open a new tab at `<url>`  
   - **`c <number>`**: close the tab at that index  
   - **`q`**: return to main menu  

Tab activation also brings the Chrome window to the foreground, even with multiple windows open.

---

## Features in Detail

### Option 1 – Notepad Automation
- **Open or reuse** Notepad via UI Automation (ValuePattern) or fallback to SendKeys.  
- **Append**, **overwrite**, or **read** the text in the active Notepad window.  
- Ideal for quick prototyping of text‑entry workflows without manual GUI scripting.

### Option 2 – Excel Helper
- **Stub** implementation in this demo.  
- Intended to wrap **NetOffice.ExcelApi** for CSV → .xlsx import.  
- Prompts for sheet name and writes a 2D array, autofits columns.

### Option 3 – Chrome Tabs
- **List**, **activate**, **open** and **close** Chrome tabs via a local HTTP bridge.  
- Menu commands:
  - `[number]` – activate that tab (inside Chrome + OS foreground).  
  - `o <url>` – open a new tab at `<url>` and focus its Chrome window.  
  - `c <number>` – close the specified tab.  
  - `q` – return to main menu.  
- Works with **multiple Chrome windows** by matching Chrome’s own `windowId` to the OS window.

### Option 4 – Open Downloads Folder
- Launches File Explorer in the **Downloads** directory under your user profile.  
- Uses `ProcessStartInfo { UseShellExecute = true }` for native association.

### Option 5 – Search & Open Folder
- Recursively indexes `%USERPROFILE%`, Desktop, Documents, and Downloads.  
- Caches the index for **60 minutes** (`_ttl`).  
- Uses **FuzzySharp** (`WeightedRatio`) on folder names for quick filtering.  
- Opens the selected folder via `explorer.exe`.

### Option 6 – Switch / Focus Window
- Fuzzy‑search open windows by **process name + title**.  
- Restores minimized windows and brings the target to the **true foreground**, working around Win32 focus‑lock rules:
  - Attaches thread inputs,  
  - Uses `SetForegroundWindow`,  
  - Falls back to `SwitchToThisWindow` or simulated Alt‑key.

### Option 7 – Apply Window Layout
- Predefined layouts:
  - **coding**: VS Code on left 50%; Notepad + Terminal on right half (split top/bottom).  
  - **research**: Chrome | Obsidian side‑by‑side.  
- Prompts once per application instance, then tiles **all** windows in one pass.  
- Respects the task‑bar working area and compensates for window frames.

### Option 8 – Task Manager
- Background sampler thread polls WMI’s `Win32_PerfFormattedData_PerfProc_Process` every second.  
- **List** top 20 processes by CPU (%) and RAM (MB).  
- **Terminate** a process by PID.

### Option 9 – Audio Mute / Unmute
- Enumerate playback endpoints via **NAudio.CoreAudioApi**.  
- Get or set the **default** playback device.  
- Read or adjust **master volume** (0–100 %).  
- Send the **Play/Pause** multimedia key via `SendInput`.

---

## Configuration & Environment Variables

| Variable                 | Effect                                                       |
|--------------------------|--------------------------------------------------------------|
| `FOLDER_SEARCH_VERBOSE=1` | Show folder samples during index build.                      |
| `FOLDER_SEARCH_TTL_MIN`   | Override folder index TTL (minutes).                         |
| `DOTNET_EnablePreviewFeatures` | Enable .NET 9 preview features if using preview SDK. |

---

## Extending the CLI

1. Add a menu item in `Program.cs`.  
2. Implement a static helper class.  
3. Update `.csproj` for new packages.  
4. For new background tasks, mimic `TaskManagerHelper`.

---

## Troubleshooting

### 1. Chrome Tab Bridge

#### Extension won’t load  
- **Symptom:** You don’t see your service‑worker in `chrome://extensions`.  
- **Cause:** Manifest errors, missing `host_permissions`, or wrong folder structure.  
- **Fix:**  
  1. Verify `manifest.json` is in the root of your `ChromeExtension/` folder.  
  2. Ensure `"manifest_version": 3`, `"background.service_worker"` points to `background.js`, and you have  
     ```json
     "permissions": ["tabs","alarms"],
     "host_permissions": ["http://127.0.0.1:9234/*"]
     ```  
  3. Reload the extension in Developer mode and inspect the service‑worker console for errors.

#### `/pending` always returns `{}`  
- **Symptom:** The console app immediately gets “Task was canceled.”  
- **Cause:** JSON field‐name mismatch—extension expects `"action"`/`"reqId"` but bridge was sending `"Action"`/`"ReqId"`.  
- **Fix:**  
  - In `ChromeBridgeServer.Handle("/pending")`, emit camel‑case:  
    ```csharp
    await RespondJson(ctx, new { action = cmd.Action, reqId = cmd.ReqId, args = cmd.Args });
    ```

#### Service‐worker unloading too quickly  
- **Symptom:** Poll loop stops after ~30 s of inactivity.  
- **Cause:** MV3 service‑workers sleep when idle.  
- **Fix:**  
  - Use `chrome.alarms.create('keepAlive', { periodInMinutes:1 })` in `background.js`.  
  - On each alarm, call your poll routine again to reset the idle timer.

---

### 2. Timeouts & Canceled Tasks

- **Symptom:** C# prints “A task was canceled.” after 8 s.  
- **Cause:** Bridge’s `RequestAsync` default timeout (`timeoutMs`) was reached.  
- **Fix:**  
  - Pass a larger `timeoutMs` to `RequestAsync(...)` if you expect delays (e.g. in `getTabTitle`).  
  - Always handle `TaskCanceledException` gracefully.

---

### 3. Window Focus Issues

#### Chrome window not coming forward  
- **Symptom:** Tab activates inside Chrome, but the OS window stays in the background.  
- **Cause:** Windows “foreground lock” prevents apps from stealing focus.  
- **Fix:**  
  - Use our `WindowHelper.FocusChromeWindow(windowId)` which:  
    1. Matches Chrome’s `windowId` via the native `ChromeWindowId` property.  
    2. Calls `AttachThreadInput` → `SetForegroundWindow` → fallback strategies.  
  - Ensure you have the latest `WindowHelper.cs` with the `FocusChromeWindow` implementation.

---

### 4. Folder Search & Open

- **Symptom:** First search is extremely slow or crashes on certain directories.  
- **Cause:** Recursive scan of large or protected folders.  
- **Fix:**  
  - Set `FOLDER_SEARCH_VERBOSE=1` to see which paths take longest.  
  - Exclude network shares or very large directories by modifying `Roots` in `FolderSearchHelper`.  
  - Increase the fuzzy‑match cutoff to reduce results.

---

### 5. Excel CSV Importer

- **Symptom:** COM exception or “Class not registered.”  
- **Cause:** Office not installed or NetOffice registration missing.  
- **Fix:**  
  - Install desktop Excel (2016+).  
  - Ensure `NetOffice` NuGet package is up to date.  
  - Run the app as a user with COM registration privileges.

---

### 6. Task Manager Helper

- **Symptom:** No processes shown or sample throws WMI errors.  
- **Cause:** WMI (Win32_PerfFormattedData_PerfProc_Process) sometimes requires elevated privileges.  
- **Fix:**  
  - Run the console as Administrator.  
  - Check WMI service (`winmgmt`) is running:  
    ```powershell
    Get-Service winmgmt | Select Status
    ```

---

### 7. Audio Manager

- **Symptom:** “Audio control failed” or no devices listed.  
- **Cause:** COM interface (`IPolicyConfig`) may fail if office codecs are missing.  
- **Fix:**  
  - Ensure you’re running on Windows 10/11 with the Audio Graph API available.  
  - If enumeration fails, catch exceptions and fall back to `Nircmd` or another external tool.

---

### 8. General .NET / Environment

- **Symptom:** Build errors referencing missing dependencies (e.g. `FuzzySharp`).  
- **Cause:** NuGet packages not restored.  
- **Fix:**  
  ```bash
  dotnet restore
  dotnet build
