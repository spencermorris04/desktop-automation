# 🖥️ Desktop Automator Demo

**Desktop Automator Demo** is a C# /.NET 9 (Windows‑only) console application
that turns your keyboard—or a future voice assistant—into a Swiss‑army knife
for everyday desktop tasks:

*   open & write to Notepad
*   kill runaway processes, watch live CPU/RAM
*   list and switch audio devices, change volume, play/pause media
*   fuzzy‑search recent downloads and open files
*   paste CSV → instantly create/open an Excel workbook
*   focus any window by fuzzy name *(“open chrome”)*  
    or tile groups of windows with a single command (e.g. *coding layout*)
*   fuzzy‑search every folder under your profile and open it
*   sample process metrics in the background while you work

The repo is intentionally **single‑executable, no external services**—just
`dotnet run` and go. Use it as:

* a reference project for UI‑automation, process telemetry, audio APIs, fuzzy
  matching, window management (Win32)  
* a starting point for a voice‑controlled desktop butler (the CLI structure is
  ready for OpenAI’s streaming API)

---

## Table of Contents
1. [Quick Start](#quick-start)
2. [Project Structure](#project-structure)
3. [Features in Detail](#features-in-detail)
4. [Configuration & Environment Variables](#configuration--environment-variables)
5. [Extending the CLI](#extending-the-cli)
6. [Troubleshooting](#troubleshooting)
7. [License](#license)

---

## Quick Start

```powershell
git clone https://github.com/YourUser/DesktopAutomatorDemo.git
cd DesktopAutomatorDemo

# build & run (requires .NET 9 preview SDK)
dotnet run
```

> **Tip:** first launch will build a folder index—takes a few seconds, then
> searches are instant for the next 60 minutes.

---

## Project Structure

| Path | Description |
|------|-------------|
| **DesktopAutomatorDemo.csproj** | SDK‑style project file; targets `net9.0-windows`; pulls NuGet dependencies such as **FuzzySharp** and **NAudio**. |
| **Program.cs** | CLI‑oriented entry point. Displays the menu and dispatches to feature handlers. Starts/stops the background sampler thread. |
| **WindowHelper.cs** | Enumerates windows, fuzzy‑searches them, and *reliably* brings one to the foreground (works around Win32 foreground‑lock rules). |
| **WindowLayoutManager.cs** | Hard‑coded “coding” and “research” layouts. Prompts once for every *UseExisting* window, then tiles all windows edge‑to‑edge inside the task‑bar‑safe working area. |
| **FolderSearchHelper.cs** | Scans every folder under your user profile (recursively, fast) and provides fuzzy search + `explorer.exe` open. Verbose progress logging on first build. |
| **TaskManagerHelper.cs** | Background thread that queries WMI **Win32_PerfFormattedData_PerfProc_Process** every second; caches CPU/RAM; supports `Kill(pid)`. |
| **AudioHelper.cs** | Uses **NAudio** + undocumented `IPolicyConfig` COM to enumerate/set playback devices, get/set master volume, and simulate the Play/Pause media key. |
| **NotepadAutomator.cs** | UI Automation + SendKeys fallbacks to open Notepad, append/replace/read its text. |
| **FileExplorerHelper.cs** | `Downloads` helper—lists latest files or fuzzy search results; opens selection with the associated program. |
| **ExcelHelper.cs** | Wraps **NetOffice** (late‑bound COM) to create/open an Excel workbook, write a 2‑D array, autofit columns, optionally reopen for the user. |
| **ComHelper.cs** | Generic “get running COM object” helper—used by Excel in earlier iterations. |
| **README.md** | (you are here) |

---

## Features in Detail

### 1  Notepad Automation
* `OpenNotepad()` launches or reuses Notepad.
* Append, overwrite, or read text.  
  Attempts **ValuePattern** → clipboard fallback → `SendKeys`.

### 2  Task Manager
* Background sampler thread (`TaskManagerHelper.StartBackgroundSampler`).
* Uses WMI so no perf‑counters registration needed.
* CLI: list top 20 processes by CPU/RAM or terminate by PID.

### 3  Audio Manager
* Enumerate playback devices, mark default.
* Set new default (works on Windows 10/11 without restarting audio graph).
* Read/set master volume (0‑100 %).
* Send Play/Pause multimedia key.

### 4  Downloads Helper
* Lists 10 most recent files or fuzzy search (`Fuzz.WeightedRatio`) results.
* Opens the selected file via `ProcessStartInfo { UseShellExecute = true }`.

### 5  Excel CSV Importer
* Paste CSV, type `ENDCSV`, choose sheet name.
* Creates **.xlsx** under `./NewWorkbook_yyyyMMddHHmmss.xlsx` (or opens an
  existing one).
* Uses `NetOffice.ExcelApi`—no Excel Interop reference required.

### 6  Window Switcher
* Menu #6 (*Switch / Focus Window*).
* Typing “chrome memes” will list every Chrome tab whose title matches.
* Brings the selected window to the real foreground (no task‑bar flashing).

### 7  Window Layouts
* **coding**: VS Code left 50 %, new Notepad top‑right, new Windows Terminal
  bottom‑right.  
  If multiple Code instances exist, you pick which.  
  Layout logic waits for *all* prompts first, then tiles in one shot.
* **research**: Chrome | Obsidian side‑by‑side.  
  Same prompt logic; uses existing windows.

Implementation notes:
* Rectangles derived from `Screen.PrimaryScreen.WorkingArea` → never overlap
  the task‑bar.  
* 6‑pixel compensation removes visible gap caused by window frames.

### 8  Folder Search & Open
* Scans roots (`%USERPROFILE%`, Desktop, Documents, Downloads) with
  `EnumerationOptions.RecurseSubdirectories=true` & `IgnoreInaccessible=true`.
* Caches index for 60 minutes; first build logs progress & timing.
* Environment variable `FOLDER_SEARCH_VERBOSE=1` prints sample paths.

---

## Configuration & Environment Variables

| Variable | Effect |
|----------|--------|
| `FOLDER_SEARCH_VERBOSE=1` | Show per‑folder samples when building folder index. |
| `FOLDER_SEARCH_TTL_MIN` | Override the 60‑minute index cache (integer). |
| `DOTNET_EnablePreviewFeatures=1` | Needed when using .NET 9 preview SDK. |

---

## Extending the CLI

1. **Add a menu item** in `Program.cs` (`switch(choice) { … }`).
2. Implement a static helper class *(follow the pattern of `AudioHelper`)*.
3. Reference additional NuGet packages by editing `DesktopAutomatorDemo.csproj`.
4. If you need a new long‑running background task, mimic
   `TaskManagerHelper.StartBackgroundSampler()`.

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| **Excel helper fails** with COM exception | Ensure desktop Excel is installed; COM registration is required for NetOffice. |
| Windows fail to tile on multi‑monitor setup | The layout code uses `PrimaryScreen`. Adapt `Screen.AllScreens` if you need per‑monitor layouts. |
| “Foreground window permission” denied | The `WindowHelper` already uses `AllowSetForegroundWindow(ASFW_ANY)`; if still blocked, disable focus‑assist or 3rd‑party utilities that mess with input queues. |
| Folder search still slow | Exclude large network shares from `Roots`, or raise `cutoff` in the CLI to reduce candidate set. |