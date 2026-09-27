# Excel Refresh Automation

Windows automation scripts that refresh Excel workbooks with live data connections (Power Query, external data sources, pivot tables) on a schedule — without needing anyone to open Excel manually.

Built to run unattended via **Windows Task Scheduler**, these scripts drive Excel through COM automation to open each file, refresh its data connections multiple times, save, and close — while handling the messy real-world edge cases: a file that's already open, a refresh that hangs, or a stuck Excel process left over from a previous run.

## Results

Once scheduled, these scripts remove the need for anyone to manually open and refresh reports before they're used downstream:

- **Power Automate email flows** that read from these Excel files can send out reports with current, refreshed data — no need to manually refresh before a scheduled email goes out
- **Power BI reports and dashboards** connected to these files as a data source always pull fresh, up-to-date figures on their own refresh cycle, rather than stale data from whenever the file was last opened by hand
- **Reports stay current even when the file's owner is away** — on leave, off sick, or simply not at their desk — since the refresh runs unattended on schedule rather than depending on someone remembering to open and refresh it

## What's in this repo

| Script | Purpose |
|---|---|
| `Refresh_All_Files.py` | Refreshes a list of Excel files, one after another, each refreshed twice before saving |
| `Single_File_refresh.py` | Refreshes a single Excel file three times before saving |

Both scripts share the same core design.

## Key features

- **Runs Excel visibly** — so you can watch it open and refresh in real time, rather than working invisibly in the background
- **Won't disturb an Excel file you already have open** — if the target file is already open in your own session, the script attaches to that session, refreshes it, and saves — without closing it afterward. If the script opened the file itself, it closes it cleanly when done
- **Detects and recovers from a stuck refresh** — each refresh attempt has a timeout (default 50 seconds). If Excel hangs on a refresh, the script force-kills that specific Excel process by its process ID (never touching any other Excel window you may have open) and retries fresh
- **Never runs the refresh twice for nothing** — a single open → refresh → refresh → save → close cycle per file, with no redundant re-opens
- **Skips invalid paths safely** — if a file path doesn't exist, it's flagged and skipped rather than the whole run failing

## Requirements

- Windows with Microsoft Excel installed
- Python 3.x
- `pywin32` package:
  ```
  pip install pywin32
  ```

## Setup

1. Clone or download this repo
2. Open the script you want to use and find the **CONFIGURATION** section at the top
3. Replace the placeholder paths with your own file paths — for example:
   ```python
   EXCEL_FILES = [
       r"C:\Users\YourName\Documents\Reports\YourFile1.xlsx",
       r"C:\Users\YourName\Documents\Reports\YourFile2.xlsx",
   ]
   ```
4. Adjust the timing settings if needed:
   - `REFRESH_WAIT` — seconds between refresh passes
   - `CLOSE_WAIT` — seconds after the final refresh before saving
   - `SETTLE_WAIT` — seconds to let Excel settle right after opening
   - `REFRESH_TIMEOUT` — max seconds allowed for one refresh/save call before it's treated as stuck

## Running manually

```
python Refresh_All_Files.py
```
or
```
python Single_File_refresh.py
```

## Scheduling with Task Scheduler

1. Open **Task Scheduler** → **Create Basic Task**
2. Set a **Daily** trigger at your preferred time (add multiple triggers on the same task if you need it to run more than once a day)
3. Action → **Start a program**:

   | Field | Value |
   |---|---|
   | Program/script | Full path to `python.exe` (find it with `where python`) |
   | Add arguments | The script filename, e.g. `Refresh_All_Files.py` |
   | Start in | The folder containing the script |

4. On the **Conditions** tab:
   - Untick **"Start the task only if the computer is on AC power"**
   - Tick **"Wake the computer to run this task"**
5. On the **General** tab, set **"Run only when user is logged on"** if you want to see the Excel window appear on screen when it runs

## Notes

- If files live in a OneDrive-synced folder, right-click each file in File Explorer → **"Always keep on this device"** to avoid slow first-time downloads when the script opens them
- If a workbook's data connections use **background refresh**, consider disabling it (**Data → Queries & Connections → right-click query → Properties → untick "Enable background refresh"**) — automation is more stable with synchronous refreshes
