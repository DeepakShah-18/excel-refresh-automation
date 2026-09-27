# ─────────────────────────────────────────
# CONFIGURATION — Edit these values only
# ─────────────────────────────────────────

EXCEL_FILE = r"File1.xlsx"

REFRESH_WAIT    = 10   # seconds between refreshes (on success)
CLOSE_WAIT      = 10   # seconds after final refresh before saving (on success)
SETTLE_WAIT     = 3    # seconds to let Excel settle right after opening
REFRESH_TIMEOUT = 50   # max seconds allowed for ONE refresh/save call before treating it as stuck

# ─────────────────────────────────────────
# SCRIPT — Do not edit below this line
# ─────────────────────────────────────────

import os
import time
import subprocess
import threading
import pythoncom
import win32com.client
import win32process


def normalize(path):
    return os.path.normcase(os.path.abspath(path))


def force_kill_pid(pid):
    if not pid:
        return
    try:
        print(f"🔪 Force-killing stuck Excel process (PID {pid})...")
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True, timeout=15)
    except Exception as e:
        print(f"⚠️ Could not force-kill Excel process: {e}")


def get_pid(excel):
    try:
        hwnd = excel.Hwnd
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid
    except Exception:
        return None


def find_open_instance(filepath):
    """Check if the target file is already open in a running Excel session."""
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return None, None
    target = normalize(filepath)
    try:
        for wb in excel.Workbooks:
            try:
                if normalize(wb.FullName) == target:
                    return excel, wb
            except Exception:
                continue
    except Exception:
        pass
    return None, None


def open_new_instance(filepath):
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    wb = excel.Workbooks.Open(
        filepath, UpdateLinks=0, ReadOnly=False,
        IgnoreReadOnlyRecommended=True, Notify=False,
    )
    return excel, wb


def get_or_open_workbook(filepath):
    excel, wb = find_open_instance(filepath)
    if wb is not None:
        print("📎 File is already open — attaching to it instead of opening a new copy.")
        return excel, wb, False  # owns_it = False
    print("🔄 Opening Excel (new instance)...")
    excel, wb = open_new_instance(filepath)
    print("✅ File opened")
    return excel, wb, True  # owns_it = True


def watchdog(pid, timeout, owns_it, stop_event, flag):
    start = time.time()
    while not stop_event.is_set():
        if time.time() - start > timeout:
            if owns_it and pid:
                force_kill_pid(pid)
                flag["killed"] = True
            else:
                flag["timed_out"] = True
            return
        time.sleep(1)


def run_with_timeout(label, func, timeout, pid, owns_it):
    """Runs func() (blocking COM call) on THIS thread; watchdog kills by PID if it hangs."""
    stop_event = threading.Event()
    flag = {}
    wd = threading.Thread(target=watchdog, args=(pid, timeout, owns_it, stop_event, flag), daemon=True)
    wd.start()
    try:
        func()
        stop_event.set()
        return True
    except Exception as e:
        stop_event.set()
        if flag.get("killed"):
            print(f"⏱️ {label} exceeded {timeout}s — process force-killed.")
        elif flag.get("timed_out"):
            print(f"⏱️ {label} exceeded {timeout}s — file was already open before script ran, so not force-killing (would affect your other open work). It may still be stuck.")
        else:
            print(f"❌ {label} error: {e}")
        return False


def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Path does not exist — check the file path: {EXCEL_FILE}")
        return

    print("=" * 50)
    print("🚀 Single File Excel Refresh Tool")
    print("=" * 50)

    pythoncom.CoInitialize()
    try:
        excel, wb, owns_it = get_or_open_workbook(EXCEL_FILE)
        pid = get_pid(excel) if owns_it else None

        print(f"⏳ Letting Excel settle {SETTLE_WAIT}s...")
        time.sleep(SETTLE_WAIT)

        for i in range(1, 4):
            print(f"\n🔄 Refresh {i}/3...")

            def do_refresh():
                wb.RefreshAll()
                excel.CalculateUntilAsyncQueriesDone()

            ok = run_with_timeout(f"Refresh {i}/3", do_refresh, REFRESH_TIMEOUT, pid, owns_it)

            if ok:
                print(f"✅ Refresh {i}/3 complete")
            else:
                if owns_it:
                    print("♻️ Reopening a fresh instance to continue...")
                    excel, wb = open_new_instance(EXCEL_FILE)
                    pid = get_pid(excel)
                    time.sleep(SETTLE_WAIT)
                else:
                    print("⚠️ Continuing with the same session (not killed, since it was already open before the script ran).")

            wait = REFRESH_WAIT if i < 3 else CLOSE_WAIT
            print(f"⏳ Waiting {wait} seconds...")
            time.sleep(wait)

        print("\n💾 Saving...")
        ok = run_with_timeout("Save", wb.Save, REFRESH_TIMEOUT, pid, owns_it)
        print("✅ Saved successfully" if ok else "❌ Save failed or hung")

        if owns_it:
            try:
                wb.Close(SaveChanges=not ok)  # if save succeeded above, no need to re-save on close
                print("🔒 Workbook closed")
            except Exception as e:
                print(f"⚠️ Workbook close failed: {e}")
            try:
                excel.Quit()
                print("🔒 Excel instance closed")
            except Exception as e:
                print(f"⚠️ Excel Quit() failed: {e}")
        else:
            print("📎 Left the file open — it was already open before this script ran.")

        print(f"\n{'=' * 50}")
        print(f"📊 RESULT: {'✅ Success' if ok else '⚠️ Completed with issues — check log above'}")
        print(f"{'=' * 50}")

    finally:
        pythoncom.CoUninitialize()

    print("\n🎉 Done!")


if __name__ == "__main__":
    main()
