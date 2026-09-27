# ─────────────────────────────────────────
# CONFIGURATION — Edit these values only
# ─────────────────────────────────────────

EXCEL_FILES = [
    r"File1.xlsx",
    r"File2.xlsx",
    r"File3.xlsx",
    r"File4.xlsx",
    r"File5.xlsx",
    r"File6.xlsx",
    r"File7.xlsx",
    r"File8.xlsx",
    r"File9.xlsx",
    r"File10.xlsx",
    r"File11.xlsx",
    r"File12.xlsx",
    r"File13.xlsx",
    r"File14.xlsx",
]

REFRESH_WAIT    = 10   # seconds between first and second refresh (on success)
CLOSE_WAIT      = 10   # seconds after second refresh before saving (on success)
SETTLE_WAIT     = 3    # seconds to let Excel settle right after opening
REFRESH_TIMEOUT = 300   # max seconds allowed for ONE refresh call before treating it as stuck

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


def force_kill_pid(pid):
    if not pid:
        return
    try:
        print(f"🔪 Force-killing stuck Excel process (PID {pid})...")
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True, timeout=15)
    except Exception as e:
        print(f"⚠️ Could not force-kill Excel process: {e}")


def run_with_timeout(target, timeout, *args):
    """Runs target(*args, result) in its own thread. Returns (result_dict, timed_out_bool)."""
    result = {}
    t = threading.Thread(target=target, args=(*args, result), daemon=True)
    t.start()
    t.join(timeout)
    return result, t.is_alive()


def _open(filepath, result):
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        wb = excel.Workbooks.Open(
            filepath, UpdateLinks=0, ReadOnly=False,
            IgnoreReadOnlyRecommended=True, Notify=False,
        )
        hwnd = excel.Hwnd
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        result["pid"] = pid
        result["excel_stream"] = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, excel)
        result["wb_stream"] = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, wb)
    except Exception as e:
        result["error"] = e
    finally:
        pythoncom.CoUninitialize()


def _refresh(excel_stream, wb_stream, result):
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(excel_stream, pythoncom.IID_IDispatch))
        wb = win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(wb_stream, pythoncom.IID_IDispatch))
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        result["excel_stream"] = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, excel)
        result["wb_stream"] = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, wb)
    except Exception as e:
        result["error"] = e
    finally:
        pythoncom.CoUninitialize()


def _save_close(excel_stream, wb_stream, result):
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(excel_stream, pythoncom.IID_IDispatch))
        wb = win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(wb_stream, pythoncom.IID_IDispatch))
        wb.Save()
        wb.Close(SaveChanges=True)
        excel.Quit()
        result["ok"] = True
    except Exception as e:
        result["error"] = e
    finally:
        pythoncom.CoUninitialize()


def do_open(filepath):
    print("🔄 Opening Excel...")
    result = {}
    _open(filepath, result)
    if "error" in result:
        print(f"❌ Open failed: {result['error']}")
        return None
    print("✅ File opened")
    return result


def refresh_file(filepath):
    if not os.path.exists(filepath):
        print(f"❌ Path does not exist — check the file path: {filepath}")
        return "❌ Skipped — invalid path"

    print(f"\n{'=' * 50}")
    print(f"📂 File : {filepath}")
    print(f"{'=' * 50}")

    open_res = do_open(filepath)
    if open_res is None:
        return "❌ Failed — could not open file"

    pid = open_res["pid"]
    excel_stream = open_res["excel_stream"]
    wb_stream = open_res["wb_stream"]

    print(f"⏳ Letting Excel settle {SETTLE_WAIT}s before first refresh...")
    time.sleep(SETTLE_WAIT)

    # ── First refresh ──
    print("\n🔄 First refresh...")
    res, hung = run_with_timeout(_refresh, REFRESH_TIMEOUT, excel_stream, wb_stream)
    if hung:
        print(f"⏱️ First refresh exceeded {REFRESH_TIMEOUT}s — stuck. Killing and moving to second refresh.")
        force_kill_pid(pid)
        reopened = do_open(filepath)
        if reopened is None:
            return "❌ Failed — first refresh hung and reopen failed"
        pid, excel_stream, wb_stream = reopened["pid"], reopened["excel_stream"], reopened["wb_stream"]
        time.sleep(SETTLE_WAIT)
    elif "error" in res:
        print(f"❌ First refresh errored: {res['error']} — killing and moving to second refresh.")
        force_kill_pid(pid)
        reopened = do_open(filepath)
        if reopened is None:
            return "❌ Failed — first refresh errored and reopen failed"
        pid, excel_stream, wb_stream = reopened["pid"], reopened["excel_stream"], reopened["wb_stream"]
        time.sleep(SETTLE_WAIT)
    else:
        excel_stream, wb_stream = res["excel_stream"], res["wb_stream"]
        print("✅ First refresh complete")
        print(f"⏳ Waiting {REFRESH_WAIT} seconds...")
        time.sleep(REFRESH_WAIT)

    # ── Second refresh ──
    print("\n🔄 Second refresh...")
    res2, hung2 = run_with_timeout(_refresh, REFRESH_TIMEOUT, excel_stream, wb_stream)
    lost_data = False
    if hung2:
        print(f"⏱️ Second refresh exceeded {REFRESH_TIMEOUT}s — stuck. Killing and moving to save/close.")
        force_kill_pid(pid)
        lost_data = True
    elif "error" in res2:
        print(f"❌ Second refresh errored: {res2['error']} — killing and moving to save/close.")
        force_kill_pid(pid)
        lost_data = True
    else:
        excel_stream, wb_stream = res2["excel_stream"], res2["wb_stream"]
        print("✅ Second refresh complete")
        print(f"⏳ Waiting {CLOSE_WAIT} seconds...")
        time.sleep(CLOSE_WAIT)

    if lost_data:
        print("♻️ Reopening once more to save & close cleanly (last refresh's data was lost).")
        reopened = do_open(filepath)
        if reopened is None:
            return "❌ Failed — second refresh hung and reopen for save failed"
        pid, excel_stream, wb_stream = reopened["pid"], reopened["excel_stream"], reopened["wb_stream"]

    # ── Save & close ──
    print("\n💾 Saving...")
    res3, hung3 = run_with_timeout(_save_close, REFRESH_TIMEOUT, excel_stream, wb_stream)
    if hung3:
        print(f"⏱️ Save/close exceeded {REFRESH_TIMEOUT}s — force-killing.")
        force_kill_pid(pid)
        return "⚠️ Refresh done but save/close hung — file force-closed, verify it saved"
    if "error" in res3:
        print(f"❌ Save/close failed: {res3['error']}")
        return f"❌ Failed at save/close — {res3['error']}"

    print("✅ Saved successfully")
    print("🔒 Workbook closed")
    print("🔒 Excel instance closed")
    return "✅ Success" if not lost_data else "⚠️ Saved, but a refresh attempt was stuck and skipped"


def main():
    print("=" * 50)
    print("🚀 Excel Auto Refresh Tool")
    print(f"📋 Files to process : {len(EXCEL_FILES)}")
    print("=" * 50)

    results = {}
    for i, filepath in enumerate(EXCEL_FILES, 1):
        print(f"\n[{i}/{len(EXCEL_FILES)}] Processing...")
        status = refresh_file(filepath)
        results[filepath] = status

    print(f"\n{'=' * 50}")
    print(f"📊 FINAL SUMMARY")
    print(f"{'=' * 50}")
    for filepath, status in results.items():
        filename = filepath.split("\\")[-1]
        print(f"   {status} — {filename}")
    print(f"{'=' * 50}")
    print(f"\n🎉 All done!")


if __name__ == "__main__":
    main()
