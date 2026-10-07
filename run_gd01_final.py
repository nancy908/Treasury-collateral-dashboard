import os
import time
import csv
import datetime
import shutil
import ctypes
import ctypes.wintypes as wintypes
import webbrowser

# Win32 structures for CreateProcessW on interactive desktop
class PROCESS_INFORMATION(ctypes.Structure):
    _fields_ = [
        ("hProcess", wintypes.HANDLE),
        ("hThread", wintypes.HANDLE),
        ("dwProcessId", wintypes.DWORD),
        ("dwThreadId", wintypes.DWORD)
    ]

class STARTUPINFOW(ctypes.Structure):
    _fields_ = [
        ("cb", wintypes.DWORD),
        ("lpReserved", wintypes.LPWSTR),
        ("lpDesktop", wintypes.LPWSTR),
        ("lpTitle", wintypes.LPWSTR),
        ("dwX", wintypes.DWORD),
        ("dwY", wintypes.DWORD),
        ("dwXSize", wintypes.DWORD),
        ("dwYSize", wintypes.DWORD),
        ("dwXCountChars", wintypes.DWORD),
        ("dwYCountChars", wintypes.DWORD),
        ("dwFillAttribute", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("wShowWindow", wintypes.WORD),
        ("cbReserved2", wintypes.WORD),
        ("lpReserved2", ctypes.POINTER(ctypes.c_byte)),
        ("hStdInput", wintypes.HANDLE),
        ("hStdOutput", wintypes.HANDLE),
        ("hStdError", wintypes.HANDLE)
    ]

STARTF_USESHOWWINDOW = 0x00000001
SW_SHOW = 5

downloads_dir = os.path.expanduser(r"~\Downloads")
chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
today = datetime.date.today().strftime("%Y-%m-%d")
dest_dir = r"G:\My Drive\Nancy\Financial Reports\0 Cash flow (New)\3 Google Antigravity\Guarantee dashboard\Data"

# Ensure destination directory exists
os.makedirs(dest_dir, exist_ok=True)

def download_via_desktop_chrome(url, timeout=60):
    start_time = time.time()
    # Get snapshot of downloads dir with mtimes
    initial_mtimes = {}
    for f in os.listdir(downloads_dir):
        if f.lower().endswith(".csv"):
            try:
                initial_mtimes[f] = os.path.getmtime(os.path.join(downloads_dir, f))
            except Exception:
                pass
    
    # Open URL with Chrome on interactive desktop (keeps user session cookies)
    print(f"Opening download URL: {url}")
    
    si = STARTUPINFOW()
    si.cb = ctypes.sizeof(STARTUPINFOW)
    si.lpDesktop = "winsta0\\default"
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = SW_SHOW
    
    pi = PROCESS_INFORMATION()
    
    chrome_profile = "Profile 7"
    cmd = f'"{chrome_path}" --profile-directory="{chrome_profile}" "{url}"'
    
    CreateProcessW = ctypes.windll.kernel32.CreateProcessW
    CreateProcessW.argtypes = [
        wintypes.LPCWSTR, wintypes.LPWSTR, ctypes.c_void_p, ctypes.c_void_p,
        wintypes.BOOL, wintypes.DWORD, ctypes.c_void_p, wintypes.LPCWSTR,
        ctypes.POINTER(STARTUPINFOW), ctypes.POINTER(PROCESS_INFORMATION)
    ]
    CreateProcessW.restype = wintypes.BOOL
    
    success = CreateProcessW(
        None,
        cmd,
        None,
        None,
        False,
        0,
        None,
        None,
        ctypes.byref(si),
        ctypes.byref(pi)
    )
    
    if success:
        ctypes.windll.kernel32.CloseHandle(pi.hProcess)
        ctypes.windll.kernel32.CloseHandle(pi.hThread)
    else:
        print(f"Interactive launch failed (Error: {ctypes.windll.kernel32.GetLastError()}). Falling back to shell execution...")
        import subprocess
        subprocess.run(f'start chrome --profile-directory="{chrome_profile}" "{url}"', shell=True)
    
    # Monitor downloads
    downloaded_file = None
    while time.time() - start_time < timeout:
        time.sleep(1)
        for f in os.listdir(downloads_dir):
            if not f.lower().endswith(".csv"):
                continue
            path = os.path.join(downloads_dir, f)
            crdownload_path = path + ".crdownload"
            tmp_path = path + ".tmp"
            
            try:
                current_mtime = os.path.getmtime(path)
            except Exception:
                continue
                
            is_new = (f not in initial_mtimes)
            is_updated = (f in initial_mtimes and current_mtime > initial_mtimes[f] + 0.01)
            
            if (is_new or is_updated) and not os.path.exists(crdownload_path) and not os.path.exists(tmp_path):
                try:
                    # check if readable (not locked by browser anymore)
                    with open(path, 'rb') as fp:
                        pass
                    downloaded_file = path
                    break
                except IOError:
                    pass
        if downloaded_file:
            break
            
    if not downloaded_file:
        raise TimeoutError(f"Download timed out or failed to detect new CSV file for URL: {url}")
        
    print(f"Detected downloaded file: {downloaded_file}")
    return downloaded_file

def run_gd01():
    print("=============================================================")
    print("STARTING GD01 DATA INGESTION PIPELINE")
    print(f"Date: {today}")
    print("=============================================================\n")
    
    # --- Task 1: Summary Sheet ---
    print("[1/5] Downloading Summary sheet...")
    url1 = "https://docs.google.com/spreadsheets/d/1fGrRYCJ5mCMeA7hyqAFV5M5dYiJB95btYyrN1ZKSJiY/export?format=csv&gid=550601387"
    file1 = download_via_desktop_chrome(url1)
    
    dest_file1 = os.path.join(dest_dir, f"Collateral calculator 2 - Summary {today}.csv")
    # Move the file, handling overwrite if it exists
    if os.path.exists(dest_file1):
        os.remove(dest_file1)
    shutil.move(file1, dest_file1)
    print(f"Saved: {dest_file1}")
    fixed_file1 = os.path.join(dest_dir, "Collateral calculator 2 - Summary.csv")
    shutil.copy2(dest_file1, fixed_file1)
    print(f"Saved: {fixed_file1}\n")
    
    # --- Task 2: CFF Treasury (O43:U57) ---
    print("[2/5] Downloading CFF Treasury sheet...")
    url2 = "https://docs.google.com/spreadsheets/d/1b333pTDQYcoo6wjYcLSOPW3q5E_xyqdqkfD41433PIE/export?format=csv&gid=1275658334"
    file2 = download_via_desktop_chrome(url2)
    
    print("Parsing and slicing CFF Treasury (O43:U57)...")
    sliced_rows2 = []
    # O43 is row 43 (index 42), col O is 15th (index 14)
    # U57 is row 57 (index 56), col U is 21st (index 20)
    with open(file2, "r", encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f)
        rows = list(reader)
        for r in range(42, 57):
            if r < len(rows):
                row = rows[r]
                sliced_row = []
                for c in range(14, 21):
                    if c < len(row):
                        sliced_row.append(row[c])
                    else:
                        sliced_row.append("")
                sliced_rows2.append(sliced_row)
                
    dest_file2 = os.path.join(dest_dir, f"CFF Treasury {today}.csv")
    with open(dest_file2, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(sliced_rows2)
        
    os.remove(file2)
    print(f"Saved: {dest_file2}")
    fixed_file2 = os.path.join(dest_dir, "CFF Treasury.csv")
    shutil.copy2(dest_file2, fixed_file2)
    print(f"Saved: {fixed_file2}\n")
    
    # --- Task 3: CFF FP&A (Columns L to Q) ---
    print("[3/5] Downloading CFF FP&A sheet...")
    url3 = "https://docs.google.com/spreadsheets/d/1b333pTDQYcoo6wjYcLSOPW3q5E_xyqdqkfD41433PIE/export?format=csv&gid=1997786154"
    file3 = download_via_desktop_chrome(url3)
    
    print("Parsing and keeping CFF FP&A columns (Period to WCFF)...")
    sliced_rows3 = []
    with open(file3, "r", encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f)
        all_rows = list(reader)

        start_col = 11
        end_col = 21  # default indices 11 to 20 (cols L to U)
        for r in all_rows[:10]:
            if len(r) > 11 and r[11].strip().lower() == 'period':
                for c in range(start_col, len(r)):
                    val = r[c].strip().lower()
                    if not val or 'consumption' in val:
                        end_col = c
                        break
                else:
                    end_col = len(r)
                break

        for row in all_rows:
            sliced_row = []
            for col_idx in range(start_col, end_col):
                if col_idx < len(row):
                    sliced_row.append(row[col_idx])
                else:
                    sliced_row.append("")
            sliced_rows3.append(sliced_row)
            
    dest_file3 = os.path.join(dest_dir, f"CFF FP&A {today}.csv")
    with open(dest_file3, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(sliced_rows3)
        
    os.remove(file3)
    print(f"Saved: {dest_file3}")
    fixed_file3 = os.path.join(dest_dir, "CFF FP&A.csv")
    shutil.copy2(dest_file3, fixed_file3)
    print(f"Saved: {fixed_file3}\n")
    
    # --- Task 4: Covenant Sheet ---
    print("[4/5] Downloading Covenant sheet...")
    url4 = "https://docs.google.com/spreadsheets/d/1fGrRYCJ5mCMeA7hyqAFV5M5dYiJB95btYyrN1ZKSJiY/export?format=csv&gid=519116704"
    file4 = download_via_desktop_chrome(url4)
    
    dest_file4 = os.path.join(dest_dir, f"Covenant forecast {today}.csv")
    if os.path.exists(dest_file4):
        os.remove(dest_file4)
    shutil.move(file4, dest_file4)
    print(f"Saved: {dest_file4}")
    fixed_file4 = os.path.join(dest_dir, "Covenant forecast.csv")
    shutil.copy2(dest_file4, fixed_file4)
    print(f"Saved: {fixed_file4}\n")
    
    # --- Task 5: Cash Position History Sheet ---
    print("[5/5] Downloading Cash position history sheet...")
    url5 = "https://docs.google.com/spreadsheets/d/1b333pTDQYcoo6wjYcLSOPW3q5E_xyqdqkfD41433PIE/export?format=csv&gid=71255572"
    file5 = download_via_desktop_chrome(url5)
    
    dest_file5 = os.path.join(dest_dir, f"Cash position history {today}.csv")
    if os.path.exists(dest_file5):
        os.remove(dest_file5)
    shutil.move(file5, dest_file5)
    print(f"Saved: {dest_file5}")
    fixed_file5 = os.path.join(dest_dir, "Cash position history.csv")
    shutil.copy2(dest_file5, fixed_file5)
    print(f"Saved: {fixed_file5}\n")
    
    print("=============================================================")
    print("GD01 DOWNLOADS COMPLETED SUCCESSFULLY!")
    print("=============================================================")

    # Automatically update dashboard after downloads
    try:
        import update_dashboard_final
        print("\nUpdating dashboard HTML, PPTX, and syncing...")
        update_dashboard_final.main()
    except Exception as e:
        print(f"Warning: Failed to auto-update dashboard: {e}")

if __name__ == "__main__":
    run_gd01()
