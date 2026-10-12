"""System optimization script for Windows PC:
1. Cleans stale PyInstaller _MEI* extraction folders in %TEMP%.
2. Disables non-essential auto-start background apps (Steam, RiotClient, qBittorrent, Edge auto-launch) in HKCU Run.
3. Creates a fast Desktop shortcut to dist/LOA-Monitoring-Tool/LOA-Monitoring-Tool.exe.
"""

import os
import shutil
import glob
import subprocess
import winreg

def clean_temp_mei_folders():
    temp_dir = os.environ.get("TEMP", "")
    if not temp_dir or not os.path.isdir(temp_dir):
        print("[-] Temp directory not found.")
        return 0

    mei_folders = glob.glob(os.path.join(temp_dir, "_MEI*"))
    deleted = 0
    freed_bytes = 0
    print(f"[*] Found {len(mei_folders)} _MEI* temporary folders in {temp_dir}.")

    for folder in mei_folders:
        try:
            for root, dirs, files in os.walk(folder):
                for f in files:
                    try:
                        freed_bytes += os.path.getsize(os.path.join(root, f))
                    except Exception:
                        pass
            shutil.rmtree(folder, ignore_errors=True)
            deleted += 1
        except Exception:
            pass

    mb_freed = round(freed_bytes / (1024 * 1024), 1)
    print(f"[+] Cleaned {deleted} temporary extraction folders (~{mb_freed} MB freed).")
    return deleted

def optimize_startup_apps():
    """Disable non-essential gaming/torrent/browser background startup keys in HKCU Run."""
    run_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    
    targets_to_disable = [
        "qBittorrent",
        "RiotClient",
        "Steam",
        "MicrosoftEdgeAutoLaunch_19FCAA5C6E54B7FB05BE845C7B4EF005",
        "MicrosoftEdgeAutoLaunch_14AC5A18FCC91E8F2B594B461773C951"
    ]

    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_path, 0, winreg.KEY_ALL_ACCESS)
    except Exception as e:
        print(f"[-] Could not open HKCU Run key: {e}")
        return

    num_values = winreg.QueryInfoKey(key)[1]
    existing_items = {}
    for i in range(num_values):
        name, val, _ = winreg.EnumValue(key, i)
        existing_items[name] = val

    print(f"[*] Checking {len(existing_items)} user startup entries...")

    disabled_count = 0
    for name in targets_to_disable:
        if name in existing_items:
            try:
                winreg.DeleteValue(key, name)
                print(f"[+] Disabled startup item: {name}")
                disabled_count += 1
            except Exception as e:
                print(f"[-] Could not disable {name}: {e}")

    winreg.CloseKey(key)
    print(f"[+] Successfully optimized startup: disabled {disabled_count} background launch items.")

def create_desktop_shortcut(target_exe, icon_path=None):
    """Creates a Windows desktop shortcut (.lnk) pointing to the fast executable."""
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.isdir(desktop):
        return

    shortcut_path = os.path.join(desktop, "LOA Monitoring Tool.lnk")
    target_exe = os.path.abspath(target_exe)
    working_dir = os.path.dirname(target_exe)
    icon_loc = os.path.abspath(icon_path) if icon_path and os.path.isfile(icon_path) else target_exe

    ps_script = f"""
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("{shortcut_path}")
    $Shortcut.TargetPath = "{target_exe}"
    $Shortcut.WorkingDirectory = "{working_dir}"
    $Shortcut.IconLocation = "{icon_loc}"
    $Shortcut.Description = "LOA Monitoring Tool - Rapid Local Launch"
    $Shortcut.Save()
    """

    res = subprocess.run(["powershell", "-NoProfile", "-Command", ps_script], capture_output=True, text=True)
    if os.path.isfile(shortcut_path):
        print(f"[+] Created instant desktop shortcut: {shortcut_path}")
    else:
        print(f"[-] Could not create desktop shortcut: {res.stderr}")

if __name__ == "__main__":
    print("=== PC & APPLICATION STARTUP OPTIMIZER ===")
    clean_temp_mei_folders()
    optimize_startup_apps()
    
    target = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "dist", "LOA-Monitoring-Tool", "LOA-Monitoring-Tool.exe")
    icon = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "icon.ico")
    if os.path.isfile(target):
        create_desktop_shortcut(target, icon)
