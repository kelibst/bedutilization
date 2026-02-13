"""
Helper script to fix common Excel COM automation issues
Checks for and resolves:
1. Zombie Excel processes
2. Files locked by Excel
3. Temporary file cleanup
"""
import os
import sys
import time

def kill_excel_processes():
    """Kill all Excel processes"""
    try:
        import psutil
        killed = 0
        for proc in psutil.process_iter(['name', 'pid']):
            if proc.info['name'] and 'excel' in proc.info['name'].lower():
                try:
                    print(f"Killing Excel process (PID: {proc.info['pid']})")
                    proc.kill()
                    killed += 1
                except (psutil.AccessDenied, psutil.NoSuchProcess):
                    pass
        
        if killed > 0:
            print(f"✓ Killed {killed} Excel process(es)")
            time.sleep(2)  # Wait for processes to fully terminate
        else:
            print("No Excel processes found")
        
        return killed
    except ImportError:
        print("ERROR: psutil not installed. Run: pip install psutil")
        return 0

def check_file_locks(file_path):
    """Check if a file is locked by any process"""
    try:
        import psutil
        abs_path = os.path.abspath(file_path)
        
        for proc in psutil.process_iter(['name', 'open_files']):
            try:
                if proc.info['open_files']:
                    for file in proc.info['open_files']:
                        if abs_path in file.path:
                            print(f"⚠️  File is locked by: {proc.info['name']} (PID: {proc.pid})")
                            return True
            except (psutil.AccessDenied, psutil.NoSuchProcess):
                pass
        
        print(f"✓ File is not locked: {file_path}")
        return False
    except ImportError:
        print("ERROR: psutil not installed. Run: pip install psutil")
        return False

def cleanup_temp_files(directory="."):
    """Remove temporary Excel files"""
    temp_patterns = ["~$*.xlsx", "~$*.xlsm", "*.tmp"]
    removed = 0
    
    for pattern in temp_patterns:
        import glob
        for file in glob.glob(os.path.join(directory, pattern)):
            try:
                os.remove(file)
                print(f"Removed temp file: {file}")
                removed += 1
            except Exception as e:
                print(f"Could not remove {file}: {e}")
    
    if removed > 0:
        print(f"✓ Removed {removed} temporary file(s)")
    else:
        print("No temporary files found")
    
    return removed

def main():
    print("=" * 60)
    print("  Excel COM Automation Troubleshooter")
    print("=" * 60)
    print()
    
    # Check for specific file if provided
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        if os.path.exists(file_path):
            print(f"Checking file: {file_path}\n")
            check_file_locks(file_path)
            print()
    
    # Kill Excel processes
    print("Step 1: Checking for Excel processes...")
    kill_excel_processes()
    print()
    
    # Cleanup temp files
    print("Step 2: Cleaning up temporary files...")
    cleanup_temp_files()
    print()
    
    print("=" * 60)
    print("  Troubleshooting complete!")
    print("  You can now try running build_workbook.py again")
    print("=" * 60)

if __name__ == "__main__":
    main()
