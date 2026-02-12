"""
Simple Excel Process Killer (No dependencies required)
Kills all Excel processes to resolve COM automation issues
"""
import subprocess
import sys
import time

def kill_excel_windows():
    """Kill all Excel processes using Windows taskkill"""
    print("=" * 60)
    print("  Excel Process Killer")
    print("=" * 60)
    print()
    
    print("Checking for Excel processes...")
    
    # Use tasklist to check for Excel
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE"],
            capture_output=True,
            text=True
        )
        
        if "EXCEL.EXE" in result.stdout:
            print("Found Excel processes. Killing them...")
            
            # Kill all Excel processes
            kill_result = subprocess.run(
                ["taskkill", "/F", "/IM", "EXCEL.EXE"],
                capture_output=True,
                text=True
            )
            
            if kill_result.returncode == 0:
                print("✓ Successfully killed Excel processes")
                time.sleep(2)  # Wait for processes to fully terminate
            else:
                print(f"⚠️  Warning: {kill_result.stderr}")
        else:
            print("✓ No Excel processes found")
    
    except Exception as e:
        print(f"Error: {e}")
        return False
    
    print()
    print("=" * 60)
    print("  Done! You can now try running build_workbook.py again")
    print("=" * 60)
    return True

if __name__ == "__main__":
    kill_excel_windows()
