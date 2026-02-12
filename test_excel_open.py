import os
import sys
import time
import win32com.client
import tempfile
import traceback

def test_excel_open(file_path):
    print(f"Testing access to: {file_path}")
    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        print(f"Error: File not found: {abs_path}")
        return

    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        print("Excel application started.")
        
        print(f"Attempting to open: {abs_path}")
        wb = excel.Workbooks.Open(abs_path)
        print("  Success! Workbook opened.")
        
        wb.Close(SaveChanges=False)
        print("  Workbook closed.")
        
    except Exception as e:
        print(f"\nERROR: {e}")
        traceback.print_exc()
    finally:
        if excel:
            try:
                excel.Quit()
                print("Excel application quit.")
            except:
                pass

def test_temp_file():
    print("Testing with temporary file...")
    # Create a temporary file
    temp_dir = tempfile.gettempdir()
    temp_file = os.path.join(temp_dir, "test_excel_auto.xlsx")
    
    # Ensure it doesn't exist
    if os.path.exists(temp_file):
        try:
            os.remove(temp_file)
        except OSError:
            print(f"Error: Could not remove existing temp file {temp_file}")
            return

    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        print("Excel application started.")
        
        # Test 1: Add a new workbook
        print("Test 1: Adding a new workbook...")
        wb = excel.Workbooks.Add()
        print("  Success! Workbook added.")
        
        # Test 2: Save it
        abs_path = os.path.abspath(temp_file)
        print(f"Test 2: Saving to {abs_path}...")
        wb.SaveAs(abs_path)
        print("  Success! Workbook saved.")
        
        wb.Close()
        print("  Workbook closed.")
        
        # Test 3: Open it back
        print(f"Test 3: Opening {abs_path}...")
        wb = excel.Workbooks.Open(abs_path)
        print("  Success! Workbook opened.")
        
        wb.Close()
        print("  Workbook closed again.")
        
    except Exception as e:
        print(f"\nERROR: {e}")
        traceback.print_exc()
        
    finally:
        if excel:
            try:
                excel.Quit()
                print("Excel application quit.")
            except:
                pass
            
        # Cleanup
        if os.path.exists(temp_file):
            try:
                os.remove(temp_file)
                print("Temp file cleaned up.")
            except:
                pass

if __name__ == "__main__":
    if len(sys.argv) > 1:
        test_excel_open(sys.argv[1])
    else:
        test_temp_file()
