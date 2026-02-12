import openpyxl
import win32com.client
import os
import time

def test_openpyxl_com_compat():
    print("Testing openpyxl -> COM compatibility...")
    
    filename = "test_compat.xlsx"
    abs_path = os.path.abspath(filename)
    
    # 1. Create simple file with openpyxl
    print("1. Creating file with openpyxl...")
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = "Hello World"
        wb.save(filename)
        print("   Saved successfully.")
    except Exception as e:
        print(f"   Failed to save with openpyxl: {e}")
        return

    # 2. Try to open with COM
    print(f"2. Opening with Excel COM: {abs_path}")
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        
        wb_com = excel.Workbooks.Open(abs_path)
        print("   Success! COM opened the openpyxl file.")
        
        wb_com.Close(SaveChanges=False)
        
    except Exception as e:
        print(f"   ERROR: Failed to open with COM: {e}")
        
    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass
        
        # Cleanup
        if os.path.exists(filename):
            try:
                os.remove(filename)
                print("   Cleanup successful.")
            except:
                pass

if __name__ == "__main__":
    test_openpyxl_com_compat()
