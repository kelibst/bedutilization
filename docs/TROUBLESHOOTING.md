# Troubleshooting Excel VBA Injection Errors

## Quick Fix Steps

### Step 1: Close All Excel Windows
The most common cause is that Excel has the file open or there are zombie Excel processes.

**Run this command:**
```bash
python kill_excel.py
```

### Step 2: Delete the Generated .xlsx File
The file might be corrupted. Delete it and regenerate:
```bash
del Bed_Utilization_2026.xlsx
```

### Step 3: Run the Build Script Again
```bash
python build_workbook.py
```

---

## Alternative: Build Without VBA First

If the issue persists, you can build the workbook without VBA macros first, then add them manually:

```bash
python build_workbook.py --skip-vba
```

This will create a `.xlsx` file without macros that you can open and verify.

---

## Common Causes & Solutions

### 1. File Already Open in Excel
**Symptom:** "Open method of Workbooks class failed"

**Solution:**
- Close all Excel windows
- Run `python kill_excel.py` to kill zombie processes
- Try again

### 2. VBA Project Access Not Enabled
**Symptom:** "Programmatic access is not trusted"

**Solution:**
1. Open Excel
2. Go to: File → Options → Trust Center → Trust Center Settings
3. Click "Macro Settings"
4. Check "Trust access to the VBA project object model"
5. Click OK and restart Excel

### 3. File Corruption
**Symptom:** File exists but Excel can't open it

**Solution:**
- Delete the `.xlsx` file
- Run `python build_workbook.py` again

### 4. Antivirus Blocking
**Symptom:** Random failures, file access denied

**Solution:**
- Temporarily disable antivirus
- Add the project folder to antivirus exclusions
- Try again

---

## Manual VBA Addition (If All Else Fails)

If automated VBA injection continues to fail:

1. Run: `python build_workbook.py --skip-vba`
2. Open the generated `.xlsx` file in Excel
3. Save it as `.xlsm` (Excel Macro-Enabled Workbook)
4. Press Alt+F11 to open VBA Editor
5. Manually add the VBA code from the `phase2_vba.py` file

---

## Testing the Fix

After applying fixes, test with:
```bash
python build_workbook.py
```

You should see:
```
--- Phase 1: Building workbook structure ---
Phase 1 complete: ...\Bed_Utilization_2026.xlsx

--- Phase 2: Injecting VBA macros ---
Performing pre-flight checks...
Starting Excel for VBA injection...
Opening workbook: ...
  ✓ Workbook opened successfully
  Injecting VBA modules...
  ...
Phase 2 complete: ...\Bed_Utilization_2026.xlsm
```
