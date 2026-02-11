# Ward Management System - User Guide

## 🎉 New Features

Your Bed Utilization System now has a **user-friendly ward management interface**! No more manual JSON editing!

## ✨ What's New

### 1. **Manage Wards Form** (In Excel VBA)
- Add new wards with a simple form
- Edit existing ward details
- Delete wards (when necessary)
- View all wards at a glance

### 2. **Export Ward Config Button**
- Save your ward configuration to `wards_config.json`
- Creates a backup of your settings
- Ready for next year's workbook rebuild

### 3. **Dynamic Ward Loading**
- Wards are loaded from the Control sheet table
- Changes take effect in forms immediately
- No need to edit Python code

## 📖 How to Use

### Adding a New Ward

1. **Open** `Bed_Utilization_2026.xlsm` in Excel
2. **Enable macros** when prompted
3. **Click** the **"Manage Wards"** button on the Control sheet
4. **Click** "New Ward" button in the form
5. **Fill in** the ward details:
   - **Code**: Short unique ID (e.g., "ICU", "MAT")
   - **Name**: Full ward name (e.g., "Intensive Care Unit")
   - **Bed Complement**: Number of beds (e.g., 15)
   - **Prev Year Remaining**: Patients remaining from last year (usually 0 for new wards)
   - **Emergency Ward**: Check if this is an emergency ward
   - **Display Order**: Order in reports (e.g., 10, 11, 12...)
6. **Click** "Save"
7. **Click** "Export Config to JSON" to save the configuration
8. **Rebuild** the workbook for changes to take full effect:
   ```bash
   python build_workbook.py --year 2026
   ```

### Editing a Ward

1. **Open** the workbook and **click** "Manage Wards"
2. **Select** the ward from the list on the left
3. **Edit** the details on the right
4. **Click** "Save"
5. **Click** "Export Config to JSON"
6. **Rebuild** if you changed ward codes or added/removed wards

### Deleting a Ward

1. **Open** the workbook and **click** "Manage Wards"
2. **Select** the ward from the list
3. **Click** "Delete"
4. **Confirm** the deletion
5. **Click** "Export Config to JSON"
6. **Rebuild** the workbook to remove the ward sheet

### Exporting Configuration

At any time, you can export your current ward configuration:

1. **Click** "Export Ward Config" button on the Control sheet
2. The file `wards_config.json` will be updated
3. You'll see a confirmation message

## 🔄 Complete Workflow

Here's the complete workflow for managing wards:

```
┌─────────────────────────────────────┐
│ 1. Open Bed_Utilization_2026.xlsm  │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│ 2. Click "Manage Wards" button      │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│ 3. Add/Edit/Delete wards in form    │
│    - Click "New Ward" to add        │
│    - Select ward to edit            │
│    - Select + "Delete" to remove    │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│ 4. Click "Export Config to JSON"    │
│    (in the form or main Control)    │
└──────────────┬──────────────────────┘
               │
               ▼
┌─────────────────────────────────────┐
│ 5. Rebuild workbook (if needed)     │
│    python build_workbook.py --year  │
│    2026                              │
└─────────────────────────────────────┘
```

## 🎯 When to Rebuild

You **MUST rebuild** the workbook when:
- ✅ Adding a new ward (to create the ward sheet)
- ✅ Deleting a ward (to remove the ward sheet)
- ✅ Changing a ward code (affects formulas)

You **DON'T need to rebuild** when:
- ⏭️ Changing ward name
- ⏭️ Updating bed complement
- ⏭️ Updating previous year remaining
- ⏭️ Changing emergency status
- ⏭️ Changing display order

## 💡 Pro Tips

1. **Export Often**: Click "Export Config to JSON" after making changes to create a backup

2. **Test Before Rollout**: Make changes in a test workbook first

3. **Consistent Naming**: Use clear, descriptive ward names

4. **Logical Order**: Assign display orders that group related wards together
   - Example: Regular wards (1-6), Special units (7-9), Emergency (10-11)

5. **Backup**: Keep a copy of `wards_config.json` before making major changes

## ⚠️ Important Notes

- **Ward codes must be unique** - you can't have two wards with the same code
- **Display order** affects how wards appear in all reports
- **Deleting a ward** doesn't delete historical data - it just removes the ward from the configuration
- **After rebuilding**, you may need to set "Previous Year Remaining" values again for new wards

## 🔧 Troubleshooting

### "Ward code already exists"
- Each ward must have a unique code
- Try a different code (e.g., "ICU2", "MAT1")

### "Wards not showing in reports"
- Make sure you clicked "Export Config to JSON"
- Rebuild the workbook
- Check that the ward is in the Control sheet table

### "Changes not taking effect"
- Some changes require rebuilding (see "When to Rebuild" above)
- Make sure you exported the configuration
- Close and reopen Excel if forms aren't refreshing

## 📞 Need Help?

See also:
- [WARD_CONFIGURATION_GUIDE.md](WARD_CONFIGURATION_GUIDE.md) - Detailed configuration guide
- [QUICK_START_WARDS.md](QUICK_START_WARDS.md) - Quick reference
- [README.md](README.md) - General system documentation

## 🎊 Benefits

✅ **No JSON editing** - Use user-friendly forms instead
✅ **Immediate feedback** - See changes right away in the form
✅ **Error prevention** - Built-in validation prevents mistakes
✅ **Easy backup** - Export to JSON anytime
✅ **Flexible** - Can still edit JSON manually if you prefer

---

**Enjoy your new ward management system!** 🎉
