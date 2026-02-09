# Quick Start: Managing Wards

## 🎯 Quick Reference

### To Add a New Ward

1. Open `wards_config.json`
2. Copy this template at the end of the wards list:
   ```json
   ,
   {
     "code": "YOUR_CODE",
     "name": "Your Ward Name",
     "bed_complement": 20,
     "is_emergency": false,
     "display_order": 10
   }
   ```
3. Replace the values:
   - `code`: Short unique ID (e.g., "ICU", "MAT")
   - `name`: Full ward name (e.g., "Intensive Care Unit")
   - `bed_complement`: Number of beds (e.g., 15, 20, 30)
   - `is_emergency`: `true` for emergency wards, `false` for regular
   - `display_order`: Order number (higher = appears later in reports)
4. Save the file
5. Run: `python build_workbook.py --year 2026`

### To Update Ward Information

1. Open `wards_config.json`
2. Find the ward you want to update
3. Change the values (e.g., increase `bed_complement` from 20 to 25)
4. Save the file
5. Run: `python build_workbook.py --year 2026`

### To Change Ward Order in Reports

1. Open `wards_config.json`
2. Change the `display_order` numbers
3. Save the file
4. Run: `python build_workbook.py --year 2026`

## ✅ Field Quick Reference

```json
{
  "code": "MW",              // Unique ID (2-6 letters)
  "name": "Male Medical",    // Full name (any text)
  "bed_complement": 32,      // Number of beds (number)
  "is_emergency": false,     // true or false (no quotes!)
  "display_order": 1         // Order number (1, 2, 3...)
}
```

## ⚠️ Common Mistakes to Avoid

❌ **Missing comma** between ward entries
```json
{...}  // Missing comma here!
{...}
```

✅ **Correct:**
```json
{...},  // Comma added
{...}
```

---

❌ **Quotes around true/false**
```json
"is_emergency": "false"  // Wrong!
```

✅ **Correct:**
```json
"is_emergency": false  // No quotes for boolean
```

---

❌ **Trailing comma** after last ward
```json
{
  "code": "FAE",
  ...
},  // Remove this comma!
]
```

✅ **Correct:**
```json
{
  "code": "FAE",
  ...
}
]
```

## 🔧 Testing Your Changes

After editing `wards_config.json`:

```bash
# Build the workbook
python build_workbook.py --year 2026

# Check for error messages in the output
# If successful, you'll see: "Loaded X wards from wards_config.json"
```

## 📞 Need More Help?

See the full guide: **WARD_CONFIGURATION_GUIDE.md**
