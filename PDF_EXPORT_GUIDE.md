# PDF Export Feature - Quick Guide

## ✅ What's New

The GUI now has **TWO TABS**:

### Tab 1: Generate Certificates 📝
(Same as before)
- Select calibration file
- Generate Excel certificates

### Tab 2: Export to PDF 📄 (NEW!)
- Select certificate Excel file
- Choose which sheets to export (one, multiple, or all)
- Export to PDF with full formatting, logos, and signatures preserved

---

## 🎯 How to Use PDF Export

### Step 1: Generate Certificates First
1. Go to "**Generate Certificates**" tab
2. Select calibration file
3. Click "Generate Certificates"
4. Wait for completion

### Step 2: Export to PDF
1. Go to "**Export to PDF**" tab
2. Click **Browse** and select your certificate Excel file (e.g., CYBER_PARK_TOWER_B_complete.xlsx)
3. The listbox will show all available sheets
4. Select sheets to export:
   - **Click individual sheets** to select them
   - **Click "Select All"** to export all certificates
   - **Hold Ctrl+Click** to select multiple specific sheets
5. Choose output folder (default: PDF_Certificates)
6. Click **"Export to PDF"**

---

## 📊 Features

### Selection Options:
- ✅ **One sheet**: Click a single sheet name
- ✅ **Multiple sheets**: Ctrl+Click multiple sheet names
- ✅ **All sheets**: Click "Select All" button

### Output:
- Each sheet becomes a separate PDF file
- PDF filename matches sheet name
- All formatting preserved (logo, signatures, borders, fonts)
- Saved to selected output folder

---

## 📁 Example Workflow

### Exporting Tower B Certificates:
```
1. Tab 2: Export to PDF
2. Browse → Select "CYBER_PARK_TOWER_B_complete.xlsx"
3. Listbox shows all 48 sheets (TowerB_xxx)
4. Click "Select All" (or select specific ones)
5. Output Folder: C:\Users\sumit\Downloads\manish\PDF_Certificates
6. Click "Export to PDF"
7. Wait for progress bar
8. Done! 48 PDF files created
```

### Exporting Specific Certificates:
```
1. Browse → Select certificate file
2. Ctrl+Click to select specific sheets:
   - TowerB_ADMIN_OFFICE
   - TowerB_CAFETERIA
   - TowerB_SERVER_ROOM
3. Click "Export to PDF"
4. Only these 3 certificates exported as PDF
```

---

## 🔧 Technical Details

### Requirements:
- ✅ **pywin32** - Installed automatically
- ✅ **Microsoft Excel** - Must be installed on your computer
- ✅ Windows OS

### PDF Export Process:
1. Uses Excel's native PDF export (via COM automation)
2. Preserves all formatting exactly as Excel shows
3. Includes logos, images, signatures
4. Each sheet exported as separate PDF file

### Performance:
- ~2-3 seconds per certificate
- Progress bar shows real-time status
- Runs in background (GUI stays responsive)

---

## 💡 Tips

### For Best Results:
1. **Close Excel** before exporting (if the file is open)
2. **Select output folder** with enough disk space
3. **Use "Select All"** for batch export of entire tower
4. **Check output folder** after export completes

### Common Use Cases:

**Export all certificates for one tower:**
```
Select file → Click "Select All" → Export to PDF
```

**Export specific meters only:**
```
Select file → Ctrl+Click specific sheets → Export to PDF
```

**Export to custom location:**
```
Click "Browse" next to Output Folder → Choose location → Export
```

---

## 🎯 Quick Start

### Launch GUI:
```bash
.venv\Scripts\python.exe gui_certificate_generator.py
```

### Generate + Export in One Go:
1. **Tab 1**: Generate certificates from calibration file
2. **Tab 2**: Export all to PDF
3. Done!

---

## 📝 Notes

- PDF files preserve **exact Excel formatting**
- Logos and signatures included automatically
- Each PDF named after sheet (e.g., "TowerB_ADMIN_OFFICE.pdf")
- Output folder created automatically if doesn't exist
- Can export same file multiple times (overwrites existing PDFs)

---

## ✅ Already Available Certificate Files

You can export these immediately:
- CYBER_PARK_TOWER_B_complete.xlsx (48 sheets)
- CYBER_PARK_TOWER_C_complete.xlsx (10 sheets)
- CYBER_PARK_GROUND_FLOOR_SHOP_complete.xlsx (10 sheets)
- CYBER_PARK_BASEMENT_complete.xlsx (7 sheets)

**Total: 75 certificates ready for PDF export!**
