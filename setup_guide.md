# 🛠️ Setup Guide — CBSE Academic Planner 2026-2027

**Designed by Principal Haidar Ali Shaikh**

---

## Option 1: Use the Excel File Directly (Recommended)

### Step 1: Download the Excel File

1. Go to the repository: [Haidaralishaikh1/EXCEL](https://github.com/Haidaralishaikh1/EXCEL)
2. Click on `CBSE_Academic_Planner_2026_2027.xlsx`
3. Click **Download** (or the download icon)
4. Save to your computer

### Step 2: Open the File

- **Microsoft Excel**: Double-click the file — it opens directly
- **Google Sheets**: Go to [sheets.google.com](https://sheets.google.com) → File → Import → Upload the .xlsx file
- **LibreOffice Calc**: File → Open → select the file

### Step 3: Navigate the Sheets

The workbook has **16 sheets** (tabs at the bottom):

| Tab | What It Contains |
|-----|-----------------|
| `INSTRUCTIONS` | How-to guide — read first! |
| `MASTER INPUT` | Enter/edit lesson data here |
| `HOLIDAYS` | Pre-loaded holidays for 2026-2027 |
| `APR 2026` ... `MAR 2027` | Monthly calendars (12 sheets) |
| `DASHBOARD` | Summary and statistics |

---

## Option 2: Run the Python Script

Use this if you want to customize lessons or regenerate with updated data.

### Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

### Step 1: Install Python (if not already installed)

Download from [python.org](https://www.python.org/downloads/)

### Step 2: Install the Required Library

```bash
pip install openpyxl>=3.1.0
```

Or use the requirements file:

```bash
pip install -r requirements.txt
```

### Step 3: Run the Script

```bash
python generate_planner.py
```

The script will print progress and save `CBSE_Academic_Planner_2026_2027.xlsx` in the same folder.

### Step 4: Open the Generated File

Open `CBSE_Academic_Planner_2026_2027.xlsx` in Excel or Google Sheets.

---

## Option 3: Google Apps Script (Google Sheets)

Use this if you prefer working entirely in Google Sheets.

### Step 1: Create a New Google Spreadsheet

1. Go to [sheets.google.com](https://sheets.google.com)
2. Click **+ Blank** to create a new spreadsheet
3. Name it: `CBSE Academic Planner 2026-2027`

### Step 2: Open Apps Script Editor

1. Click **Extensions** in the menu bar
2. Click **Apps Script**

### Step 3: Paste the Script

1. In the Apps Script editor, select all existing code (Ctrl+A) and delete it
2. Copy the entire contents of `Academic_Planner_Setup.gs`
3. Paste it into the editor
4. Click **Save** (Ctrl+S)

### Step 4: Run the Generator

1. In the Apps Script editor, click **Run → generateAcademicPlanner**
2. **Grant permissions** when the browser asks (click "Allow")
3. Wait 1-2 minutes — all sheets are being generated

### Step 5: Use the Custom Menu

After running once, a **"🏫 Academic Planner"** menu appears in your Google Sheet:

- **▶️ Generate Full Planner** — Regenerates all sheets from scratch
- **🔄 Refresh Dashboard** — Updates summary statistics
- **📅 Update Monthly Sheets** — Refreshes all 12 calendar sheets
- **🎉 Update Holidays** — Updates the holidays sheet

---

## How to Modify Lessons in MASTER INPUT

1. Open the **MASTER INPUT** sheet
2. Find the row for the subject/lesson you want to change
3. Edit any of these columns:
   - **Column A**: Subject name
   - **Column C**: Lesson name
   - **Columns D–P**: Number of days for each grade (use 0 if not applicable)
   - **Column Q**: Start Month (select from dropdown)
   - **Column R**: Priority/Sequence number

4. After editing:
   - **For Excel**: The monthly sheets will need to be manually updated (or re-run the Python script)
   - **For Google Sheets**: Use **🏫 Academic Planner → 📅 Update Monthly Sheets** from the menu

### Tips for Editing

- Use **0** in a grade column if that lesson is not taught for that grade
- Keep **Priority** numbers sequential within each subject-month combination
- The **Start Month** determines which monthly calendar sheet the lesson appears in
- Days should be realistic:
  - Nursery/LKG/UKG: 4–7 days per lesson
  - Grade 1–3: 6–10 days per lesson
  - Grade 4–5: 8–12 days per lesson
  - Grade 6–8: 10–14 days per lesson
  - Grade 9–10: 12–16 days per lesson

---

## How to Add a New Holiday

1. Open the **HOLIDAYS** sheet
2. Scroll to the bottom of the existing data
3. Add a new row with:
   - **Date**: The holiday date
   - **Day**: Day of week (Monday, Tuesday, etc.)
   - **Holiday Name**: Name of the holiday
   - **Type**: National / State (Maharashtra) / School / Optional / Vacation
   - **Applicable To**: All
4. Re-run the Python script to regenerate monthly sheets with the new holiday

---

## Troubleshooting

| Issue | Solution |
|-------|---------|
| "Module not found: openpyxl" | Run `pip install openpyxl` |
| Excel file won't open | Ensure you have Excel 2016+ or latest LibreOffice |
| Google Script permission denied | Run as yourself (not in incognito mode) |
| Monthly sheets not showing lessons | Check that Start Month in MASTER INPUT matches exactly |
| Script takes too long | Google Apps Script has a 6-minute timeout; run in sections |

---

*🏫 CBSE Academic Planner 2026-2027 — Designed by Principal Haidar Ali Shaikh*
