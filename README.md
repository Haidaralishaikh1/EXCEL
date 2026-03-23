# 🏫 CBSE Academic Planner 2026-2027

**Designed by Principal Haidar Ali Shaikh**

A complete, professional academic planning system for CBSE schools — Nursery to Grade 10.
Covers the full academic year **April 2026 to March 2027** with all Indian national and Maharashtra state holidays pre-loaded, lesson scheduling across working days, and monthly calendars for all 13 grades.

---

## 📋 Table of Contents
1. [Project Overview](#project-overview)
2. [Files in This Repository](#files-in-this-repository)
3. [How to Use the Excel File](#how-to-use-the-excel-file)
4. [How to Run the Python Script](#how-to-run-the-python-script)
5. [How to Use the Google Apps Script](#how-to-use-the-google-apps-script)
6. [Holidays Included](#holidays-included)
7. [CBSE Academic Planning Notes](#cbse-academic-planning-notes)

---

## 📌 Project Overview

This planner generates a **16-sheet Excel workbook** (.xlsx) with:
- **INSTRUCTIONS** — Step-by-step usage guide
- **MASTER INPUT** — Enter lesson names and days per grade
- **HOLIDAYS** — All 30+ pre-loaded holidays for 2026-2027
- **APR 2026 → MAR 2027** — 12 monthly calendar sheets with lesson assignments
- **DASHBOARD** — Grade-wise summary and month-wise working days

**All 13 grades covered:** Nursery, LKG, UKG, Grade 1 through Grade 10

---

## 📁 Files in This Repository

| File | Description |
|------|-------------|
| `CBSE_Academic_Planner_2026_2027.xlsx` | **The ready-to-use Excel workbook** — download and open directly |
| `generate_planner.py` | Python script to regenerate the Excel file with customizations |
| `requirements.txt` | Python dependencies (`openpyxl>=3.1.0`) |
| `Academic_Planner_Setup.gs` | Google Apps Script — alternative for Google Sheets users |
| `setup_guide.md` | Detailed setup and customization instructions |
| `.github/workflows/generate.yml` | GitHub Actions — auto-regenerates `.xlsx` on every push |

---

## 📊 How to Use the Excel File

1. **Download** `CBSE_Academic_Planner_2026_2027.xlsx` from this repository
2. Open in **Microsoft Excel** (recommended) or **Google Sheets** (File → Import)
3. Go to **MASTER INPUT** sheet and update lesson data for your school
4. Check monthly sheets (**APR 2026** through **MAR 2027**) for the academic calendar
5. Review the **DASHBOARD** for a summary view

### Color Coding in Monthly Sheets
| Color | Meaning |
|-------|---------|
| 🔵 Light Blue | Sunday (non-working) |
| 🟠 Light Orange | 2nd/4th Saturday (non-working, CBSE pattern) |
| 🔴 Light Red | Public Holiday |
| 🟣 Light Purple | Vacation Period |
| 🟢 Light Green | Working Day — Lesson assigned |
| 🟡 Light Yellow | Revision / Buffer Day |

---

## 🐍 How to Run the Python Script

Use this to regenerate or customize the Excel file:

```bash
# 1. Install Python dependency
pip install openpyxl>=3.1.0

# 2. Run the script
python generate_planner.py

# 3. Open the generated file
# CBSE_Academic_Planner_2026_2027.xlsx will be created in the same folder
```

See [`setup_guide.md`](setup_guide.md) for detailed customization instructions.

---

## 📱 How to Use the Google Apps Script

For those who prefer Google Sheets:

1. Open **Google Sheets** → Create a new spreadsheet
2. Go to **Extensions → Apps Script**
3. Paste the contents of `Academic_Planner_Setup.gs`
4. Click **Save**, then click **Run → generateAcademicPlanner**
5. Grant permissions when prompted
6. Wait 1-2 minutes — all 16+ sheets will be auto-generated

A custom menu **"🏫 Academic Planner"** will appear with options to:
- ▶️ Generate Full Planner
- 🔄 Refresh Dashboard
- 📅 Update Monthly Sheets
- 🎉 Update Holidays

---

## 🎉 Holidays Included (2026-2027)

### National Holidays
| Date | Holiday |
|------|---------|
| Apr 3, 2026 | Good Friday |
| Apr 14, 2026 | Dr. Babasaheb Ambedkar Jayanti |
| May 1, 2026 | Maharashtra Day / Labour Day |
| May 28, 2026 | Eid ul-Zuha (Bakrid) |
| Jun 26, 2026 | Muharram |
| Aug 12, 2026 | Raksha Bandhan |
| Aug 14, 2026 | Janmashtami / Dahi Handi |
| Aug 15, 2026 | Independence Day |
| Aug 26, 2026 | Id-e-Milad (Prophet's Birthday) |
| Oct 2, 2026 | Mahatma Gandhi Jayanti |
| Oct 21, 2026 | Dussehra / Vijaya Dashami |
| Nov 8, 2026 | Diwali / Lakshmi Puja |
| Nov 9, 2026 | Bali Pratipada |
| Nov 24, 2026 | Guru Nanak Jayanti |
| Dec 25, 2026 | Christmas |
| Jan 26, 2027 | Republic Day |
| Feb 15, 2027 | Maha Shivaratri |
| Mar 5, 2027 | Holi |
| Mar 19, 2027 | Eid-ul-Fitr (Ramzan Eid) |
| Mar 28, 2027 | Ram Navami |

### Maharashtra State Holidays
| Date | Holiday |
|------|---------|
| Jul 17, 2026 | Ashadhi Ekadashi |
| Sep 14-15, 2026 | Ganesh Chaturthi (2 days) |
| Sep 23, 2026 | Ganesh Visarjan |
| Nov 10, 2026 | Bhau Beej |
| Feb 19, 2027 | Chhatrapati Shivaji Maharaj Jayanti |
| Mar 22, 2027 | Gudi Padwa (Marathi New Year) |

### Vacation Periods
| Period | Dates |
|--------|-------|
| Summer Vacation | May 8 – June 14, 2026 |
| Diwali Vacation | November 6–15, 2026 |
| Winter Vacation | December 24, 2026 – January 2, 2027 |

---

## 📚 CBSE Academic Planning Notes

| Parameter | Approx. Days |
|-----------|-------------|
| Total Academic Year | 365 days |
| Total Working Days | ~220 days |
| Public Holidays | ~20 days |
| Vacation Days (Summer + Diwali + Winter) | ~50 days |
| Examination Days (Unit Tests + SA1 + SA2) | ~30 days |
| Net Teaching Days | ~175 days |
| Buffer / Revision Days | ~15 days |

### CBSE Saturday Pattern
- **1st, 3rd, 5th Saturdays** → Working days
- **2nd and 4th Saturdays** → Non-working (school holiday)

### Grade-wise Lesson Days
| Grade Group | Days Per Lesson |
|-------------|----------------|
| Nursery / LKG / UKG | 4–7 days |
| Grade 1–3 | 6–10 days |
| Grade 4–5 | 8–12 days |
| Grade 6–8 | 10–14 days |
| Grade 9–10 | 12–16 days |

---

*🏫 CBSE Academic Planner 2026-2027 — Designed by Principal Haidar Ali Shaikh*
