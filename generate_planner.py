#!/usr/bin/env python3
"""
CBSE Academic Planner 2026-2027 Generator
Designed by Principal Haidar Ali Shaikh

Generates a complete Excel workbook (.xlsx) with:
- INSTRUCTIONS sheet
- MASTER INPUT sheet (lesson data)
- HOLIDAYS sheet
- 12 Monthly Calendar sheets (APR 2026 - MAR 2027)
- DASHBOARD sheet
"""

import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles.numbers import FORMAT_DATE_DDMMYY
import datetime
import calendar

# ─── COLOR PALETTE ───────────────────────────────────────────────────────────
DARK_BLUE   = "1A237E"
MED_BLUE    = "1565C0"
LIGHT_BLUE  = "BBDEFB"
LIGHT_GREEN = "C8E6C9"
LIGHT_YELLOW= "FFF9C4"
LIGHT_RED   = "FFCDD2"
LIGHT_ORANGE= "FFE0B2"
LIGHT_PURPLE= "E1BEE7"
LIGHT_GRAY  = "F5F5F5"
WHITE       = "FFFFFF"
DARK_TEXT   = "212121"
AMBER       = "FFB300"

def make_fill(hex_color):
    return PatternFill(fill_type="solid", fgColor=hex_color)

def make_font(bold=False, color=DARK_TEXT, size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name="Calibri")

def make_border(thin=True):
    s = Side(style="thin" if thin else "medium")
    return Border(left=s, right=s, top=s, bottom=s)

def center_align(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def left_align(wrap=False):
    return Alignment(horizontal="left", vertical="center", wrap_text=wrap)

# ─── HOLIDAYS DATA ────────────────────────────────────────────────────────────
HOLIDAYS = [
    # (date, holiday_name, type, applicable_to)
    (datetime.date(2026,  4,  3), "Good Friday",                        "National",          "All"),
    (datetime.date(2026,  4, 14), "Dr. Babasaheb Ambedkar Jayanti",     "National",          "All"),
    (datetime.date(2026,  5,  1), "Maharashtra Day / Labour Day",        "State + National",  "All"),
    (datetime.date(2026,  5, 28), "Eid ul-Zuha (Bakrid)",               "National",          "All"),
    (datetime.date(2026,  6, 26), "Muharram",                           "National",          "All"),
    (datetime.date(2026,  7, 17), "Ashadhi Ekadashi",                   "State (Maharashtra)","All"),
    (datetime.date(2026,  8, 12), "Raksha Bandhan",                     "National",          "All"),
    (datetime.date(2026,  8, 14), "Janmashtami / Dahi Handi",           "National",          "All"),
    (datetime.date(2026,  8, 15), "Independence Day",                   "National",          "All"),
    (datetime.date(2026,  8, 16), "Parsi New Year",                     "Regional",          "All"),
    (datetime.date(2026,  8, 26), "Id-e-Milad (Prophet's Birthday)",    "National",          "All"),
    (datetime.date(2026,  9, 14), "Ganesh Chaturthi",                   "State (Maharashtra)","All"),
    (datetime.date(2026,  9, 15), "Ganesh Chaturthi Day 2",             "State (Maharashtra)","All"),
    (datetime.date(2026,  9, 23), "Ganesh Visarjan",                    "State (Maharashtra)","All"),
    (datetime.date(2026, 10,  2), "Mahatma Gandhi Jayanti",             "National",          "All"),
    (datetime.date(2026, 10, 21), "Dussehra / Vijaya Dashami",          "National",          "All"),
    (datetime.date(2026, 10, 22), "Dussehra Holiday",                   "School",            "All"),
    (datetime.date(2026, 11,  8), "Diwali / Lakshmi Puja",              "National",          "All"),
    (datetime.date(2026, 11,  9), "Bali Pratipada",                     "National",          "All"),
    (datetime.date(2026, 11, 10), "Bhau Beej",                          "State (Maharashtra)","All"),
    (datetime.date(2026, 11, 24), "Guru Nanak Jayanti",                 "National",          "All"),
    (datetime.date(2026, 12, 25), "Christmas",                          "National",          "All"),
    (datetime.date(2027,  1,  1), "New Year's Day",                     "Optional",          "All"),
    (datetime.date(2027,  1, 26), "Republic Day",                       "National",          "All"),
    (datetime.date(2027,  2, 15), "Maha Shivaratri",                    "National",          "All"),
    (datetime.date(2027,  2, 19), "Chhatrapati Shivaji Maharaj Jayanti","State (Maharashtra)","All"),
    (datetime.date(2027,  3,  5), "Holi",                               "National",          "All"),
    (datetime.date(2027,  3, 19), "Eid-ul-Fitr (Ramzan Eid)",           "National",          "All"),
    (datetime.date(2027,  3, 22), "Gudi Padwa (Marathi New Year)",      "State (Maharashtra)","All"),
    (datetime.date(2027,  3, 28), "Ram Navami",                         "National",          "All"),
    # Vacation markers
    (datetime.date(2026,  5,  8), "Summer Vacation Begins",             "Vacation",          "All"),
    (datetime.date(2026,  6, 14), "Summer Vacation Ends",               "Vacation",          "All"),
    (datetime.date(2026, 11,  6), "Diwali Vacation Begins",             "Vacation",          "All"),
    (datetime.date(2026, 11, 15), "Diwali Vacation Ends",               "Vacation",          "All"),
    (datetime.date(2026, 12, 24), "Winter Vacation Begins",             "Vacation",          "All"),
    (datetime.date(2027,  1,  2), "Winter Vacation Ends",               "Vacation",          "All"),
]

# Vacation ranges for quick lookup
VACATION_RANGES = [
    (datetime.date(2026, 5, 8),   datetime.date(2026, 6, 14)),
    (datetime.date(2026, 11, 6),  datetime.date(2026, 11, 15)),
    (datetime.date(2026, 12, 24), datetime.date(2027, 1, 2)),
]

# Build a dict: date -> (holiday_name, type) for non-vacation holidays
HOLIDAY_DICT = {}
for hdate, hname, htype, happlicable in HOLIDAYS:
    if htype != "Vacation":
        HOLIDAY_DICT[hdate] = (hname, htype)

def is_vacation(d):
    for start, end in VACATION_RANGES:
        if start <= d <= end:
            return True
    return False

def is_holiday(d):
    return d in HOLIDAY_DICT

def is_sunday(d):
    return d.weekday() == 6

def is_nth_saturday(d, ns=(2, 4)):
    """Check if date is a 2nd or 4th Saturday"""
    if d.weekday() != 5:
        return False
    # Which Saturday in the month?
    first_day = d.replace(day=1)
    # Find first Saturday
    first_sat = first_day + datetime.timedelta(days=(5 - first_day.weekday()) % 7)
    sat_num = ((d - first_sat).days // 7) + 1
    return sat_num in ns

def day_status(d):
    """Returns: 'sunday', 'vacation', 'holiday', '2nd4th_sat', 'working'"""
    if is_sunday(d):
        return 'sunday'
    if is_vacation(d):
        return 'vacation'
    if is_holiday(d):
        return 'holiday'
    if is_nth_saturday(d, (2, 4)):
        return '2nd4th_sat'
    return 'working'

# ─── LESSON DATA ──────────────────────────────────────────────────────────────
GRADES = [
    "Nursery", "LKG", "UKG",
    "Grade 1", "Grade 2", "Grade 3", "Grade 4", "Grade 5",
    "Grade 6", "Grade 7", "Grade 8", "Grade 9", "Grade 10"
]

MONTHS_ORDER = [
    "April", "May", "June", "July", "August", "September",
    "October", "November", "December", "January", "February", "March"
]

# Days per lesson by grade tier
# Nursery/LKG/UKG: 4-7, Grade 1-3: 6-10, Grade 4-5: 8-12, Grade 6-8: 10-14, Grade 9-10: 12-16

MASTER_DATA = [
    # Subject, LessonNo, LessonName, [Nursery,LKG,UKG,G1,G2,G3,G4,G5,G6,G7,G8,G9,G10], StartMonth, Priority
    ["English", 1, "My Family",         [5,6,7,8,8,9,10,10,11,12,12,14,14], "April",   1],
    ["English", 2, "The Park",          [5,6,7,8,8,9,10,10,11,12,12,14,14], "April",   2],
    ["English", 3, "Animals Around Us", [5,6,7,8,8,10,10,11,12,12,13,14,15],"May",     3],
    ["English", 4, "Seasons",           [5,6,7,8,9,10,10,11,12,12,13,15,15],"June",    4],
    ["English", 5, "Our Helpers",       [5,6,7,8,9,10,10,11,12,13,13,15,16],"July",    5],
    ["English", 6, "Food We Eat",       [5,6,7,8,9,10,11,11,12,13,14,15,16],"August",  6],

    ["Mathematics", 1, "Numbers (1-10)",  [6,7,8,9,9,10,11,11,12,13,13,14,15], "April",  1],
    ["Mathematics", 2, "Shapes",          [5,6,7,8,9,9,10,10,11,12,12,13,14],  "April",  2],
    ["Mathematics", 3, "Addition",        [0,6,7,9,9,10,11,11,12,13,13,14,15],  "May",    3],
    ["Mathematics", 4, "Subtraction",     [0,6,7,9,9,10,11,11,12,13,13,14,15],  "June",   4],
    ["Mathematics", 5, "Measurement",     [0,0,7,9,9,10,11,12,12,13,14,14,15],  "July",   5],
    ["Mathematics", 6, "Time",            [0,0,6,8,9,10,11,12,13,13,14,15,16],  "August", 6],

    ["Hindi", 1, "Varnmala",        [6,7,8,9,9,10,11,11,12,12,13,14,14], "April",  1],
    ["Hindi", 2, "Matra",           [5,6,7,8,9,10,10,11,12,12,13,14,14], "April",  2],
    ["Hindi", 3, "Shabd Rachna",    [5,6,7,8,9,10,10,11,12,12,13,14,15], "May",    3],
    ["Hindi", 4, "Vaaky Rachna",    [0,5,6,8,9,10,10,11,12,13,13,14,15], "June",   4],
    ["Hindi", 5, "Kahani Lekhan",   [0,0,6,8,9,10,11,12,12,13,14,15,16], "July",   5],
    ["Hindi", 6, "Nibandh",         [0,0,0,8,9,10,11,12,12,13,14,15,16], "August", 6],

    ["EVS/Science", 1, "My Body",            [5,6,7,8,9,10,10,11,12,12,13,14,15], "April",   1],
    ["EVS/Science", 2, "My Family",          [5,6,7,8,9,10,10,11,12,12,13,14,15], "April",   2],
    ["EVS/Science", 3, "Plants Around Us",   [5,6,7,8,9,10,11,11,12,13,13,14,15], "May",     3],
    ["EVS/Science", 4, "Animals",            [5,6,7,8,9,10,11,12,12,13,14,15,16], "June",    4],
    ["EVS/Science", 5, "Water",              [5,6,7,8,9,10,11,12,12,13,14,15,16], "July",    5],
    ["EVS/Science", 6, "Food and Health",    [5,6,7,8,9,10,11,12,13,13,14,15,16], "August",  6],

    ["Social Studies", 1, "My Neighbourhood",       [0,0,0,0,0,10,11,12,12,13,13,14,15], "April",   1],
    ["Social Studies", 2, "Our Country",             [0,0,0,0,0,10,11,12,12,13,13,14,15], "May",     2],
    ["Social Studies", 3, "Maps and Directions",     [0,0,0,0,0,10,11,12,13,13,14,15,16], "June",    3],
    ["Social Studies", 4, "History of India",        [0,0,0,0,0,10,11,12,13,13,14,15,16], "July",    4],
    ["Social Studies", 5, "Indian Constitution",     [0,0,0,0,0,0,11,12,13,14,14,15,16],  "August",  5],
    ["Social Studies", 6, "Geography of India",      [0,0,0,0,0,0,11,12,13,14,14,15,16],  "September",6],

    ["Computer", 1, "Parts of Computer",        [0,0,0,8,9,10,11,11,12,12,13,14,14], "April",   1],
    ["Computer", 2, "Using Mouse and Keyboard", [0,0,0,8,9,10,11,12,12,13,13,14,15], "May",     2],
    ["Computer", 3, "MS Paint",                 [0,0,0,8,9,10,11,12,12,13,13,14,15], "June",    3],
    ["Computer", 4, "MS Word",                  [0,0,0,0,9,10,11,12,13,13,14,15,16], "July",    4],
    ["Computer", 5, "Internet Basics",           [0,0,0,0,0,10,11,12,13,14,14,15,16], "August",  5],
    ["Computer", 6, "Coding Basics",             [0,0,0,0,0,10,11,12,13,14,14,15,16], "September",6],
]

# ─── MONTHS TO GENERATE ───────────────────────────────────────────────────────
ACADEMIC_MONTHS = [
    (2026,  4, "APR 2026", "APRIL 2026"),
    (2026,  5, "MAY 2026", "MAY 2026"),
    (2026,  6, "JUN 2026", "JUNE 2026"),
    (2026,  7, "JUL 2026", "JULY 2026"),
    (2026,  8, "AUG 2026", "AUGUST 2026"),
    (2026,  9, "SEP 2026", "SEPTEMBER 2026"),
    (2026, 10, "OCT 2026", "OCTOBER 2026"),
    (2026, 11, "NOV 2026", "NOVEMBER 2026"),
    (2026, 12, "DEC 2026", "DECEMBER 2026"),
    (2027,  1, "JAN 2027", "JANUARY 2027"),
    (2027,  2, "FEB 2027", "FEBRUARY 2027"),
    (2027,  3, "MAR 2027", "MARCH 2027"),
]

# ─── HELPER: apply header style ───────────────────────────────────────────────
def style_header_cell(cell, text, bold=True, size=10, bg=DARK_BLUE, fg=WHITE, wrap=True):
    cell.value = text
    cell.font = make_font(bold=bold, color=fg, size=size)
    cell.fill = make_fill(bg)
    cell.alignment = center_align(wrap=wrap)
    cell.border = make_border()

def style_data_cell(cell, value, bold=False, bg=WHITE, fg=DARK_TEXT, align="left", wrap=False, fmt=None):
    cell.value = value
    cell.font = make_font(bold=bold, color=fg)
    cell.fill = make_fill(bg)
    cell.alignment = left_align(wrap) if align == "left" else center_align(wrap)
    cell.border = make_border()
    if fmt:
        cell.number_format = fmt


# ─── SHEET 1: INSTRUCTIONS ────────────────────────────────────────────────────
def create_instructions_sheet(wb):
    ws = wb.create_sheet("INSTRUCTIONS")
    ws.sheet_view.showGridLines = False

    # Title
    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = "🏫 CBSE Academic Planner 2026-2027"
    c.font = Font(bold=True, color=WHITE, size=18, name="Calibri")
    c.fill = make_fill(DARK_BLUE)
    c.alignment = center_align()

    # Subtitle
    ws.merge_cells("A2:H2")
    c = ws["A2"]
    c.value = "Designed by Principal Haidar Ali Shaikh"
    c.font = Font(bold=True, color=WHITE, size=12, name="Calibri", italic=True)
    c.fill = make_fill(MED_BLUE)
    c.alignment = center_align()

    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 30

    # Blank row
    ws.row_dimensions[3].height = 10

    # Instructions heading
    ws.merge_cells("A4:H4")
    c = ws["A4"]
    c.value = "📋 HOW TO USE THIS PLANNER"
    c.font = Font(bold=True, color=WHITE, size=13, name="Calibri")
    c.fill = make_fill(DARK_BLUE)
    c.alignment = center_align()
    ws.row_dimensions[4].height = 28

    instructions = [
        ("STEP 1", "Open the 'MASTER INPUT' sheet", "Enter your lesson names, number of days per grade, and start month."),
        ("STEP 2", "Review the 'HOLIDAYS' sheet", "All Indian national and Maharashtra state holidays for 2026-2027 are pre-loaded. Add or remove as needed."),
        ("STEP 3", "View Monthly Calendar Sheets", "Sheets APR 2026 through MAR 2027 show day-by-day lesson assignments per grade, with holidays highlighted."),
        ("STEP 4", "Check the DASHBOARD", "See grade-wise summary of planned lessons, working days, and buffer status at a glance."),
        ("STEP 5", "Customize as Needed", "Add more lessons in MASTER INPUT. Run the Python script (generate_planner.py) to regenerate with new data."),
        ("NOTE",   "2nd & 4th Saturdays", "Following CBSE practice, 2nd and 4th Saturdays are marked as non-working days (orange rows)."),
        ("NOTE",   "Color Coding", "🔵 Blue = Sunday | 🟠 Orange = 2nd/4th Saturday | 🔴 Red = Holiday | 🟣 Purple = Vacation | 🟢 Green = Working Day with Lesson | 🟡 Yellow = Revision/Buffer"),
        ("NOTE",   "Vacation Periods", "Summer (May 8 - Jun 14), Diwali (Nov 6-15), Winter (Dec 24 - Jan 2)"),
        ("GRADES", "13 Grades Covered", "Nursery, LKG, UKG, Grade 1 through Grade 10"),
        ("CBSE",   "Academic Year", "220 working days approx. | ~30 days for exams | ~15 buffer/revision days | ~175 net teaching days"),
    ]

    for i, (step, heading, detail) in enumerate(instructions):
        row = 5 + i
        ws.row_dimensions[row].height = 38
        bg = LIGHT_GRAY if i % 2 == 0 else WHITE

        ws["A" + str(row)].value = step
        ws["A" + str(row)].font = Font(bold=True, color=WHITE, size=10, name="Calibri")
        ws["A" + str(row)].fill = make_fill(MED_BLUE)
        ws["A" + str(row)].alignment = center_align(wrap=True)
        ws["A" + str(row)].border = make_border()

        ws["B" + str(row)].value = heading
        ws["B" + str(row)].font = Font(bold=True, size=10, name="Calibri")
        ws["B" + str(row)].fill = make_fill(bg)
        ws["B" + str(row)].alignment = left_align(wrap=True)
        ws["B" + str(row)].border = make_border()

        ws.merge_cells(f"C{row}:H{row}")
        ws["C" + str(row)].value = detail
        ws["C" + str(row)].font = Font(size=10, name="Calibri")
        ws["C" + str(row)].fill = make_fill(bg)
        ws["C" + str(row)].alignment = left_align(wrap=True)
        ws["C" + str(row)].border = make_border()

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 28
    for col in "CDEFGH":
        ws.column_dimensions[col].width = 20

    # Footer
    last_row = 5 + len(instructions) + 1
    ws.merge_cells(f"A{last_row}:H{last_row}")
    c = ws[f"A{last_row}"]
    c.value = "📧 For support: Contact your school IT coordinator | Generated by generate_planner.py using openpyxl"
    c.font = Font(italic=True, size=9, color="757575", name="Calibri")
    c.alignment = center_align()


# ─── SHEET 2: MASTER INPUT ────────────────────────────────────────────────────
def create_master_input_sheet(wb):
    ws = wb.create_sheet("MASTER INPUT")
    ws.sheet_view.showGridLines = False

    headers = ["Subject", "Lesson No.", "Lesson Name"] + GRADES + ["Start Month", "Priority"]
    col_widths = [15, 10, 28] + [9] * 13 + [12, 8]

    # Header row
    for col_idx, header in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col_idx)
        style_header_cell(c, header, size=9)

    ws.row_dimensions[1].height = 32

    # Apply column widths
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Data rows
    for row_idx, lesson in enumerate(MASTER_DATA, start=2):
        subject, lesson_no, lesson_name, days_list, start_month, priority = lesson
        bg = WHITE if row_idx % 2 == 0 else LIGHT_GRAY
        row_data = [subject, lesson_no, lesson_name] + days_list + [start_month, priority]
        for col_idx, val in enumerate(row_data, start=1):
            c = ws.cell(row=row_idx, column=col_idx)
            style_data_cell(c, val, bg=bg)
            if col_idx >= 4 and col_idx <= 16:
                c.alignment = center_align()

    ws.row_dimensions[1].height = 30

    # Freeze Row 1
    ws.freeze_panes = "A2"

    # Auto-filter
    last_col = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col}{len(MASTER_DATA)+1}"

    # Data validation for Start Month (column Q = 17)
    months_str = '"April,May,June,July,August,September,October,November,December,January,February,March"'
    dv = DataValidation(type="list", formula1=months_str, allow_blank=True)
    dv.sqref = f"Q2:Q{len(MASTER_DATA)+50}"
    ws.add_data_validation(dv)


# ─── SHEET 3: HOLIDAYS ────────────────────────────────────────────────────────
def create_holidays_sheet(wb):
    ws = wb.create_sheet("HOLIDAYS")
    ws.sheet_view.showGridLines = False

    headers = ["Date", "Day", "Holiday Name", "Type", "Applicable To"]
    widths   = [15,     12,    35,             20,      18]

    for col_idx, (header, width) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=1, column=col_idx)
        style_header_cell(c, header)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"

    # Sort holidays by date
    sorted_holidays = sorted(HOLIDAYS, key=lambda x: x[0])

    for row_idx, (hdate, hname, htype, happlicable) in enumerate(sorted_holidays, start=2):
        day_name = hdate.strftime("%A")

        # Color by type
        if htype == "Vacation":
            bg = LIGHT_PURPLE
        elif "National" in htype:
            bg = LIGHT_RED
        elif "State" in htype:
            bg = LIGHT_ORANGE
        elif htype == "Optional":
            bg = LIGHT_YELLOW
        else:
            bg = LIGHT_GRAY

        c_date = ws.cell(row=row_idx, column=1, value=hdate)
        c_date.number_format = "DD-MMM-YYYY"
        c_date.font = make_font()
        c_date.fill = make_fill(bg)
        c_date.alignment = center_align()
        c_date.border = make_border()

        for col_idx, val in enumerate([day_name, hname, htype, happlicable], start=2):
            c = ws.cell(row=row_idx, column=col_idx)
            style_data_cell(c, val, bg=bg)
            if col_idx in (2, 4, 5):
                c.alignment = center_align()

        ws.row_dimensions[row_idx].height = 20


# ─── LESSON ASSIGNMENT LOGIC ──────────────────────────────────────────────────
def build_lesson_schedule():
    """
    Build a per-grade, per-month lesson schedule.
    Returns: { month_name: { grade_idx: [(lesson_name, days), ...] } }
    """
    schedule = {}
    for month in MONTHS_ORDER:
        grade_lessons = {}
        for g_idx in range(13):
            lessons = []
            for row in MASTER_DATA:
                _, _, lesson_name, days_list, start_month, _ = row
                if start_month == month and days_list[g_idx] > 0:
                    lessons.append((lesson_name, days_list[g_idx]))
            grade_lessons[g_idx] = lessons
        schedule[month] = grade_lessons
    return schedule

LESSON_SCHEDULE = build_lesson_schedule()


def get_lesson_for_day(month_name, grade_idx, working_day_num):
    """
    Given month, grade, and cumulative working day number (1-based),
    return the lesson name to display.
    """
    lessons = LESSON_SCHEDULE.get(month_name, {}).get(grade_idx, [])
    if not lessons:
        return "📗 Revision / Buffer"
    cumulative = 0
    for lesson_name, days in lessons:
        cumulative += days
        if working_day_num <= cumulative:
            return lesson_name
    return "📗 Revision / Buffer"


# ─── SHEETS 4-15: MONTHLY CALENDARS ──────────────────────────────────────────
def create_monthly_sheet(wb, year, month, sheet_name, month_full):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    # Month name string for lesson lookup
    month_name_str = datetime.date(year, month, 1).strftime("%B")

    # Column headers
    col_headers = ["Date", "Day", "Holiday / Event", "Working Day?", "Day #"] + GRADES
    col_widths   = [14,     10,    28,                 11,             6] + [18] * 13
    NCOLS = len(col_headers)

    # Row 1: Title (merged)
    ws.merge_cells(f"A1:{get_column_letter(NCOLS)}1")
    c = ws["A1"]
    c.value = f"{month_full} — ACADEMIC CALENDAR"
    c.font = Font(bold=True, color=WHITE, size=14, name="Calibri")
    c.fill = make_fill(DARK_BLUE)
    c.alignment = center_align()
    ws.row_dimensions[1].height = 34

    # Row 2: Column headers
    for col_idx, (header, width) in enumerate(zip(col_headers, col_widths), start=1):
        c = ws.cell(row=2, column=col_idx)
        style_header_cell(c, header, size=9)
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[2].height = 28

    # Freeze rows 1-2 and columns A-E
    ws.freeze_panes = "F3"

    # Auto-filter on row 2
    ws.auto_filter.ref = f"A2:{get_column_letter(NCOLS)}2"

    # Generate all dates in the month
    num_days = calendar.monthrange(year, month)[1]
    working_day_count = 0

    for day in range(1, num_days + 1):
        d = datetime.date(year, month, day)
        row = day + 2  # data starts at row 3

        status = day_status(d)
        day_name = d.strftime("%A")

        # Determine row styling and cell values
        if status == 'sunday':
            row_bg = LIGHT_BLUE
            holiday_event = ""
            working_day_str = "NO"
            day_num = ""
            grade_display = ["🔴 SUNDAY"] * 13
        elif status == 'vacation':
            row_bg = LIGHT_PURPLE
            holiday_event = "VACATION"
            working_day_str = "NO"
            day_num = ""
            grade_display = ["🟣 VACATION"] * 13
        elif status == 'holiday':
            row_bg = LIGHT_RED
            hname, htype = HOLIDAY_DICT[d]
            holiday_event = hname
            working_day_str = "NO"
            day_num = ""
            grade_display = [f"🔴 {hname}"] * 13
        elif status == '2nd4th_sat':
            row_bg = LIGHT_ORANGE
            holiday_event = "2nd/4th Saturday"
            working_day_str = "NO"
            day_num = ""
            grade_display = ["🔴 2nd/4th Saturday"] * 13
        else:  # working
            row_bg = WHITE
            holiday_event = ""
            working_day_str = "YES"
            working_day_count += 1
            day_num = working_day_count
            grade_display = []
            for g_idx in range(13):
                lesson = get_lesson_for_day(month_name_str, g_idx, working_day_count)
                grade_display.append(lesson)

        # Write Date cell
        c_date = ws.cell(row=row, column=1, value=d)
        c_date.number_format = "DD-MMM-YYYY"
        c_date.font = make_font(bold=(status in ('holiday', '2nd4th_sat')))
        c_date.fill = make_fill(row_bg)
        c_date.alignment = center_align()
        c_date.border = make_border()

        # Write Day cell
        c_day = ws.cell(row=row, column=2, value=day_name)
        c_day.font = make_font()
        c_day.fill = make_fill(row_bg)
        c_day.alignment = center_align()
        c_day.border = make_border()

        # Write Holiday/Event cell
        c_hol = ws.cell(row=row, column=3, value=holiday_event)
        c_hol.font = make_font(bold=bool(holiday_event))
        c_hol.fill = make_fill(row_bg)
        c_hol.alignment = left_align(wrap=True)
        c_hol.border = make_border()

        # Write Working Day cell
        c_wd = ws.cell(row=row, column=4, value=working_day_str)
        c_wd.font = make_font(bold=True, color=(DARK_TEXT if working_day_str == "YES" else "B71C1C"))
        c_wd.fill = make_fill(LIGHT_GREEN if working_day_str == "YES" else row_bg)
        c_wd.alignment = center_align()
        c_wd.border = make_border()

        # Write Day # cell
        c_daynum = ws.cell(row=row, column=5, value=day_num)
        c_daynum.font = make_font(bold=True)
        c_daynum.fill = make_fill(row_bg)
        c_daynum.alignment = center_align()
        c_daynum.border = make_border()

        # Write Grade columns
        for g_idx, lesson_text in enumerate(grade_display):
            col_idx = 6 + g_idx
            c_g = ws.cell(row=row, column=col_idx, value=lesson_text)

            if status == 'working':
                if "Revision" in lesson_text or "Buffer" in lesson_text:
                    cell_bg = LIGHT_YELLOW
                else:
                    cell_bg = LIGHT_GREEN
            else:
                cell_bg = row_bg

            c_g.font = make_font(size=9)
            c_g.fill = make_fill(cell_bg)
            c_g.alignment = left_align(wrap=True)
            c_g.border = make_border()

        ws.row_dimensions[row].height = 22

    return working_day_count


# ─── SHEET 16: DASHBOARD ──────────────────────────────────────────────────────
def create_dashboard_sheet(wb, month_stats):
    """
    month_stats: list of (month_label, year, month_num, total_days, working_days,
                          holidays, sundays, sat_2nd4th) tuples
    """
    ws = wb.create_sheet("DASHBOARD")
    ws.sheet_view.showGridLines = False

    # Title
    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = "📊 ACADEMIC PLANNER DASHBOARD — 2026-2027"
    c.font = Font(bold=True, color=WHITE, size=16, name="Calibri")
    c.fill = make_fill(DARK_BLUE)
    c.alignment = center_align()
    ws.row_dimensions[1].height = 38

    ws.merge_cells("A2:H2")
    c = ws["A2"]
    c.value = "Designed by Principal Haidar Ali Shaikh | CBSE Academic Year 2026-2027"
    c.font = Font(italic=True, color=WHITE, size=11, name="Calibri")
    c.fill = make_fill(MED_BLUE)
    c.alignment = center_align()
    ws.row_dimensions[2].height = 26

    # ── Section 1: Grade-wise Summary ─────────────────────────────────────────
    ws.merge_cells("A4:H4")
    c = ws["A4"]
    c.value = "📌 SECTION 1: GRADE-WISE LESSON SUMMARY"
    c.font = Font(bold=True, color=WHITE, size=12, name="Calibri")
    c.fill = make_fill(DARK_BLUE)
    c.alignment = left_align()
    ws.row_dimensions[4].height = 26

    g_headers = ["Grade", "Total Lessons", "Total Days Planned",
                 "Working Days Available", "Buffer Days", "Status"]
    g_widths   = [14,      14,              18,              22,                    12,           14]

    for col_idx, (h, w) in enumerate(zip(g_headers, g_widths), start=1):
        c = ws.cell(row=5, column=col_idx)
        style_header_cell(c, h, size=10)
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[5].height = 26

    # Calculate total working days across all months
    total_working_days = sum(s[4] for s in month_stats)

    for g_idx, grade in enumerate(GRADES):
        row = 6 + g_idx
        total_lessons = sum(1 for row_d in MASTER_DATA if row_d[3][g_idx] > 0)
        total_days_planned = sum(row_d[3][g_idx] for row_d in MASTER_DATA)
        buffer_days = total_working_days - total_days_planned

        if total_days_planned == 0:
            status = "⚠️ No Data"
        elif buffer_days >= 15:
            status = "📗 Has Buffer"
        elif buffer_days >= 0:
            status = "✅ On Track"
        else:
            status = "⚠️ Overloaded"

        bg = LIGHT_GREEN if "✅" in status or "📗" in status else LIGHT_RED
        row_data = [grade, total_lessons, total_days_planned, total_working_days, buffer_days, status]

        for col_idx, val in enumerate(row_data, start=1):
            c = ws.cell(row=row, column=col_idx)
            c.value = val
            c.font = make_font(bold=(col_idx == 1))
            c.fill = make_fill(bg if col_idx == 6 else (LIGHT_GRAY if g_idx % 2 == 0 else WHITE))
            c.alignment = center_align()
            c.border = make_border()
        ws.row_dimensions[row].height = 20

    # ── Section 2: Month-wise Working Days ────────────────────────────────────
    sec2_start = 6 + len(GRADES) + 2

    ws.merge_cells(f"A{sec2_start}:H{sec2_start}")
    c = ws[f"A{sec2_start}"]
    c.value = "📅 SECTION 2: MONTH-WISE WORKING DAYS SUMMARY"
    c.font = Font(bold=True, color=WHITE, size=12, name="Calibri")
    c.fill = make_fill(DARK_BLUE)
    c.alignment = left_align()
    ws.row_dimensions[sec2_start].height = 26

    m_headers = ["Month", "Total Days", "Working Days", "Holidays", "Sundays", "2nd/4th Saturdays"]
    for col_idx, h in enumerate(m_headers, start=1):
        c = ws.cell(row=sec2_start + 1, column=col_idx)
        style_header_cell(c, h, size=10)
    ws.row_dimensions[sec2_start + 1].height = 26

    totals = [0, 0, 0, 0, 0]
    for i, (month_label, yr, mo, total_d, working_d, holidays_d, sundays_d, sat24_d) in enumerate(month_stats):
        row = sec2_start + 2 + i
        bg = LIGHT_GRAY if i % 2 == 0 else WHITE
        row_data = [month_label, total_d, working_d, holidays_d, sundays_d, sat24_d]
        for col_idx, val in enumerate(row_data, start=1):
            c = ws.cell(row=row, column=col_idx)
            c.value = val
            c.font = make_font(bold=(col_idx == 1))
            c.fill = make_fill(bg)
            c.alignment = center_align()
            c.border = make_border()
        ws.row_dimensions[row].height = 20
        totals[0] += total_d
        totals[1] += working_d
        totals[2] += holidays_d
        totals[3] += sundays_d
        totals[4] += sat24_d

    # Totals row
    total_row = sec2_start + 2 + len(month_stats)
    ws.cell(row=total_row, column=1).value = "TOTAL"
    ws.cell(row=total_row, column=1).font = make_font(bold=True, color=WHITE)
    ws.cell(row=total_row, column=1).fill = make_fill(MED_BLUE)
    ws.cell(row=total_row, column=1).alignment = center_align()
    ws.cell(row=total_row, column=1).border = make_border()
    for i, val in enumerate(totals, start=2):
        c = ws.cell(row=total_row, column=i)
        c.value = val
        c.font = make_font(bold=True, color=WHITE)
        c.fill = make_fill(MED_BLUE)
        c.alignment = center_align()
        c.border = make_border()
    ws.row_dimensions[total_row].height = 22

    # ── Section 3: Quick Stats ─────────────────────────────────────────────────
    sec3_start = total_row + 2

    ws.merge_cells(f"A{sec3_start}:H{sec3_start}")
    c = ws[f"A{sec3_start}"]
    c.value = "⚡ SECTION 3: QUICK STATS"
    c.font = Font(bold=True, color=WHITE, size=12, name="Calibri")
    c.fill = make_fill(DARK_BLUE)
    c.alignment = left_align()
    ws.row_dimensions[sec3_start].height = 26

    total_vacation = sum(
        (end - start).days + 1
        for start, end in VACATION_RANGES
    )
    net_teaching = totals[1] - 30  # subtract ~30 exam days

    quick_stats = [
        ("📅 Total Working Days in Academic Year", totals[1], LIGHT_GREEN),
        ("🎉 Total Public Holidays",               totals[2], LIGHT_RED),
        ("🏖️ Total Vacation Days",                 total_vacation, LIGHT_PURPLE),
        ("📚 Net Teaching Days (est. after exams)", net_teaching, LIGHT_YELLOW),
        ("📆 Total Sundays",                        totals[3], LIGHT_BLUE),
        ("🗓️ 2nd/4th Saturdays Off",               totals[4], LIGHT_ORANGE),
    ]

    for i, (label, val, bg) in enumerate(quick_stats):
        row = sec3_start + 1 + i
        ws.merge_cells(f"A{row}:E{row}")
        c = ws[f"A{row}"]
        c.value = label
        c.font = make_font(bold=True, size=11)
        c.fill = make_fill(bg)
        c.alignment = left_align()
        c.border = make_border()

        ws.merge_cells(f"F{row}:H{row}")
        c2 = ws[f"F{row}"]
        c2.value = val
        c2.font = Font(bold=True, size=14, name="Calibri")
        c2.fill = make_fill(bg)
        c2.alignment = center_align()
        c2.border = make_border()
        ws.row_dimensions[row].height = 26


# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    print("🏫 CBSE Academic Planner 2026-2027 Generator")
    print("   Designed by Principal Haidar Ali Shaikh")
    print("=" * 55)

    wb = openpyxl.Workbook()
    # Remove the default sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    print("📋 Creating INSTRUCTIONS sheet...")
    create_instructions_sheet(wb)

    print("📝 Creating MASTER INPUT sheet...")
    create_master_input_sheet(wb)

    print("🎉 Creating HOLIDAYS sheet...")
    create_holidays_sheet(wb)

    month_stats = []
    for year, month, sheet_name, month_full in ACADEMIC_MONTHS:
        print(f"📅 Creating {sheet_name} calendar sheet...")
        month_name_str = datetime.date(year, month, 1).strftime("%B")
        num_days = calendar.monthrange(year, month)[1]

        # Count stats for this month
        working_days = 0
        holidays_count = 0
        sundays_count = 0
        sat24_count = 0

        for day in range(1, num_days + 1):
            d = datetime.date(year, month, day)
            status = day_status(d)
            if status == 'working':
                working_days += 1
            elif status == 'holiday':
                holidays_count += 1
            elif status == 'sunday':
                sundays_count += 1
            elif status == '2nd4th_sat':
                sat24_count += 1
            # vacations are not separately counted in month stats (they overlap with other days)

        working_days_actual = create_monthly_sheet(wb, year, month, sheet_name, month_full)
        month_stats.append((
            month_full, year, month, num_days,
            working_days_actual, holidays_count, sundays_count, sat24_count
        ))

    print("📊 Creating DASHBOARD sheet...")
    create_dashboard_sheet(wb, month_stats)

    filename = "CBSE_Academic_Planner_2026_2027.xlsx"
    wb.save(filename)
    print(f"\n✅ Excel workbook saved as: {filename}")
    print(f"   Total sheets: {len(wb.sheetnames)}")
    print(f"   Sheets: {', '.join(wb.sheetnames)}")
    print("\n🎓 Ready for use by Principal Haidar Ali Shaikh!")


if __name__ == "__main__":
    main()
