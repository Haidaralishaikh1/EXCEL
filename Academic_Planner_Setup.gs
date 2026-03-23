/**
 * CBSE Academic Planner 2026-2027
 * Google Apps Script — Auto-generates all sheets in Google Sheets
 * Designed by Principal Haidar Ali Shaikh
 *
 * HOW TO USE:
 * 1. Open Google Sheets → Extensions → Apps Script
 * 2. Paste this entire script
 * 3. Click Save, then click Run → generateAcademicPlanner
 * 4. Grant permissions when prompted
 */

// ─── COLOR CONSTANTS ─────────────────────────────────────────────────────────
const DARK_BLUE   = "#1A237E";
const MED_BLUE    = "#1565C0";
const LIGHT_BLUE  = "#BBDEFB";
const LIGHT_GREEN = "#C8E6C9";
const LIGHT_YELLOW= "#FFF9C4";
const LIGHT_RED   = "#FFCDD2";
const LIGHT_ORANGE= "#FFE0B2";
const LIGHT_PURPLE= "#E1BEE7";
const LIGHT_GRAY  = "#F5F5F5";
const WHITE       = "#FFFFFF";

// ─── GRADES ─────────────────────────────────────────────────────────────────
const GRADES = [
  "Nursery","LKG","UKG",
  "Grade 1","Grade 2","Grade 3","Grade 4","Grade 5",
  "Grade 6","Grade 7","Grade 8","Grade 9","Grade 10"
];

// ─── HOLIDAYS ────────────────────────────────────────────────────────────────
const HOLIDAYS_DATA = [
  [new Date(2026,3,3),  "Good Friday",                         "National",           "All"],
  [new Date(2026,3,14), "Dr. Babasaheb Ambedkar Jayanti",      "National",           "All"],
  [new Date(2026,4,1),  "Maharashtra Day / Labour Day",         "State + National",   "All"],
  [new Date(2026,4,28), "Eid ul-Zuha (Bakrid)",                "National",           "All"],
  [new Date(2026,5,26), "Muharram",                            "National",           "All"],
  [new Date(2026,6,17), "Ashadhi Ekadashi",                    "State (Maharashtra)","All"],
  [new Date(2026,7,12), "Raksha Bandhan",                      "National",           "All"],
  [new Date(2026,7,14), "Janmashtami / Dahi Handi",            "National",           "All"],
  [new Date(2026,7,15), "Independence Day",                    "National",           "All"],
  [new Date(2026,7,16), "Parsi New Year",                      "Regional",           "All"],
  [new Date(2026,7,26), "Id-e-Milad (Prophet's Birthday)",     "National",           "All"],
  [new Date(2026,8,14), "Ganesh Chaturthi",                    "State (Maharashtra)","All"],
  [new Date(2026,8,15), "Ganesh Chaturthi Day 2",              "State (Maharashtra)","All"],
  [new Date(2026,8,23), "Ganesh Visarjan",                     "State (Maharashtra)","All"],
  [new Date(2026,9,2),  "Mahatma Gandhi Jayanti",              "National",           "All"],
  [new Date(2026,9,21), "Dussehra / Vijaya Dashami",           "National",           "All"],
  [new Date(2026,9,22), "Dussehra Holiday",                    "School",             "All"],
  [new Date(2026,10,8), "Diwali / Lakshmi Puja",               "National",           "All"],
  [new Date(2026,10,9), "Bali Pratipada",                      "National",           "All"],
  [new Date(2026,10,10),"Bhau Beej",                           "State (Maharashtra)","All"],
  [new Date(2026,10,24),"Guru Nanak Jayanti",                  "National",           "All"],
  [new Date(2026,11,25),"Christmas",                           "National",           "All"],
  [new Date(2027,0,1),  "New Year's Day",                      "Optional",           "All"],
  [new Date(2027,0,26), "Republic Day",                        "National",           "All"],
  [new Date(2027,1,15), "Maha Shivaratri",                     "National",           "All"],
  [new Date(2027,1,19), "Chhatrapati Shivaji Maharaj Jayanti", "State (Maharashtra)","All"],
  [new Date(2027,2,5),  "Holi",                                "National",           "All"],
  [new Date(2027,2,19), "Eid-ul-Fitr (Ramzan Eid)",            "National",           "All"],
  [new Date(2027,2,22), "Gudi Padwa (Marathi New Year)",       "State (Maharashtra)","All"],
  [new Date(2027,2,28), "Ram Navami",                          "National",           "All"],
];

// Vacation ranges [start, end] (inclusive)
const VACATION_RANGES = [
  [new Date(2026,4,8),  new Date(2026,5,14)],
  [new Date(2026,10,6), new Date(2026,10,15)],
  [new Date(2026,11,24),new Date(2027,0,2)],
];

// ─── LESSON DATA ─────────────────────────────────────────────────────────────
// [Subject, LessonNo, LessonName, [days x 13 grades], StartMonth, Priority]
const MASTER_DATA = [
  ["English",1,"My Family",         [5,6,7,8,8,9,10,10,11,12,12,14,14],"April",1],
  ["English",2,"The Park",          [5,6,7,8,8,9,10,10,11,12,12,14,14],"April",2],
  ["English",3,"Animals Around Us", [5,6,7,8,8,10,10,11,12,12,13,14,15],"May",3],
  ["English",4,"Seasons",           [5,6,7,8,9,10,10,11,12,12,13,15,15],"June",4],
  ["English",5,"Our Helpers",       [5,6,7,8,9,10,10,11,12,13,13,15,16],"July",5],
  ["English",6,"Food We Eat",       [5,6,7,8,9,10,11,11,12,13,14,15,16],"August",6],
  ["Mathematics",1,"Numbers (1-10)",[6,7,8,9,9,10,11,11,12,13,13,14,15],"April",1],
  ["Mathematics",2,"Shapes",        [5,6,7,8,9,9,10,10,11,12,12,13,14],"April",2],
  ["Mathematics",3,"Addition",      [0,6,7,9,9,10,11,11,12,13,13,14,15],"May",3],
  ["Mathematics",4,"Subtraction",   [0,6,7,9,9,10,11,11,12,13,13,14,15],"June",4],
  ["Mathematics",5,"Measurement",   [0,0,7,9,9,10,11,12,12,13,14,14,15],"July",5],
  ["Mathematics",6,"Time",          [0,0,6,8,9,10,11,12,13,13,14,15,16],"August",6],
  ["Hindi",1,"Varnmala",            [6,7,8,9,9,10,11,11,12,12,13,14,14],"April",1],
  ["Hindi",2,"Matra",               [5,6,7,8,9,10,10,11,12,12,13,14,14],"April",2],
  ["Hindi",3,"Shabd Rachna",        [5,6,7,8,9,10,10,11,12,12,13,14,15],"May",3],
  ["Hindi",4,"Vaaky Rachna",        [0,5,6,8,9,10,10,11,12,13,13,14,15],"June",4],
  ["Hindi",5,"Kahani Lekhan",       [0,0,6,8,9,10,11,12,12,13,14,15,16],"July",5],
  ["Hindi",6,"Nibandh",             [0,0,0,8,9,10,11,12,12,13,14,15,16],"August",6],
  ["EVS/Science",1,"My Body",            [5,6,7,8,9,10,10,11,12,12,13,14,15],"April",1],
  ["EVS/Science",2,"My Family",          [5,6,7,8,9,10,10,11,12,12,13,14,15],"April",2],
  ["EVS/Science",3,"Plants Around Us",   [5,6,7,8,9,10,11,11,12,13,13,14,15],"May",3],
  ["EVS/Science",4,"Animals",            [5,6,7,8,9,10,11,12,12,13,14,15,16],"June",4],
  ["EVS/Science",5,"Water",              [5,6,7,8,9,10,11,12,12,13,14,15,16],"July",5],
  ["EVS/Science",6,"Food and Health",    [5,6,7,8,9,10,11,12,13,13,14,15,16],"August",6],
  ["Social Studies",1,"My Neighbourhood",   [0,0,0,0,0,10,11,12,12,13,13,14,15],"April",1],
  ["Social Studies",2,"Our Country",         [0,0,0,0,0,10,11,12,12,13,13,14,15],"May",2],
  ["Social Studies",3,"Maps and Directions", [0,0,0,0,0,10,11,12,13,13,14,15,16],"June",3],
  ["Social Studies",4,"History of India",    [0,0,0,0,0,10,11,12,13,13,14,15,16],"July",4],
  ["Social Studies",5,"Indian Constitution", [0,0,0,0,0,0,11,12,13,14,14,15,16],"August",5],
  ["Social Studies",6,"Geography of India",  [0,0,0,0,0,0,11,12,13,14,14,15,16],"September",6],
  ["Computer",1,"Parts of Computer",        [0,0,0,8,9,10,11,11,12,12,13,14,14],"April",1],
  ["Computer",2,"Using Mouse and Keyboard", [0,0,0,8,9,10,11,12,12,13,13,14,15],"May",2],
  ["Computer",3,"MS Paint",                 [0,0,0,8,9,10,11,12,12,13,13,14,15],"June",3],
  ["Computer",4,"MS Word",                  [0,0,0,0,9,10,11,12,13,13,14,15,16],"July",4],
  ["Computer",5,"Internet Basics",          [0,0,0,0,0,10,11,12,13,14,14,15,16],"August",5],
  ["Computer",6,"Coding Basics",            [0,0,0,0,0,10,11,12,13,14,14,15,16],"September",6],
];

const MONTH_SHEETS = [
  {year:2026,month:3, short:"APR 2026",full:"APRIL 2026"},
  {year:2026,month:4, short:"MAY 2026",full:"MAY 2026"},
  {year:2026,month:5, short:"JUN 2026",full:"JUNE 2026"},
  {year:2026,month:6, short:"JUL 2026",full:"JULY 2026"},
  {year:2026,month:7, short:"AUG 2026",full:"AUGUST 2026"},
  {year:2026,month:8, short:"SEP 2026",full:"SEPTEMBER 2026"},
  {year:2026,month:9, short:"OCT 2026",full:"OCTOBER 2026"},
  {year:2026,month:10,short:"NOV 2026",full:"NOVEMBER 2026"},
  {year:2026,month:11,short:"DEC 2026",full:"DECEMBER 2026"},
  {year:2027,month:0, short:"JAN 2027",full:"JANUARY 2027"},
  {year:2027,month:1, short:"FEB 2027",full:"FEBRUARY 2027"},
  {year:2027,month:2, short:"MAR 2027",full:"MARCH 2027"},
];

// ─── UTILITY FUNCTIONS ────────────────────────────────────────────────────────
function isVacation(d) {
  const dt = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  for (const [start, end] of VACATION_RANGES) {
    if (dt >= start && dt <= end) return true;
  }
  return false;
}

function getHoliday(d) {
  const dt = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  for (const [hdate, hname, htype] of HOLIDAYS_DATA) {
    if (hdate.getTime() === dt.getTime()) return {name: hname, type: htype};
  }
  return null;
}

function isSunday(d) { return d.getDay() === 0; }

function isNthSaturday(d, ns) {
  if (d.getDay() !== 6) return false;
  const firstDay = new Date(d.getFullYear(), d.getMonth(), 1);
  const firstSat = new Date(firstDay);
  while (firstSat.getDay() !== 6) firstSat.setDate(firstSat.getDate() + 1);
  const satNum = Math.floor((d.getDate() - firstSat.getDate()) / 7) + 1;
  return ns.includes(satNum);
}

function getDayStatus(d) {
  if (isSunday(d)) return 'sunday';
  if (isVacation(d)) return 'vacation';
  if (getHoliday(d)) return 'holiday';
  if (isNthSaturday(d, [2, 4])) return '2nd4th_sat';
  return 'working';
}

function getDaysInMonth(year, month) {
  return new Date(year, month + 1, 0).getDate();
}

function getMonthName(monthIndex) {
  const names = ["January","February","March","April","May","June",
                 "July","August","September","October","November","December"];
  return names[monthIndex];
}

function getLessonForDay(monthName, gradeIdx, workingDayNum) {
  const lessons = MASTER_DATA.filter(r => r[4] === monthName && r[3][gradeIdx] > 0);
  if (lessons.length === 0) return "📗 Revision / Buffer";
  let cumulative = 0;
  for (const row of lessons) {
    cumulative += row[3][gradeIdx];
    if (workingDayNum <= cumulative) return row[2];
  }
  return "📗 Revision / Buffer";
}

function styleCell(range, bg, fg, bold, fontSize, hAlign, vAlign, wrap) {
  range.setBackground(bg || WHITE)
       .setFontColor(fg || "#212121")
       .setFontWeight(bold ? "bold" : "normal")
       .setFontSize(fontSize || 10)
       .setHorizontalAlignment(hAlign || "left")
       .setVerticalAlignment(vAlign || "middle")
       .setWrap(wrap !== false);
}

// ─── CREATE CUSTOM MENU ───────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🏫 Academic Planner")
    .addItem("▶️ Generate Full Planner", "generateAcademicPlanner")
    .addSeparator()
    .addItem("🔄 Refresh Dashboard", "refreshDashboard")
    .addItem("📅 Update Monthly Sheets", "updateMonthlySheets")
    .addItem("🎉 Update Holidays", "updateHolidays")
    .addToUi();
}

// ─── MAIN GENERATOR ──────────────────────────────────────────────────────────
function generateAcademicPlanner() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  ui.alert("🏫 CBSE Academic Planner", 
    "Starting to generate all sheets...\nThis may take 1-2 minutes. Please wait.",
    ui.ButtonSet.OK);

  // Delete all existing sheets except the first one
  const allSheets = ss.getSheets();
  if (allSheets.length > 1) {
    for (let i = 1; i < allSheets.length; i++) {
      ss.deleteSheet(allSheets[i]);
    }
  }

  // Rename first sheet temporarily
  allSheets[0].setName("TEMP");

  createInstructionsSheet(ss);
  createMasterInputSheet(ss);
  createHolidaysSheet(ss);

  for (const m of MONTH_SHEETS) {
    createMonthlySheet(ss, m.year, m.month, m.short, m.full);
  }

  createDashboardSheet(ss);

  // Delete temp sheet
  ss.deleteSheet(ss.getSheetByName("TEMP"));

  ss.setActiveSheet(ss.getSheetByName("INSTRUCTIONS"));

  ui.alert("✅ Done!", 
    "CBSE Academic Planner 2026-2027 has been generated successfully!\n\nDesigned by Principal Haidar Ali Shaikh",
    ui.ButtonSet.OK);
}

// ─── SHEET: INSTRUCTIONS ─────────────────────────────────────────────────────
function createInstructionsSheet(ss) {
  const ws = ss.insertSheet("INSTRUCTIONS", 0);

  ws.getRange("A1:H1").merge()
    .setValue("🏫 CBSE Academic Planner 2026-2027")
    .setBackground(DARK_BLUE).setFontColor(WHITE).setFontWeight("bold")
    .setFontSize(18).setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.setRowHeight(1, 50);

  ws.getRange("A2:H2").merge()
    .setValue("Designed by Principal Haidar Ali Shaikh")
    .setBackground(MED_BLUE).setFontColor(WHITE).setFontWeight("bold")
    .setFontSize(12).setFontStyle("italic")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.setRowHeight(2, 35);

  ws.getRange("A4:H4").merge()
    .setValue("📋 HOW TO USE THIS PLANNER")
    .setBackground(DARK_BLUE).setFontColor(WHITE).setFontWeight("bold")
    .setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.setRowHeight(4, 35);

  const instructions = [
    ["STEP 1","Open MASTER INPUT sheet","Enter lesson names, number of days per grade, and start month."],
    ["STEP 2","Review HOLIDAYS sheet","All Indian national and Maharashtra state holidays for 2026-2027 are pre-loaded."],
    ["STEP 3","View Monthly Calendar Sheets","APR 2026 through MAR 2027 show day-by-day lesson assignments with holidays highlighted."],
    ["STEP 4","Check the DASHBOARD","See grade-wise summary of planned lessons, working days, and buffer status."],
    ["STEP 5","Customize as Needed","Add more lessons in MASTER INPUT then use 🔄 Refresh from the menu."],
    ["NOTE","2nd & 4th Saturdays","Following CBSE practice, 2nd and 4th Saturdays are marked as non-working days."],
    ["NOTE","Color Coding","🔵 Sunday | 🟠 2nd/4th Sat | 🔴 Holiday | 🟣 Vacation | 🟢 Working Day | 🟡 Revision"],
    ["NOTE","Vacations","Summer (May 8-Jun 14), Diwali (Nov 6-15), Winter (Dec 24-Jan 2)"],
    ["CBSE","Academic Year","~220 working days | ~30 exam days | ~15 buffer days | ~175 net teaching days"],
  ];

  instructions.forEach(([step, heading, detail], i) => {
    const row = 5 + i;
    const bg = i % 2 === 0 ? LIGHT_GRAY : WHITE;
    ws.getRange(row, 1).setValue(step).setBackground(MED_BLUE).setFontColor(WHITE)
      .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
    ws.getRange(row, 2).setValue(heading).setBackground(bg).setFontWeight("bold")
      .setVerticalAlignment("middle").setWrap(true);
    ws.getRange(row, 3, 1, 6).merge().setValue(detail).setBackground(bg)
      .setVerticalAlignment("middle").setWrap(true);
    ws.setRowHeight(row, 45);
  });

  ws.setColumnWidth(1, 80);
  ws.setColumnWidth(2, 200);
  for (let c = 3; c <= 8; c++) ws.setColumnWidth(c, 150);
  ws.setGridlines(false);
}

// ─── SHEET: MASTER INPUT ─────────────────────────────────────────────────────
function createMasterInputSheet(ss) {
  const ws = ss.insertSheet("MASTER INPUT");
  const headers = ["Subject","Lesson No.","Lesson Name",...GRADES,"Start Month","Priority"];
  const widths = [120, 80, 220, ...Array(13).fill(70), 100, 70];

  ws.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(DARK_BLUE).setFontColor(WHITE).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  ws.setRowHeight(1, 35);

  MASTER_DATA.forEach((row, i) => {
    const r = i + 2;
    const bg = i % 2 === 0 ? WHITE : LIGHT_GRAY;
    const rowData = [row[0], row[1], row[2], ...row[3], row[4], row[5]];
    ws.getRange(r, 1, 1, rowData.length).setValues([rowData]).setBackground(bg)
      .setVerticalAlignment("middle");
    ws.getRange(r, 4, 1, 13).setHorizontalAlignment("center");
    ws.setRowHeight(r, 20);
  });

  headers.forEach((_, i) => ws.setColumnWidth(i + 1, widths[i]));
  ws.setFrozenRows(1);
  ws.setGridlines(false);

  // Month dropdown
  const monthRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["April","May","June","July","August","September",
                         "October","November","December","January","February","March"])
    .setAllowInvalid(false).build();
  ws.getRange(2, headers.length - 1, 100, 1).setDataValidation(monthRule);
}

// ─── SHEET: HOLIDAYS ─────────────────────────────────────────────────────────
function createHolidaysSheet(ss) {
  const ws = ss.insertSheet("HOLIDAYS");
  const headers = ["Date","Day","Holiday Name","Type","Applicable To"];
  const widths = [110, 90, 260, 150, 130];

  ws.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(DARK_BLUE).setFontColor(WHITE).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.setRowHeight(1, 30);
  ws.setFrozenRows(1);

  // Add vacation rows too
  const vacationEntries = [
    [new Date(2026,4,8),  "Summer Vacation Begins","Vacation","All"],
    [new Date(2026,5,14), "Summer Vacation Ends",  "Vacation","All"],
    [new Date(2026,10,6), "Diwali Vacation Begins","Vacation","All"],
    [new Date(2026,10,15),"Diwali Vacation Ends",  "Vacation","All"],
    [new Date(2026,11,24),"Winter Vacation Begins","Vacation","All"],
    [new Date(2027,0,2),  "Winter Vacation Ends",  "Vacation","All"],
  ];

  const allHolidays = [...HOLIDAYS_DATA.map(h => [h[0],h[1],h[2],h[3]]), ...vacationEntries];
  allHolidays.sort((a, b) => a[0] - b[0]);

  const DAYS = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];

  allHolidays.forEach(([hdate, hname, htype, happlicable], i) => {
    const row = i + 2;
    const dayName = DAYS[hdate.getDay()];
    let bg;
    if (htype === "Vacation") bg = LIGHT_PURPLE;
    else if (htype.includes("National")) bg = LIGHT_RED;
    else if (htype.includes("State")) bg = LIGHT_ORANGE;
    else if (htype === "Optional") bg = LIGHT_YELLOW;
    else bg = LIGHT_GRAY;

    ws.getRange(row, 1).setValue(hdate).setNumberFormat("DD-MMM-YYYY")
      .setBackground(bg).setHorizontalAlignment("center").setVerticalAlignment("middle");
    ws.getRange(row, 2).setValue(dayName).setBackground(bg).setHorizontalAlignment("center");
    ws.getRange(row, 3).setValue(hname).setBackground(bg);
    ws.getRange(row, 4).setValue(htype).setBackground(bg).setHorizontalAlignment("center");
    ws.getRange(row, 5).setValue(happlicable).setBackground(bg).setHorizontalAlignment("center");
    ws.setRowHeight(row, 22);
  });

  headers.forEach((_, i) => ws.setColumnWidth(i + 1, widths[i]));
  ws.setGridlines(false);
}

// ─── SHEET: MONTHLY CALENDAR ─────────────────────────────────────────────────
function createMonthlySheet(ss, year, month, sheetName, monthFull) {
  const ws = ss.insertSheet(sheetName);
  const colHeaders = ["Date","Day","Holiday / Event","Working Day?","Day #",...GRADES];
  const colWidths = [110, 80, 200, 85, 50, ...Array(13).fill(140)];
  const NCOLS = colHeaders.length;

  // Row 1: Title
  ws.getRange(1, 1, 1, NCOLS).merge()
    .setValue(`${monthFull} — ACADEMIC CALENDAR`)
    .setBackground(DARK_BLUE).setFontColor(WHITE).setFontWeight("bold")
    .setFontSize(14).setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.setRowHeight(1, 40);

  // Row 2: Headers
  ws.getRange(2, 1, 1, NCOLS).setValues([colHeaders])
    .setBackground(DARK_BLUE).setFontColor(WHITE).setFontWeight("bold")
    .setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  ws.setRowHeight(2, 32);

  ws.setFrozenRows(2);
  ws.setFrozenColumns(5);
  colHeaders.forEach((_, i) => ws.setColumnWidth(i + 1, colWidths[i]));

  const DAYS_FULL = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  const monthName = getMonthName(month);
  const numDays = getDaysInMonth(year, month);
  let workingDayCount = 0;

  const batchData = [];
  const batchFormats = [];

  for (let day = 1; day <= numDays; day++) {
    const d = new Date(year, month, day);
    const status = getDayStatus(d);
    const dayName = DAYS_FULL[d.getDay()];
    const row = day + 2;
    let rowBg, holidayEvent, workingDayStr, dayNum;
    let gradeDisplays;

    if (status === 'sunday') {
      rowBg = LIGHT_BLUE; holidayEvent = ""; workingDayStr = "NO"; dayNum = "";
      gradeDisplays = Array(13).fill("🔴 SUNDAY");
    } else if (status === 'vacation') {
      rowBg = LIGHT_PURPLE; holidayEvent = "VACATION"; workingDayStr = "NO"; dayNum = "";
      gradeDisplays = Array(13).fill("🟣 VACATION");
    } else if (status === 'holiday') {
      const h = getHoliday(d);
      rowBg = LIGHT_RED; holidayEvent = h.name; workingDayStr = "NO"; dayNum = "";
      gradeDisplays = Array(13).fill(`🔴 ${h.name}`);
    } else if (status === '2nd4th_sat') {
      rowBg = LIGHT_ORANGE; holidayEvent = "2nd/4th Saturday"; workingDayStr = "NO"; dayNum = "";
      gradeDisplays = Array(13).fill("🔴 2nd/4th Saturday");
    } else {
      rowBg = WHITE; holidayEvent = ""; workingDayStr = "YES";
      workingDayCount++;
      dayNum = workingDayCount;
      gradeDisplays = GRADES.map((_, gIdx) => getLessonForDay(monthName, gIdx, workingDayCount));
    }

    // Apply row
    const range = ws.getRange(row, 1, 1, NCOLS);
    range.setBackground(rowBg).setVerticalAlignment("middle");
    ws.setRowHeight(row, 24);

    ws.getRange(row, 1).setValue(d).setNumberFormat("DD-MMM-YYYY")
      .setHorizontalAlignment("center").setFontWeight("bold");
    ws.getRange(row, 2).setValue(dayName).setHorizontalAlignment("center");
    ws.getRange(row, 3).setValue(holidayEvent).setWrap(true);
    ws.getRange(row, 4).setValue(workingDayStr).setHorizontalAlignment("center")
      .setFontColor(workingDayStr === "YES" ? "#1B5E20" : "#B71C1C").setFontWeight("bold");
    ws.getRange(row, 5).setValue(dayNum).setHorizontalAlignment("center").setFontWeight("bold");

    gradeDisplays.forEach((lesson, gIdx) => {
      const cellBg = status === 'working' 
        ? (lesson.includes("Revision") || lesson.includes("Buffer") ? LIGHT_YELLOW : LIGHT_GREEN)
        : rowBg;
      ws.getRange(row, 6 + gIdx).setValue(lesson).setBackground(cellBg)
        .setFontSize(9).setWrap(true).setVerticalAlignment("middle");
    });
  }

  ws.setGridlines(false);
}

// ─── SHEET: DASHBOARD ────────────────────────────────────────────────────────
function createDashboardSheet(ss) {
  const ws = ss.insertSheet("DASHBOARD");

  ws.getRange("A1:H1").merge()
    .setValue("📊 ACADEMIC PLANNER DASHBOARD — 2026-2027")
    .setBackground(DARK_BLUE).setFontColor(WHITE).setFontWeight("bold")
    .setFontSize(16).setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.setRowHeight(1, 45);

  ws.getRange("A2:H2").merge()
    .setValue("Designed by Principal Haidar Ali Shaikh | CBSE Academic Year 2026-2027")
    .setBackground(MED_BLUE).setFontColor(WHITE).setFontStyle("italic")
    .setFontSize(11).setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.setRowHeight(2, 30);

  // Section 1: Grade-wise Summary
  ws.getRange("A4:H4").merge()
    .setValue("📌 SECTION 1: GRADE-WISE LESSON SUMMARY")
    .setBackground(DARK_BLUE).setFontColor(WHITE).setFontWeight("bold")
    .setFontSize(12).setHorizontalAlignment("left").setVerticalAlignment("middle");
  ws.setRowHeight(4, 30);

  const gHeaders = ["Grade","Total Lessons","Total Days Planned","Working Days Available","Buffer Days","Status"];
  ws.getRange(5, 1, 1, gHeaders.length).setValues([gHeaders])
    .setBackground(DARK_BLUE).setFontColor(WHITE).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.setRowHeight(5, 28);

  const totalWorkingDays = 220; // approximate CBSE academic year
  GRADES.forEach((grade, gIdx) => {
    const row = 6 + gIdx;
    const totalLessons = MASTER_DATA.filter(r => r[3][gIdx] > 0).length;
    const totalDays = MASTER_DATA.reduce((s, r) => s + r[3][gIdx], 0);
    const buffer = totalWorkingDays - totalDays;
    const status = totalDays === 0 ? "⚠️ No Data" : buffer >= 15 ? "📗 Has Buffer" : buffer >= 0 ? "✅ On Track" : "⚠️ Overloaded";
    const bg = status.includes("✅") || status.includes("📗") ? LIGHT_GREEN : LIGHT_RED;
    const rowBg = gIdx % 2 === 0 ? WHITE : LIGHT_GRAY;

    [grade, totalLessons, totalDays, totalWorkingDays, buffer, status].forEach((val, ci) => {
      const c = ws.getRange(row, ci + 1);
      c.setValue(val).setBackground(ci === 5 ? bg : rowBg)
       .setHorizontalAlignment("center").setVerticalAlignment("middle");
      if (ci === 0) c.setFontWeight("bold");
    });
    ws.setRowHeight(row, 22);
  });

  ws.setColumnWidths(1, 6, [120, 110, 140, 170, 100, 120]);
  ws.setGridlines(false);
}

// ─── MENU ACTION STUBS ───────────────────────────────────────────────────────
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("DASHBOARD");
  if (!ws) { SpreadsheetApp.getUi().alert("DASHBOARD sheet not found. Run Generate first."); return; }
  createDashboardSheet(ss);
  SpreadsheetApp.getUi().alert("✅ Dashboard refreshed!");
}

function updateMonthlySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const m of MONTH_SHEETS) {
    let ws = ss.getSheetByName(m.short);
    if (ws) ss.deleteSheet(ws);
    createMonthlySheet(ss, m.year, m.month, m.short, m.full);
  }
  SpreadsheetApp.getUi().alert("✅ All monthly sheets updated!");
}

function updateHolidays() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName("HOLIDAYS");
  if (ws) ss.deleteSheet(ws);
  createHolidaysSheet(ss);
  SpreadsheetApp.getUi().alert("✅ Holidays sheet updated!");
}
