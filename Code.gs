// ============================================================
// BeGifted Package Alert Dashboard — Code.gs
// ============================================================

const SPREADSHEET_ID_CREDITS   = "100bidSt63ynf_y7Iq-nRQUj3MMQoNpnltOQdSiwHN-0";
const SPREADSHEET_ID_ANALYTICS = "15XTOU1kYDib4stuiFzbTOT1MAeRlCw20maOXk5irfsc";

const SHEET_AGGREGATIONS      = "Aggregations";
const SHEET_CREDIT_CONTROL    = "Credit_Control";
const SHEET_UPCOMING          = "Upcoming Sessions";
const SHEET_STUDENTS          = "Students";
const SHEET_STUDENTS_COURSES  = "Students & Courses";

const ALERT_THRESHOLD    = 2;
const NOTIFY_WINDOW_DAYS = 30;

// ============================================================
// ENTRY POINT
// ============================================================

function doGet() {
  return HtmlService
    .createTemplateFromFile("dashboard")
    .evaluate()
    .setTitle("BeGifted — Package Expiry Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// MAIN DATA FUNCTION
// ============================================================

function getStudentData() {
  try {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const ssCredits   = SpreadsheetApp.openById(SPREADSHEET_ID_CREDITS);
    const ssAnalytics = SpreadsheetApp.openById(SPREADSHEET_ID_ANALYTICS);

    const aggData      = getSheetData(ssCredits,   SHEET_AGGREGATIONS);
    const ccData       = getSheetData(ssAnalytics, SHEET_CREDIT_CONTROL);
    const upcomingData = getSheetData(ssAnalytics, SHEET_UPCOMING);
    const studentsData = getSheetData(ssAnalytics, SHEET_STUDENTS);
    const scData       = getSheetData(ssAnalytics, SHEET_STUDENTS_COURSES);

    const aggCols = getColMap(aggData[0]);
    const ccCols  = getColMap(ccData[0]);
    const upCols  = getColMap(upcomingData[0]);
    const scCols  = getColMap(scData[0]);

    // ── หา header row จริงของ Students sheet ─────────────
    const stuHeaderIdx = studentsData.findIndex(row =>
      row.some(c => String(c).trim().toLowerCase() === 'student_name')
    );
    if (stuHeaderIdx === -1) throw new Error('ไม่พบ header "student_name" ใน sheet Students');
    const stuCols = getColMap(studentsData[stuHeaderIdx]);
    const stuRows = studentsData.slice(stuHeaderIdx + 1);

    // ── Build active student set ──────────────────────────
    const activeStudents = new Set();
    for (let i = 0; i < stuRows.length; i++) {
      const row  = stuRows[i];
      const name = String(row[stuCols["student_name"]]      || "").trim();
      const rem  = String(row[stuCols["Remaining Credits"]] || "").trim().toUpperCase();
      if (!name) continue;
      if (rem !== "N/A" && rem !== "") {
        activeStudents.add(name);
      }
    }
    Logger.log("Active students count: " + activeStudents.size);

    // ── Build pretest exclusion set ───────────────────────
    // key: "studentName|||packageName" → exclude ถ้า Class Name มีคำว่า Pretest
    const pretestKeys = new Set();
    for (let i = 1; i < scData.length; i++) {
      const row       = scData[i];
      const student   = String(row[scCols["Student Name"]] || "").trim();
      const className = String(row[scCols["Class Name"]]   || "").trim();
      const classSub  = String(row[scCols["Class Subject"]]|| "").trim();
      if (!student || !classSub) continue;
      if (className.toLowerCase().includes("pretest")) {
        pretestKeys.add(`${student}|||${classSub}`);
      }
    }
    Logger.log("Pretest keys excluded: " + pretestKeys.size);

    // ── Pending deductions ────────────────────────────────
    const pendingMap = {};
    for (let i = 1; i < ccData.length; i++) {
      const row      = ccData[i];
      const student  = String(row[ccCols["Student Name"]]    || "").trim();
      const pkg      = String(row[ccCols["Package/Program"]] || "").trim();
      const status   = String(row[ccCols["final_status"]]    || "").trim().toUpperCase();
      const feedback = String(row[ccCols["teacher_feedback"]]|| "").trim();
      const consumed = parseFloat(row[ccCols["credits_consumed"]]) || 0;
      const duration = parseFloat(row[ccCols["session_duration"]]) || 60;
      const sessDate = parseDate(row[ccCols["session_date"]]);

      if (!student || !pkg || !sessDate) continue;
      if (!activeStudents.has(student)) continue;
      if (pretestKeys.has(`${student}|||${pkg}`)) continue;  // skip pretest
      if (sessDate > today) continue;

      if (status === "ENDED" && (feedback === "0" || feedback === "") && consumed === 0) {
        const key = `${student}|||${pkg}`;
        pendingMap[key] = (pendingMap[key] || 0) + (duration / 60);
      }
    }

    // ── Upcoming sessions map ─────────────────────────────
    const upcomingMap = {};
    for (let i = 1; i < upcomingData.length; i++) {
      const row        = upcomingData[i];
      const student    = String(row[upCols["Student Name"]]    || "").trim();
      const pkg        = String(row[upCols["Package/Program"]] || "").trim();
      const sessStatus = String(row[upCols["Session Status"]]  || "").trim().toUpperCase();
      const duration   = parseFloat(row[upCols["Session Duration"]]) || 60;
      const schedDate  = parseDate(row[upCols["Scheduled Date"]]);

      if (!student || !pkg || !schedDate) continue;
      if (!activeStudents.has(student)) continue;
      if (pretestKeys.has(`${student}|||${pkg}`)) continue;  // skip pretest
      if (sessStatus !== "UPCOMING") continue;
      if (schedDate <= today) continue;

      const key = `${student}|||${pkg}`;
      if (!upcomingMap[key]) upcomingMap[key] = [];
      upcomingMap[key].push({ date: schedDate, durationMin: duration });
    }

    for (const key in upcomingMap) {
      upcomingMap[key].sort((a, b) => a.date - b.date);
    }

    // ── Build result จาก Aggregations ────────────────────
    const studentMap = {};

    for (let i = 1; i < aggData.length; i++) {
      const row     = aggData[i];
      const student = String(row[aggCols["Student Name"]]                   || "").trim();
      const parent  = String(row[aggCols["Parent Name"]]                    || "").trim();
      const pkg     = String(row[aggCols["Class Subject"]]                  || "").trim();
      const curRem  = parseFloat(row[aggCols["Current Remaining Credits"]]) || 0;
      const total   = parseFloat(row[aggCols["Current Total Credits"]])     || 0;

      if (!student || !pkg) continue;
      if (!activeStudents.has(student)) continue;
      if (pretestKeys.has(`${student}|||${pkg}`)) continue;  // skip pretest

      const key        = `${student}|||${pkg}`;
      const pendDeduct = Math.round((pendingMap[key] || 0) * 10) / 10;
      const adjRem     = Math.round(Math.max(0, curRem - pendDeduct) * 10) / 10;
      const sessions   = upcomingMap[key] || [];
      const projection = computeProjection(adjRem, sessions, today);

      const pkgObj = {
        name:              pkg,
        currentRemaining:  curRem,
        pendingDeduction:  pendDeduct,
        adjustedRemaining: adjRem,
        totalCredits:      total,
        alertDate:         projection.alertDate,
        exhaustDate:       projection.exhaustDate,
        daysUntilAlert:    projection.daysUntilAlert,
        status:            projection.status,
        projection:        projection.rows,
      };

      if (!studentMap[student]) {
        studentMap[student] = { student, parent, packages: [] };
      }

      // Deduplicate — ถ้า package ชื่อเดียวกันมีอยู่แล้ว เก็บอันที่ totalCredits มากกว่า
      const existingIdx = studentMap[student].packages.findIndex(x => x.name === pkg);
      if (existingIdx !== -1) {
        if (total > studentMap[student].packages[existingIdx].totalCredits) {
          studentMap[student].packages[existingIdx] = pkgObj;
        }
        // ถ้า total น้อยกว่าหรือเท่ากัน → skip ไป
      } else {
        studentMap[student].packages.push(pkgObj);
      }
    }

    const statusOrder = { notify: 0, watch: 1, ok: 2, nodata: 3 };
    return Object.values(studentMap).sort((a, b) => {
      return (statusOrder[worstStatus(a.packages)] || 9) - (statusOrder[worstStatus(b.packages)] || 9);
    });

  } catch (e) {
    Logger.log("getStudentData error: " + e.toString());
    throw new Error("ดึงข้อมูลไม่ได้: " + e.message);
  }
}

// ============================================================
// HELPERS
// ============================================================

function computeProjection(startBal, sessions, today) {
  if (!sessions.length) {
    if (startBal < ALERT_THRESHOLD) {
      return { alertDate: formatDate(today), exhaustDate: null, daysUntilAlert: 0, status: "notify", rows: [] };
    }
    return { alertDate: null, exhaustDate: null, daysUntilAlert: null, status: "nodata", rows: [] };
  }

  let bal = startBal;
  let alertDate = null, exhaustDate = null;
  const rows = [];

  for (const sess of sessions) {
    const credit = Math.round((sess.durationMin / 60) * 100) / 100;
    bal = Math.round((bal - credit) * 100) / 100;

    const flag = [];
    if (!alertDate && bal < ALERT_THRESHOLD)  { alertDate = sess.date;   flag.push("alert"); }
    if (!exhaustDate && bal <= 0)             { exhaustDate = sess.date; flag.push("exhaust"); }

    rows.push({ date: formatDate(sess.date), dur: sess.durationMin, deduct: credit, bal, flag: flag.join(" ") });
    if (exhaustDate && rows.length >= 6) break;
  }

  const daysUntilAlert = alertDate
    ? Math.round((alertDate - today) / (1000 * 60 * 60 * 24))
    : null;

  const status = startBal < ALERT_THRESHOLD ? "notify"
    : alertDate && daysUntilAlert <= NOTIFY_WINDOW_DAYS ? "watch"
    : "ok";

  return {
    alertDate:      alertDate   ? formatDate(alertDate)   : null,
    exhaustDate:    exhaustDate ? formatDate(exhaustDate) : null,
    daysUntilAlert,
    status,
    rows,
  };
}

function worstStatus(packages) {
  if (packages.some(p => p.status === "notify")) return "notify";
  if (packages.some(p => p.status === "watch"))  return "watch";
  if (packages.some(p => p.status === "ok"))     return "ok";
  return "nodata";
}

function getSheetData(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error(`ไม่พบ sheet "${sheetName}"`);
  return sheet.getDataRange().getValues().filter(row => row.some(c => c !== "" && c !== null));
}

function getColMap(headerRow) {
  const map = {};
  headerRow.forEach((h, i) => { if (h) map[String(h).trim()] = i; });
  return map;
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  const d = new Date(val);
  return isNaN(d) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function formatDate(d) {
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd") : null;
}
