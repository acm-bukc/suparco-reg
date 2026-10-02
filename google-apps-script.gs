const SHEET_NAME = "Industrial Visit Registrations";
const MAX_REGISTRATIONS = 40;
// Change to false and redeploy the web app to stop accepting registrations.
const REGISTRATION_OPEN = true;
const ALLOWED_DEPARTMENTS = ["cs", "it", "ai"];
const ALLOWED_SEMESTERS = ["1", "2", "3", "4", "5", "6", "7", "8"];

function doGet(event) {
  const status = getStatus();
  const callback = event && event.parameter && event.parameter.callback;
  const output = JSON.stringify(status);
  if (callback && /^[a-zA-Z_$][\w$]*$/.test(callback)) {
    return ContentService.createTextOutput(`${callback}(${output})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JSON);
}

function doPost(event) {
  if (!REGISTRATION_OPEN) return reply(false, "Registration is currently closed.");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet();
    const records = sheet.getDataRange().getValues().slice(1);
    if (records.length >= MAX_REGISTRATIONS) return reply(false, "Registration is closed because all 40 places are filled.", 0);

    const data = JSON.parse(event.postData.contents);
    const name = normalize(data.name);
    const enrollment = normalize(data.enrollment);
    const cnic = normalizeCnic(data.cnic);
    const semester = String(data.semester || "").trim();
    const department = normalize(data.department);
    const phone = normalizePhone(data.phone);
    const website = String(data.website || "").trim();

    if (website) return reply(false, "Registration could not be submitted.", MAX_REGISTRATIONS - records.length);
    if (!/^[a-zA-Z][a-zA-Z .'-]{2,59}$/.test(name)) return reply(false, "Enter a valid name.");
    if (!/^[a-zA-Z0-9][a-zA-Z0-9\/_-]{2,29}$/.test(enrollment)) return reply(false, "Enter a valid enrollment number.");
    if (!/^\d{5}-\d{7}-\d$/.test(cnic)) return reply(false, "Enter a valid CNIC number.");
    if (!ALLOWED_SEMESTERS.includes(semester)) return reply(false, "Select a valid semester.");
    if (!ALLOWED_DEPARTMENTS.includes(department)) return reply(false, "Select CS, IT, or AI.");
    if (!/^03\d{9}$/.test(phone)) return reply(false, "Enter a valid Pakistani phone number.");

    const duplicate = records.some((row) => normalize(row[1]) === name || normalize(row[2]) === enrollment || normalizeCnic(row[3]) === cnic || normalizePhone(row[6]) === phone);
    if (duplicate) return reply(false, "This name, enrollment, CNIC, or phone number is already registered.", MAX_REGISTRATIONS - records.length);

    sheet.appendRow([new Date(), name, enrollment, cnic, semester, department, phone]);
    return reply(true, "Registration confirmed.", MAX_REGISTRATIONS - records.length - 1);
  } catch (error) {
    return reply(false, "The registration could not be saved. Please try again.");
  } finally {
    lock.releaseLock();
  }
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(["Timestamp", "Name", "Enrollment", "CNIC", "Semester", "Department", "Phone"]);
  return sheet;
}

function getStatus() {
  const records = getSheet().getDataRange().getValues().slice(1);
  const count = records.length;
  return { success: true, count, remaining: Math.max(0, MAX_REGISTRATIONS - count), open: REGISTRATION_OPEN && count < MAX_REGISTRATIONS };
}

function normalize(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function normalizePhone(value) {
  return String(value || "").replace(/[^\d+]/g, "").replace(/^\+92/, "0").replace(/^92/, "0");
}

function normalizeCnic(value) {
  return String(value || "").replace(/[^\d-]/g, "");
}

function reply(success, message, remaining) {
  return ContentService.createTextOutput(JSON.stringify({ success, message, remaining }))
    .setMimeType(ContentService.MimeType.JSON);
}
