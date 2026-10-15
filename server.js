"use strict";

const crypto = require("node:crypto");
const fileSystem = require("node:fs");
const fs = require("node:fs/promises");
const { DatabaseSync } = require("node:sqlite");
const { gunzipSync, gzipSync } = require("node:zlib");
const http = require("node:http");
const os = require("node:os");
const path = require("node:path");
const { promisify } = require("node:util");
const ExcelJS = require("exceljs");
const XLSX = require("@e965/xlsx");

const DOWNLOADS_DIRECTORY = path.join(os.homedir(), "Downloads");
const DEFAULT_DATABASE_PATH = path.join(__dirname, "grade5-portal-database.sqlite");
const APPLICATION_WORKBOOK_PATH = path.join(__dirname, "grade5-portal-database.xlsx");
const DOWNLOADS_DATABASE_PATH = path.join(DOWNLOADS_DIRECTORY, "grade5-portal-database.xlsx");
const LEGACY_DATABASE_PATH = path.join(DOWNLOADS_DIRECTORY, "grade5-portal-database-template.xlsx");
const configuredDatabasePath = process.env.PORTAL_DATABASE_PATH
  ? path.resolve(process.env.PORTAL_DATABASE_PATH)
  : DEFAULT_DATABASE_PATH;
const DATABASE_PATH = /\.(?:xlsx|xlsm?)$/i.test(configuredDatabasePath)
  ? `${configuredDatabasePath}.sqlite`
  : configuredDatabasePath;
const MIGRATION_WORKBOOK_PATH = process.env.PORTAL_MIGRATION_WORKBOOK
  ? path.resolve(process.env.PORTAL_MIGRATION_WORKBOOK)
  : /\.(?:xlsx|xlsm?)$/i.test(configuredDatabasePath)
    ? configuredDatabasePath
    : fileSystem.existsSync(APPLICATION_WORKBOOK_PATH)
      ? APPLICATION_WORKBOOK_PATH
      : fileSystem.existsSync(DOWNLOADS_DATABASE_PATH)
        ? DOWNLOADS_DATABASE_PATH
        : fileSystem.existsSync(LEGACY_DATABASE_PATH)
          ? LEGACY_DATABASE_PATH
          : path.join(__dirname, "grade5-portal-database-template.xlsx");
const INDEX_PATH = path.join(__dirname, "index.html");
const MAX_SF1_FILE_BYTES = 10 * 1024 * 1024;
const MAX_SF2_FILE_BYTES = 10 * 1024 * 1024;
const MAX_SF1_STUDENTS = 1000;
const MAX_PROFILE_PHOTO_BYTES = 2 * 1024 * 1024;
const SF2_TEMPLATE_PATH = path.join(__dirname, "SF2.xlsx");
const SF2_FIRST_MALE_STUDENT_ROW = 13;
const SF2_LAST_MALE_STUDENT_ROW = 62;
const SF2_FIRST_FEMALE_STUDENT_ROW = 64;
const SF2_LAST_FEMALE_STUDENT_ROW = 113;
const SF2_STUDENT_ROWS = [
  ...Array.from(
    { length: SF2_LAST_MALE_STUDENT_ROW - SF2_FIRST_MALE_STUDENT_ROW + 1 },
    (_, index) => SF2_FIRST_MALE_STUDENT_ROW + index,
  ),
  ...Array.from(
    { length: SF2_LAST_FEMALE_STUDENT_ROW - SF2_FIRST_FEMALE_STUDENT_ROW + 1 },
    (_, index) => SF2_FIRST_FEMALE_STUDENT_ROW + index,
  ),
];
const SF2_MAX_STUDENTS_PER_GENDER = 50;
const SF2_FIRST_DAY_COLUMN = 7;
const SF2_LAST_DAY_COLUMN = 31;
const SF2_WEEKDAY_NAMES = ["M", "T", "W", "TH", "F"];
const SF2_REMARKS_COLUMN = 34;
const SF2_ABSENCE_REASONS_MARKER = "\n[Grade Portal absence reasons]\n";
const SF2_STATUS_DATE_MARKER = "\n[Grade Portal status date: ";
const SF2_STATUS_BY_MARK = new Map([
  ["", "Present"],
  ["x", "Absent"],
  ["h", "Half-day"],
  ["t", "Tardy"],
  ["e", "Excused"],
  ["nc", "No Class"],
  ["n/a", "Locked"],
]);
const SF2_MARK_BY_STATUS = new Map([
  ["Present", ""],
  ["Absent", "x"],
  ["Half-day", "h"],
  ["Tardy", "t"],
  ["Excused", "e"],
]);
const PASSWORD_ITERATIONS = 210000;
const PASSWORD_BYTES = 32;
const SESSION_COOKIE = "grade5_session";
const SESSION_DURATION_MS = 8 * 60 * 60 * 1000;
const GRADE_TERMS = ["Term 1", "Term 2", "Term 3"];
const READING_PERIODS = ["BOSY", "MOSY", "EOSY"];
const READING_LANGUAGES = ["English", "Filipino"];
const READING_LEVELS = ["Non-reader", "Frustration", "Instructional", "Independent"];
const ATTENDANCE_STATUSES = ["Present", "Absent", "Tardy", "Half-day", "Excused"];
const GRADE_SUBJECTS = [
  ["filipino", "Filipino"],
  ["english", "English"],
  ["mathematics", "Mathematics"],
  ["science", "Science"],
  ["aralingPanlipunan", "Araling Panlipunan (AP)"],
  ["gmrc", "GMRC/Values Education"],
  ["eppTle", "EPP / TLE"],
  ["musicAndArts", "Music and Arts"],
  ["physicalEducationAndHealth", "Physical Education and Health"],
];
const STUDENT_FIELDS = [
  ["lrn", "LRN"],
  ["lastName", "Last name"],
  ["firstName", "First name"],
  ["middleName", "Middle name"],
  ["gender", "Gender"],
  ["birthdate", "Birthdate"],
  ["mothersMaidenName", "Mother's Maiden name"],
  ["guardianRelationship", "Guardian relationship"],
  ["father", "Father"],
  ["currentResidence", "Current residence"],
  ["religion", "Religion"],
  ["motherTongue", "Mother tongue"],
  ["dialects", "Dialects"],
  ["ethnicities", "Ethnicities"],
  ["teacherUserId", "Teacher User ID"],
];
const pbkdf2 = promisify(crypto.pbkdf2);

class HttpError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}

function headerColumns(worksheet) {
  const headers = new Map();
  worksheet.getRow(1).eachCell((cell, columnNumber) => {
    if (typeof cell.value === "string") {
      headers.set(normalizeHeader(cell.value), columnNumber);
    }
  });
  return headers;
}

function normalizeHeader(name) {
  return name.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function ensureColumn(worksheet, headers, name, aliases = []) {
  for (const candidate of [name, ...aliases]) {
    const column = headers.get(normalizeHeader(candidate));
    if (column) return column;
  }
  const column = worksheet.columnCount + 1;
  const headerCell = worksheet.getRow(1).getCell(column);
  headerCell.value = name;
  const previousHeader = worksheet.getRow(1).getCell(column - 1);
  if (previousHeader.style) headerCell.style = structuredClone(previousHeader.style);
  headers.set(normalizeHeader(name), column);
  return column;
}

function getColumn(headers, name, aliases = []) {
  const column = [name, ...aliases]
    .map(normalizeHeader)
    .map((candidate) => headers.get(candidate))
    .find(Boolean);
  if (!column) {
    throw new Error(`The workbook is missing the "${name}" column.`);
  }
  return column;
}

function cellText(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "object" && "text" in value) return String(value.text).trim();
  return String(value).trim();
}

function getUserSheet(workbook) {
  const sheet = workbook.getWorksheet("Users");
  if (!sheet) throw new Error('The workbook is missing the "Users" sheet.');
  const headers = headerColumns(sheet);
  const passwordColumn = [...headers].find(([name]) => name.startsWith("password"))?.[1];
  if (!passwordColumn) throw new Error('The workbook is missing the "Password" column.');
  return {
    sheet,
    columns: {
      id: getColumn(headers, "User ID"),
      role: getColumn(headers, "Role"),
      name: getColumn(headers, "Full Name", ["Name"]),
      email: getColumn(headers, "Email"),
      password: passwordColumn,
      classId: getColumn(headers, "Class ID"),
      status: getColumn(headers, "Status"),
      created: getColumn(headers, "Created", ["Created At", "CreatedAt"]),
      signIns: ensureColumn(sheet, headers, "Sign-ins"),
      lastSignIn: ensureColumn(sheet, headers, "Last sign-in"),
      emailCheck: ensureColumn(sheet, headers, "Email check"),
    },
  };
}

function setPasswordDocumentation(workbook, passwordColumn) {
  const usersSheet = workbook.getWorksheet("Users");
  usersSheet.getRow(1).getCell(passwordColumn).value = "Password hash (PBKDF2)";
  const readme = workbook.getWorksheet("README");
  if (readme && readme.getCell("A24").value) {
    readme.getCell("A24").value =
      "Account passwords are stored as PBKDF2 hashes. The portal creates user records only when an account is registered.";
  }
}

function findUser(sheet, columns, email) {
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    if (cellText(row.getCell(columns.email).value).toLowerCase() === email) {
      return row;
    }
  }
  return null;
}

function getClassSheet(workbook) {
  const sheet = workbook.getWorksheet("Classes");
  if (!sheet) throw new Error('The workbook is missing the "Classes" sheet.');
  const headers = headerColumns(sheet);
  return {
    sheet,
    columns: {
      id: getColumn(headers, "Class ID"),
      classroom: ensureColumn(sheet, headers, "Classroom number"),
      code: getColumn(headers, "Class code (parents enter this)", ["Parent Join Code"]),
      teacherId: getColumn(headers, "Teacher User ID", ["Teacher ID"]),
      teacherName: getColumn(headers, "Teacher name"),
      students: ensureColumn(sheet, headers, "Students"),
      parents: ensureColumn(sheet, headers, "Parents signed up"),
      grade: ensureColumn(sheet, headers, "Grade"),
      section: ensureColumn(sheet, headers, "Section"),
      schoolYear: ensureColumn(sheet, headers, "School Year"),
      schoolName: ensureColumn(sheet, headers, "Name of School"),
      schoolId: ensureColumn(sheet, headers, "School ID"),
      schoolHead: ensureColumn(sheet, headers, "School Head"),
    },
  };
}

function getProfilePhotoFormat(buffer) {
  if (buffer.length >= 8 && buffer.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))) {
    return { extension: "png", contentType: "image/png" };
  }
  if (buffer.length >= 3 && buffer[0] === 0xff && buffer[1] === 0xd8 && buffer[2] === 0xff) {
    return { extension: "jpg", contentType: "image/jpeg" };
  }
  if (buffer.length >= 12 && buffer.toString("ascii", 0, 4) === "RIFF"
    && buffer.toString("ascii", 8, 12) === "WEBP") {
    return { extension: "webp", contentType: "image/webp" };
  }
  throw new HttpError(400, "Upload a valid JPEG, PNG, or WebP profile picture.");
}

function getStudentSheet(workbook) {
  const sheet = workbook.getWorksheet("Students");
  if (!sheet) throw new Error('The workbook is missing the "Students" sheet.');
  const headers = headerColumns(sheet);
  const required = {
    id: getColumn(headers, "Student ID"),
    name: ensureColumn(sheet, headers, "Student name"),
    classId: getColumn(headers, "Class ID"),
    parentUserId: ensureColumn(sheet, headers, "Parent User ID"),
  };
  const columns = { ...required };

  for (const [key, label] of STUDENT_FIELDS) {
    const normalizedLabel = label.toLowerCase();
    let column = headers.get(normalizeHeader(label));
    if (!column && key === "teacherUserId") {
      column = headers.get(normalizeHeader("Teacher ID"));
    }
    if (!column) {
      column = ensureColumn(sheet, headers, label);
    }
    columns[key] = column;
  }

  return { sheet, columns };
}

function getOwnedStudents(sheet, columns, session) {
  const students = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const ownerId = cellText(row.getCell(columns.teacherUserId).value);
    if (cellText(row.getCell(columns.classId).value) !== session.classId
      || (ownerId && ownerId !== session.userId)) continue;

    const student = {
      id: cellText(row.getCell(columns.id).value),
      name: cellText(row.getCell(columns.name).value),
      parentUserId: cellText(row.getCell(columns.parentUserId).value),
    };
    if (!student.id || !student.name) continue;
    for (const [key] of STUDENT_FIELDS) {
      if (key === "teacherUserId") continue;
      const value = row.getCell(columns[key]).value;
      student[key] = key === "birthdate" && value instanceof Date
        ? value.toISOString().slice(0, 10)
        : cellText(value);
    }
    students.push(student);
  }
  return students;
}

function getGradeSheet(workbook) {
  let sheet = workbook.getWorksheet("Term Grades");
  if (!sheet) {
    sheet = workbook.addWorksheet("Term Grades");
    sheet.addRow([
      "Grade ID",
      "Student ID",
      "Class ID",
      "Teacher User ID",
      "Term",
      "Subject",
      "Grade",
      "Updated",
    ]);
    sheet.getRow(1).font = { bold: true };
  }
  const headers = headerColumns(sheet);
  return {
    sheet,
    columns: {
      id: getColumn(headers, "Grade ID"),
      studentId: getColumn(headers, "Student ID"),
      classId: getColumn(headers, "Class ID"),
      teacherUserId: getColumn(headers, "Teacher User ID", ["Teacher ID"]),
      term: getColumn(headers, "Term"),
      subject: getColumn(headers, "Subject", ["Subject Key"]),
      grade: getColumn(headers, "Grade"),
      updated: getColumn(headers, "Updated", ["Updated At"]),
    },
  };
}

function getGradeLockSheet(workbook) {
  let sheet = workbook.getWorksheet("Term Grade Locks");
  if (!sheet) {
    sheet = workbook.addWorksheet("Term Grade Locks");
    sheet.addRow(["Class ID", "Teacher User ID", "Term", "Student ID", "Locked", "Updated"]);
    sheet.getRow(1).font = { bold: true };
  }
  const headers = headerColumns(sheet);
  return {
    sheet,
    columns: {
      classId: getColumn(headers, "Class ID"),
      teacherUserId: getColumn(headers, "Teacher User ID"),
      term: getColumn(headers, "Term"),
      studentId: ensureColumn(sheet, headers, "Student ID"),
      locked: getColumn(headers, "Locked"),
      updated: getColumn(headers, "Updated"),
    },
  };
}

function getReadingAssessmentSheet(workbook) {
  let sheet = workbook.getWorksheet("Reading Assessments");
  if (!sheet) {
    sheet = workbook.addWorksheet("Reading Assessments");
    sheet.addRow([
      "Assessment ID",
      "Student ID",
      "Class ID",
      "Teacher User ID",
      "Assessment Period",
      "Language",
      "Reading Level",
      "Assessed Date",
      "Updated",
    ]);
    sheet.getRow(1).font = { bold: true };
  }
  const headers = headerColumns(sheet);
  return {
    sheet,
    columns: {
      id: getColumn(headers, "Assessment ID"),
      studentId: getColumn(headers, "Student ID"),
      classId: getColumn(headers, "Class ID"),
      teacherUserId: getColumn(headers, "Teacher User ID"),
      period: getColumn(headers, "Assessment Period"),
      language: getColumn(headers, "Language"),
      level: getColumn(headers, "Reading Level"),
      assessed: getColumn(headers, "Assessed Date"),
      attempt: ensureColumn(sheet, headers, "Attempt Number"),
      updated: getColumn(headers, "Updated"),
    },
  };
}

function getAttendanceSheet(workbook) {
  let sheet = workbook.getWorksheet("Attendance Records");
  if (!sheet) {
    const legacySheet = workbook.getWorksheet("Attendance");
    const legacyHeaders = legacySheet ? headerColumns(legacySheet) : new Map();
    if (legacySheet && legacyHeaders.has(normalizeHeader("Attendance ID"))) {
      legacySheet.name = "Attendance Records";
      sheet = legacySheet;
    }
  }
  if (!sheet) {
    sheet = workbook.addWorksheet("Attendance Records");
    sheet.addRow([
      "Attendance ID",
      "Student ID",
      "Class ID",
      "Teacher User ID",
      "Attendance Date",
      "Status",
      "Reason",
      "Updated",
      "Attendance Key",
    ]);
    sheet.getRow(1).font = { bold: true };
  }
  const headers = headerColumns(sheet);
  return {
    sheet,
    columns: {
      id: getColumn(headers, "Attendance ID"),
      studentId: getColumn(headers, "Student ID"),
      classId: getColumn(headers, "Class ID"),
      teacherUserId: getColumn(headers, "Teacher User ID"),
      date: getColumn(headers, "Attendance Date"),
      status: getColumn(headers, "Status"),
      reason: ensureColumn(sheet, headers, "Reason"),
      updated: getColumn(headers, "Updated"),
      key: ensureColumn(sheet, headers, "Attendance Key"),
    },
  };
}

function getAttendanceNotesSheet(workbook) {
  let sheet = workbook.getWorksheet("Attendance Monthly Notes");
  if (!sheet) {
    sheet = workbook.addWorksheet("Attendance Monthly Notes");
    sheet.addRow([
      "Note ID",
      "Student ID",
      "Class ID",
      "Teacher User ID",
      "Month",
      "Remarks",
      "Updated",
      "Attendance Key",
    ]);
    sheet.getRow(1).font = { bold: true };
  }
  const headers = headerColumns(sheet);
  return {
    sheet,
    columns: {
      id: getColumn(headers, "Note ID"),
      studentId: getColumn(headers, "Student ID"),
      classId: getColumn(headers, "Class ID"),
      teacherUserId: getColumn(headers, "Teacher User ID"),
      month: getColumn(headers, "Month"),
      remarks: getColumn(headers, "Remarks"),
      statusDate: ensureColumn(sheet, headers, "Status Date"),
      updated: getColumn(headers, "Updated"),
      key: ensureColumn(sheet, headers, "Attendance Key"),
    },
  };
}

function getAttendanceCalendarSheet(workbook) {
  let sheet = workbook.getWorksheet("Attendance Calendar");
  if (!sheet) {
    sheet = workbook.addWorksheet("Attendance Calendar");
    sheet.addRow([
      "Calendar ID",
      "Class ID",
      "Teacher User ID",
      "Attendance Date",
      "Day Type",
      "Updated",
      "Calendar Key",
    ]);
    sheet.getRow(1).font = { bold: true };
  }
  const headers = headerColumns(sheet);
  return {
    sheet,
    columns: {
      id: getColumn(headers, "Calendar ID"),
      classId: getColumn(headers, "Class ID"),
      teacherUserId: getColumn(headers, "Teacher User ID"),
      date: getColumn(headers, "Attendance Date"),
      dayType: getColumn(headers, "Day Type"),
      updated: getColumn(headers, "Updated"),
      key: getColumn(headers, "Calendar Key"),
    },
  };
}

function findClass(workbook, role, value) {
  const { sheet, columns } = getClassSheet(workbook);
  const lookupColumn = role === "teacher" ? columns.classroom : columns.code;

  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    if (cellText(row.getCell(lookupColumn).value).toLowerCase() === value.toLowerCase()) {
      return {
        row,
        id: cellText(row.getCell(columns.id).value),
        code: cellText(row.getCell(columns.code).value),
        columns,
      };
    }
  }

  if (role === "teacher") return null;
  throw new HttpError(400, "That class code is not registered yet. Ask your teacher to sign up first and share the class code.");
}

async function hashPassword(password) {
  const salt = crypto.randomBytes(16);
  const derivedKey = await pbkdf2(password, salt, PASSWORD_ITERATIONS, PASSWORD_BYTES, "sha256");
  return `pbkdf2$${PASSWORD_ITERATIONS}$${salt.toString("base64")}$${derivedKey.toString("base64")}`;
}

async function verifyPassword(password, storedPassword) {
  const [, iterationsText, saltText, keyText] = storedPassword.split("$");
  if (!storedPassword.startsWith("pbkdf2$")) return false;
  const iterations = Number(iterationsText);
  if (!Number.isInteger(iterations) || iterations < 100000 || iterations > 1000000) return false;

  const salt = Buffer.from(saltText || "", "base64");
  const expected = Buffer.from(keyText || "", "base64");
  if (salt.length !== 16 || expected.length !== PASSWORD_BYTES) return false;

  const actual = await pbkdf2(password, salt, iterations, expected.length, "sha256");
  return crypto.timingSafeEqual(actual, expected);
}

function nextUserId(sheet, idColumn, role) {
  const prefix = role === "teacher" ? "T" : "P";
  let highest = 0;
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const id = cellText(sheet.getRow(rowNumber).getCell(idColumn).value);
    const match = id.match(new RegExp(`^${prefix}(\\d+)$`, "i"));
    if (match) highest = Math.max(highest, Number(match[1]));
  }
  return `${prefix}${String(highest + 1).padStart(3, "0")}`;
}

function nextClassId(sheet, idColumn) {
  let highest = 0;
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const id = cellText(sheet.getRow(rowNumber).getCell(idColumn).value);
    const match = id.match(/^C(\d+)$/i);
    if (match) highest = Math.max(highest, Number(match[1]));
  }
  return `C${String(highest + 1).padStart(3, "0")}`;
}

function nextStudentId(sheet, idColumn) {
  let highest = 0;
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const id = cellText(sheet.getRow(rowNumber).getCell(idColumn).value);
    const match = id.match(/^S(\d+)$/i);
    if (match) highest = Math.max(highest, Number(match[1]));
  }
  return `S${String(highest + 1).padStart(3, "0")}`;
}

function nextGradeId(sheet, idColumn) {
  let highest = 0;
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const id = cellText(sheet.getRow(rowNumber).getCell(idColumn).value);
    const match = id.match(/^G(\d+)$/i);
    if (match) highest = Math.max(highest, Number(match[1]));
  }
  return `G${String(highest + 1).padStart(5, "0")}`;
}

function nextReadingAssessmentId(sheet, idColumn) {
  let highest = 0;
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const id = cellText(sheet.getRow(rowNumber).getCell(idColumn).value);
    const match = id.match(/^R(\d+)$/i);
    if (match) highest = Math.max(highest, Number(match[1]));
  }
  return `R${String(highest + 1).padStart(5, "0")}`;
}

function createTeacherClass(workbook, classroomNumber, teacherId, teacherName) {
  const { sheet, columns } = getClassSheet(workbook);
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const existing = sheet.getRow(rowNumber);
    if (cellText(existing.getCell(columns.classroom).value).toLowerCase() === classroomNumber.toLowerCase()) {
      throw new HttpError(409, "That classroom number is already registered. Ask your teacher for the class code.");
    }
  }

  let code;
  do {
    code = `G5-${crypto.randomBytes(4).toString("hex").toUpperCase()}`;
  } while (Array.from({ length: Math.max(0, sheet.rowCount - 1) }, (_, index) =>
    cellText(sheet.getRow(index + 2).getCell(columns.code).value).toLowerCase(),
  ).includes(code.toLowerCase()));

  const classId = nextClassId(sheet, columns.id);
  const values = [];
  values[columns.id] = classId;
  values[columns.classroom] = classroomNumber;
  values[columns.code] = code;
  values[columns.teacherId] = teacherId;
  values[columns.teacherName] = teacherName;
  values[columns.students] = 0;
  values[columns.parents] = 0;
  writeRow(sheet, [columns.id, columns.classroom, columns.code], values);
  return { id: classId, code };
}

function incrementCell(cell) {
  if (
    cell.value
    && typeof cell.value === "object"
    && ("formula" in cell.value || "sharedFormula" in cell.value)
  ) return;
  const value = cell.value;
  const current = typeof value === "object" && value !== null && "result" in value
    ? Number(value.result)
    : Number(value);
  cell.value = (Number.isFinite(current) ? current : 0) + 1;
}

function firstAvailableRow(sheet, keyColumns) {
  for (let rowNumber = 2; rowNumber <= sheet.rowCount + 1; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    if (keyColumns.every((column) => !cellText(row.getCell(column).value))) return row;
  }
  throw new Error(`The "${sheet.name}" sheet has no empty rows available.`);
}

function writeRow(sheet, keyColumns, values) {
  const row = firstAvailableRow(sheet, keyColumns);
  for (const [column, value] of Object.entries(values)) {
    row.getCell(Number(column)).value = value;
  }
  return row;
}

function logSignIn(workbook, entry) {
  const sheet = workbook.getWorksheet("SignInLog");
  if (!sheet) throw new Error('The workbook is missing the "SignInLog" sheet.');
  const headers = headerColumns(sheet);
  const values = [];
  values[getColumn(headers, "Timestamp")] = new Date();
  values[getColumn(headers, "Email")] = entry.email;
  values[getColumn(headers, "Role chosen", ["Role"])] = entry.role === "teacher" ? "Teacher" : "Parent";
  values[getColumn(headers, "Result", ["Success"])] = entry.result;
  values[getColumn(headers, "User ID")] = entry.userId || "";
  values[getColumn(headers, "Note", ["User Agent"])] = entry.note;
  writeRow(
    sheet,
    [getColumn(headers, "Timestamp"), getColumn(headers, "Email")],
    values,
  );
}

function getClassCode(workbook, classId) {
  const { sheet, columns } = getClassSheet(workbook);
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    if (cellText(row.getCell(columns.id).value) === classId) {
      return cellText(row.getCell(columns.code).value);
    }
  }
  return "";
}

function encodeWorkbookState(workbook) {
  return gzipSync(Buffer.from(JSON.stringify(workbook.model)));
}

function decodeWorkbookState(compressedState) {
  const model = JSON.parse(gunzipSync(compressedState).toString("utf8"));
  for (const sheet of model.worksheets || []) {
    for (const row of sheet.rows || []) {
      for (const cell of row.cells || []) {
        if (cell.type === ExcelJS.ValueType.Date && typeof cell.value === "string") {
          cell.value = new Date(cell.value);
        }
      }
    }
  }
  const workbook = new ExcelJS.Workbook();
  workbook.model = model;
  return workbook;
}

function openPortalDatabase(databasePath) {
  const database = new DatabaseSync(databasePath);
  database.exec(`
    PRAGMA journal_mode = WAL;
    PRAGMA foreign_keys = ON;
    PRAGMA busy_timeout = 5000;
    CREATE TABLE IF NOT EXISTS portal_state (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      schema_version INTEGER NOT NULL,
      workbook_state BLOB NOT NULL,
      updated_at TEXT NOT NULL
    );
  `);
  return database;
}

function readPortalDatabaseWorkbook(databasePath) {
  const database = new DatabaseSync(databasePath);
  try {
    const result = database.prepare(
      "SELECT workbook_state FROM portal_state WHERE id = 1",
    ).get();
    if (!result) throw new Error(`The SQLite database ${path.basename(databasePath)} has not been initialized.`);
    return decodeWorkbookState(result.workbook_state);
  } finally {
    database.close();
  }
}

function writePortalDatabaseWorkbook(databasePath, workbook) {
  const database = new DatabaseSync(databasePath);
  try {
    const result = database.prepare(
      "UPDATE portal_state SET workbook_state = ?, updated_at = ? WHERE id = 1",
    ).run(encodeWorkbookState(workbook), new Date().toISOString());
    if (result.changes !== 1) throw new Error("The SQLite portal database has not been initialized.");
  } finally {
    database.close();
  }
}

async function readJson(request, maxBytes = 16384) {
  let body = "";
  for await (const chunk of request) {
    body += chunk;
    if (Buffer.byteLength(body, "utf8") > maxBytes) {
      throw new HttpError(413, "The request is too large.");
    }
  }

  try {
    const data = JSON.parse(body);
    if (!data || typeof data !== "object" || Array.isArray(data)) throw new Error("Invalid body.");
    return data;
  } catch {
    throw new HttpError(400, "Enter valid account information.");
  }
}

function sendJson(response, status, data) {
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    "X-Content-Type-Options": "nosniff",
  });
  response.end(JSON.stringify(data));
}

async function readProfilePhoto(request) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > MAX_PROFILE_PHOTO_BYTES) {
      throw new HttpError(413, "Profile pictures must be 2 MB or smaller.");
    }
    chunks.push(chunk);
  }
  if (!size) throw new HttpError(400, "Choose a profile picture to upload.");
  const buffer = Buffer.concat(chunks);
  const format = getProfilePhotoFormat(buffer);
  const declaredType = request.headers["content-type"]?.split(";")[0].trim().toLowerCase();
  if (declaredType !== format.contentType) {
    throw new HttpError(400, "The selected image does not match its file type.");
  }
  return { buffer, ...format };
}

function validateAccount(data, registration) {
  const name = typeof data.name === "string" ? data.name.trim() : "";
  const email = typeof data.email === "string" ? data.email.trim().toLowerCase() : "";
  const password = typeof data.password === "string" ? data.password : "";
  const role = data.role;
  const extra = typeof data.extra === "string" ? data.extra.trim() : "";

  if (!/^\S+@\S+\.\S+$/.test(email) || email.length > 254) {
    throw new HttpError(400, "Enter a valid email address.");
  }
  if (password.length < 6 || password.length > 256) {
    throw new HttpError(400, "Use a password between 6 and 256 characters.");
  }
  if (role !== "teacher" && role !== "parent") {
    throw new HttpError(400, "Choose Teacher or Parent.");
  }
  if (registration && (!name || name.length > 120 || !extra || extra.length > 100)) {
    throw new HttpError(400, "Enter your name and classroom information.");
  }

  return { name, email, password, role, extra };
}

function validateStudent(data) {
  const student = {};
  for (const [key, label] of STUDENT_FIELDS) {
    if (key === "teacherUserId") continue;
    const value = data[key];
    student[key] = typeof value === "string" ? value.trim() : "";
    if (student[key].length > 200) throw new HttpError(400, `${label} must be 200 characters or fewer.`);
  }

  if (!/^\d{12}$/.test(student.lrn)) {
    throw new HttpError(400, "Enter the learner's 12-digit LRN.");
  }
  if (!student.lastName || !student.firstName) {
    throw new HttpError(400, "Enter the learner's first and last name.");
  }
  if (!["Male", "Female"].includes(student.gender)) {
    throw new HttpError(400, "Choose Male or Female for gender.");
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(student.birthdate)) {
    throw new HttpError(400, "Enter a valid birthdate.");
  }
  const birthdate = new Date(`${student.birthdate}T00:00:00.000Z`);
  if (Number.isNaN(birthdate.getTime()) || birthdate.toISOString().slice(0, 10) !== student.birthdate) {
    throw new HttpError(400, "Enter a valid birthdate.");
  }
  if (birthdate > new Date(new Date().toISOString().slice(0, 10))) {
    throw new HttpError(400, "Birthdate cannot be in the future.");
  }

  return { ...student, birthdate };
}

function validateGradeTerm(value) {
  if (!GRADE_TERMS.includes(value)) {
    throw new HttpError(400, "Choose a valid term (Term 1 through Term 3).");
  }
  return value;
}

function validateStudentGrades(data) {
  if (!data || typeof data !== "object" || Array.isArray(data)) {
    throw new HttpError(400, "Enter grades for the selected student.");
  }
  const grades = {};
  for (const [key, label] of GRADE_SUBJECTS) {
    const value = data[key];
    if (value === "" || value === null || value === undefined) {
      grades[key] = null;
      continue;
    }
    const grade = typeof value === "number" ? value
      : typeof value === "string" && value.trim() ? Number(value)
        : Number.NaN;
    if (!Number.isFinite(grade) || grade < 0 || grade > 100) {
      throw new HttpError(400, `${label} grade must be a number from 0 to 100.`);
    }
    grades[key] = Math.round(grade * 100) / 100;
  }
  return grades;
}

function findOwnedStudent(sheet, columns, session, studentId) {
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    if (
      cellText(row.getCell(columns.id).value) === studentId
      && cellText(row.getCell(columns.classId).value) === session.classId
      && (!cellText(row.getCell(columns.teacherUserId).value)
        || cellText(row.getCell(columns.teacherUserId).value) === session.userId)
    ) {
      return row;
    }
  }
  throw new HttpError(404, "That student is not in your class roster.");
}

function sf1CellText(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function parseSF1Birthdate(value) {
  let year;
  let month;
  let day;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    year = value.getUTCFullYear();
    month = value.getUTCMonth() + 1;
    day = value.getUTCDate();
  } else if (typeof value === "number") {
    const excelDate = XLSX.SSF.parse_date_code(value);
    if (!excelDate) return "";
    ({ y: year, m: month, d: day } = excelDate);
  } else {
    const dateText = sf1CellText(value);
    let parts = dateText.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (parts) {
      [, year, month, day] = parts;
    } else {
      parts = dateText.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
      if (!parts) return "";
      [, month, day, year] = parts;
    }
    year = Number(year);
    month = Number(month);
    day = Number(day);
  }

  const date = new Date(Date.UTC(year, month - 1, day));
  return Number.isNaN(date.getTime()) || date.getUTCFullYear() !== year
    || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day
    ? ""
    : date.toISOString().slice(0, 10);
}

function parseSF1Workbook(buffer) {
  const isLegacyExcel = buffer.length >= 8
    && buffer.subarray(0, 8).equals(Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]));
  const isOpenXmlExcel = buffer.length >= 4
    && buffer[0] === 0x50
    && buffer[1] === 0x4b
    && ((buffer[2] === 0x03 && buffer[3] === 0x04)
      || (buffer[2] === 0x05 && buffer[3] === 0x06)
      || (buffer[2] === 0x07 && buffer[3] === 0x08));
  if (!isLegacyExcel && !isOpenXmlExcel) {
    throw new HttpError(400, "The selected file is not a supported Excel workbook. Choose an Excel SF1 file.");
  }
  let workbook;
  try {
    workbook = XLSX.read(buffer, { type: "buffer", cellDates: true, dense: true });
  } catch {
    throw new HttpError(400, "The selected file could not be read. Choose a valid SF1 .xls or .xlsx file.");
  }

  for (const sheetName of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      raw: true,
      defval: "",
    });
    if (rows.length > 5000) {
      throw new HttpError(400, "The SF1 sheet has too many rows to import.");
    }

    const headerIndex = rows.slice(0, 30).findIndex((row) =>
      sf1CellText(row[0]).toLowerCase() === "lrn"
      && sf1CellText(row[2]).toLowerCase().startsWith("name")
      && sf1CellText(row[6]).toLowerCase().includes("sex")
      && sf1CellText(row[7]).toLowerCase().includes("birth"),
    );
    if (headerIndex < 0) continue;

    const students = [];
    const skipped = [];
    const addIssue = (rowNumber, reason) => skipped.push({ row: rowNumber, reason });

    for (let index = headerIndex + 1; index < rows.length; index += 1) {
      const row = rows[index];
      const lrnValue = sf1CellText(row[0]).replace(/^'/, "");
      const rawName = sf1CellText(row[2]);
      if (/^<={2,}/.test(rawName) || /^TOTAL\b/i.test(rawName)) continue;
      if (!rawName && !/^\d{1,12}$/.test(lrnValue)) continue;
      if (!lrnValue && !rawName) continue;
      if (students.length + skipped.length >= MAX_SF1_STUDENTS) {
        throw new HttpError(400, `The SF1 file contains more than ${MAX_SF1_STUDENTS} student rows.`);
      }

      const nameParts = rawName.split(",").map((part) => part.trim());
      const genderValue = sf1CellText(row[6]).toLowerCase();
      const address = row.slice(15, 26).map(sf1CellText).filter(Boolean).join(", ");
      const student = {
        lrn: /^\d+$/.test(lrnValue) ? lrnValue.padStart(12, "0") : lrnValue,
        lastName: nameParts.length > 1 ? nameParts[0] : "",
        firstName: nameParts.length > 1 ? nameParts[1] : "",
        middleName: nameParts.slice(2).filter((part) => part !== "-" && part !== "—").join(" "),
        gender: genderValue === "m" ? "Male" : genderValue === "f" ? "Female" : row[6],
        birthdate: parseSF1Birthdate(row[7]),
        motherTongue: sf1CellText(row[11]),
        ethnicities: sf1CellText(row[13]),
        religion: sf1CellText(row[14]),
        currentResidence: address,
        father: sf1CellText(row[27]),
        mothersMaidenName: sf1CellText(row[31]),
        guardianRelationship: sf1CellText(row[40]),
      };

      try {
        students.push({ row: index + 1, student: validateStudent(student) });
      } catch (error) {
        if (!(error instanceof HttpError)) throw error;
        addIssue(index + 1, error.message);
      }
    }

    return { students, skipped };
  }

  throw new HttpError(400, "No SF1 student register was found. Use the SF1 template with LRN, NAME, Sex, and BIRTH DATE columns.");
}

async function readSF1Upload(request) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > MAX_SF1_FILE_BYTES) {
      throw new HttpError(413, "SF1 files must be 10 MB or smaller.");
    }
    chunks.push(chunk);
  }
  if (!size) throw new HttpError(400, "Choose an SF1 file to import.");
  return Buffer.concat(chunks, size);
}

async function readSF2Upload(request) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > MAX_SF2_FILE_BYTES) {
      throw new HttpError(413, "SF2 files must be 10 MB or smaller.");
    }
    chunks.push(chunk);
  }
  if (!size) throw new HttpError(400, "Choose an SF2 file to import.");
  const buffer = Buffer.concat(chunks, size);
  if (buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4b) {
    throw new HttpError(400, "Upload the SF2 Excel workbook in .xlsx format.");
  }
  return buffer;
}

function createServer(options = {}) {
  const databasePath = options.databasePath || DATABASE_PATH;
  const migrationWorkbookPath = options.migrationWorkbookPath || MIGRATION_WORKBOOK_PATH;
  const profilePhotoDirectories = databasePath === DEFAULT_DATABASE_PATH
    ? [
      `${databasePath}.profiles`,
      `${APPLICATION_WORKBOOK_PATH}.profiles`,
      `${DOWNLOADS_DATABASE_PATH}.profiles`,
      `${LEGACY_DATABASE_PATH}.profiles`,
    ]
    : [`${databasePath}.profiles`];
  const profilePhotoDirectory = profilePhotoDirectories.find((directory) =>
    fileSystem.existsSync(directory)) || profilePhotoDirectories[0];
  let workbookQueue = Promise.resolve();
  const sessions = new Map();
  let database = null;
  let databaseInitialization = null;

  async function initializeDatabase() {
    if (databaseInitialization) return databaseInitialization;
    databaseInitialization = (async () => {
      await fs.mkdir(path.dirname(databasePath), { recursive: true });
      database = openPortalDatabase(databasePath);
      const existingState = database.prepare(
        "SELECT workbook_state FROM portal_state WHERE id = 1",
      ).get();
      if (existingState) {
        decodeWorkbookState(existingState.workbook_state);
        return;
      }
      if (!fileSystem.existsSync(migrationWorkbookPath)) {
        throw new Error(
          `The SQLite database is empty and the migration workbook was not found: ${migrationWorkbookPath}`,
        );
      }
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(migrationWorkbookPath);
      getUserSheet(workbook);
      if (!workbook.getWorksheet("Classes") || !workbook.getWorksheet("SignInLog")) {
        throw new Error("The migration workbook is missing required portal worksheets.");
      }
      database.prepare(
        "INSERT INTO portal_state (id, schema_version, workbook_state, updated_at) VALUES (1, 1, ?, ?)",
      ).run(encodeWorkbookState(workbook), new Date().toISOString());
      console.log(`Migrated portal data from ${path.resolve(migrationWorkbookPath)} into SQLite.`);
    })();
    return databaseInitialization;
  }

  async function readDatabaseWorkbook() {
    await initializeDatabase();
    const result = database.prepare(
      "SELECT workbook_state FROM portal_state WHERE id = 1",
    ).get();
    if (!result) throw new Error("The SQLite portal database has no saved state.");
    return decodeWorkbookState(result.workbook_state);
  }

  async function saveDatabaseWorkbook(workbook) {
    await initializeDatabase();
    const result = database.prepare(
      "UPDATE portal_state SET workbook_state = ?, updated_at = ? WHERE id = 1",
    ).run(encodeWorkbookState(workbook), new Date().toISOString());
    if (result.changes !== 1) throw new Error("The SQLite portal database has no saved state.");
  }

  function serializeWorkbook(operation) {
    const result = workbookQueue.then(operation);
    workbookQueue = result.then(() => undefined, () => undefined);
    return result;
  }

  async function checkWorkbook() {
    const workbook = await readDatabaseWorkbook();
    getUserSheet(workbook);
    if (!workbook.getWorksheet("Classes")) {
      throw new Error('The workbook is missing the "Classes" sheet.');
    }
    if (!workbook.getWorksheet("SignInLog")) {
      throw new Error('The workbook is missing the "SignInLog" sheet.');
    }
  }

  function createSessionCookie(response, sessionId, request) {
    const secure = request.socket.encrypted ? "; Secure" : "";
    response.setHeader(
      "Set-Cookie",
      `${SESSION_COOKIE}=${sessionId}; Path=/; HttpOnly; SameSite=Strict; Max-Age=${SESSION_DURATION_MS / 1000}${secure}`,
    );
  }

  function getSession(request) {
    const cookieHeader = request.headers.cookie || "";
    const cookie = cookieHeader.split(";").map((part) => part.trim())
      .find((part) => part.startsWith(`${SESSION_COOKIE}=`));
    if (!cookie) return null;
    const sessionId = cookie.slice(SESSION_COOKIE.length + 1);
    const session = sessions.get(sessionId);
    if (!session) return null;
    if (session.expiresAt <= Date.now()) {
      sessions.delete(sessionId);
      return null;
    }
    session.expiresAt = Date.now() + SESSION_DURATION_MS;
    return { sessionId, ...session };
  }

  function requireTeacherSession(request) {
    const session = getSession(request);
    if (!session) throw new HttpError(401, "Your session has expired. Sign in again.");
    if (session.role !== "teacher") throw new HttpError(403, "Only teacher accounts can manage student records.");
    return session;
  }

  function requireParentSession(request) {
    const session = getSession(request);
    if (!session) throw new HttpError(401, "Your session has expired. Sign in again.");
    if (session.role !== "parent") throw new HttpError(403, "Only parent accounts can access the Parent Portal.");
    return session;
  }

  function openSession(response, request, user) {
    const sessionId = crypto.randomBytes(32).toString("hex");
    sessions.set(sessionId, {
      userId: user.id,
      name: user.name,
      role: user.role,
      classId: user.classId,
      classCode: user.classCode || "",
      expiresAt: Date.now() + SESSION_DURATION_MS,
    });
    createSessionCookie(response, sessionId, request);
  }

  async function register(data) {
    const account = validateAccount(data, true);
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet, columns } = getUserSheet(workbook);
      if (findUser(sheet, columns, account.email)) {
        throw new HttpError(409, "That email already has an account. Try signing in.");
      }

      const userId = nextUserId(sheet, columns.id, account.role);
      const classInfo = account.role === "teacher"
        ? createTeacherClass(workbook, account.extra, userId, account.name)
        : findClass(workbook, account.role, account.extra);
      const passwordHash = await hashPassword(account.password);
      const rowValues = [];
      rowValues[columns.id] = userId;
      rowValues[columns.role] = account.role === "teacher" ? "Teacher" : "Parent";
      rowValues[columns.name] = account.name;
      rowValues[columns.email] = account.email;
      rowValues[columns.password] = passwordHash;
      rowValues[columns.classId] = classInfo.id;
      rowValues[columns.status] = "Active";
      rowValues[columns.created] = new Date();
      rowValues[columns.signIns] = 0;
      rowValues[columns.emailCheck] = "OK";
      writeRow(sheet, [columns.id, columns.email], rowValues);
      setPasswordDocumentation(workbook, columns.password);

      if (account.role === "parent") {
        incrementCell(classInfo.row.getCell(classInfo.columns.parents));
      }

      await saveDatabaseWorkbook(workbook);
      return {
        user: {
          id: userId,
          name: account.name,
          role: account.role,
          classId: classInfo.id,
          classCode: account.role === "teacher" ? classInfo.code : undefined,
        },
      };
    });
  }

  async function login(data) {
    const account = validateAccount(data, false);
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet, columns } = getUserSheet(workbook);
      const row = findUser(sheet, columns, account.email);
      const userId = row ? cellText(row.getCell(columns.id).value) : "";
      if (!row || cellText(row.getCell(columns.status).value).toLowerCase() !== "active") {
        logSignIn(workbook, {
          email: account.email,
          role: account.role,
          result: "Failed",
          userId,
          note: row ? "Inactive account" : "Account not found",
        });
        await saveDatabaseWorkbook(workbook);
        throw new HttpError(401, "Email or password doesn't match. Check both and try again.");
      }

      const storedPassword = cellText(row.getCell(columns.password).value);
      if (!(await verifyPassword(account.password, storedPassword))) {
        logSignIn(workbook, {
          email: account.email,
          role: account.role,
          result: "Failed",
          userId,
          note: "Password mismatch",
        });
        await saveDatabaseWorkbook(workbook);
        throw new HttpError(401, "Email or password doesn't match. Check both and try again.");
      }

      const storedRole = cellText(row.getCell(columns.role).value).toLowerCase();
      if (storedRole !== account.role) {
        logSignIn(workbook, {
          email: account.email,
          role: account.role,
          result: "Failed",
          userId,
          note: "Role mismatch",
        });
        await saveDatabaseWorkbook(workbook);
        throw new HttpError(
          403,
          `This email is a ${storedRole} account. Choose ${storedRole === "teacher" ? "Teacher" : "Parent"} above.`,
        );
      }

      incrementCell(row.getCell(columns.signIns));
      row.getCell(columns.lastSignIn).value = new Date();
      logSignIn(workbook, {
        email: account.email,
        role: account.role,
        result: "Success",
        userId,
        note: "",
      });
      setPasswordDocumentation(workbook, columns.password);
      await saveDatabaseWorkbook(workbook);

      return {
        user: {
          id: userId,
          name: cellText(row.getCell(columns.name).value),
          role: account.role,
          classId: cellText(row.getCell(columns.classId).value),
          classCode: account.role === "teacher"
            ? getClassCode(workbook, cellText(row.getCell(columns.classId).value))
            : undefined,
        },
      };
    });
  }

  async function listStudents(session) {
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      const { sheet: userSheet, columns: userColumns } = getUserSheet(workbook);
      const parents = [];
      for (let rowNumber = 2; rowNumber <= userSheet.rowCount; rowNumber += 1) {
        const row = userSheet.getRow(rowNumber);
        if (
          cellText(row.getCell(userColumns.role).value).toLowerCase() !== "parent"
          || cellText(row.getCell(userColumns.classId).value) !== session.classId
          || cellText(row.getCell(userColumns.status).value).toLowerCase() !== "active"
        ) continue;
        parents.push({
          id: cellText(row.getCell(userColumns.id).value),
          name: cellText(row.getCell(userColumns.name).value),
          email: cellText(row.getCell(userColumns.email).value),
        });
      }
      return {
        students: getOwnedStudents(studentSheet, studentColumns, session),
        parents,
      };
    });
  }

  async function verifyRosterClearPassword(session, data) {
    const password = typeof data.password === "string" ? data.password : "";
    if (password.length < 6 || password.length > 256) {
      throw new HttpError(400, "Enter your current password.");
    }
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet, columns } = getUserSheet(workbook);
      let userRow = null;
      for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const row = sheet.getRow(rowNumber);
        if (cellText(row.getCell(columns.id).value) === session.userId) {
          userRow = row;
          break;
        }
      }
      if (!userRow || cellText(userRow.getCell(columns.role).value).toLowerCase() !== "teacher") {
        throw new HttpError(404, "The teacher account could not be found.");
      }
      const matches = await verifyPassword(password, cellText(userRow.getCell(columns.password).value));
      if (!matches) throw new HttpError(400, "Your password doesn't match.");
      return { verified: true };
    });
  }

  async function clearStudentRoster(session, data) {
    const password = typeof data.password === "string" ? data.password : "";
    if (password.length < 6 || password.length > 256) {
      throw new HttpError(400, "Enter your current password.");
    }
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: userSheet, columns: userColumns } = getUserSheet(workbook);
      let userRow = null;
      for (let rowNumber = 2; rowNumber <= userSheet.rowCount; rowNumber += 1) {
        const row = userSheet.getRow(rowNumber);
        if (cellText(row.getCell(userColumns.id).value) === session.userId) {
          userRow = row;
          break;
        }
      }
      if (!userRow || cellText(userRow.getCell(userColumns.role).value).toLowerCase() !== "teacher") {
        throw new HttpError(404, "The teacher account could not be found.");
      }
      const matches = await verifyPassword(password, cellText(userRow.getCell(userColumns.password).value));
      if (!matches) throw new HttpError(400, "Your password doesn't match.");

      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      const studentIds = new Set();
      const studentRowsToDelete = [];
      for (let rowNumber = 2; rowNumber <= studentSheet.rowCount; rowNumber += 1) {
        const row = studentSheet.getRow(rowNumber);
        const ownerId = cellText(row.getCell(studentColumns.teacherUserId).value);
        if (
          cellText(row.getCell(studentColumns.classId).value) === session.classId
          && (!ownerId || ownerId === session.userId)
        ) {
          const studentId = cellText(row.getCell(studentColumns.id).value);
          if (studentId) studentIds.add(studentId);
          studentRowsToDelete.push(rowNumber);
        }
      }
      if (!studentRowsToDelete.length) return { deleted: 0 };

      const relatedSheets = [
        ["Term Grades", getGradeSheet, "studentId"],
        ["Term Grade Locks", getGradeLockSheet, "studentId"],
        ["Reading Assessments", getReadingAssessmentSheet, "studentId"],
        ["Attendance Records", getAttendanceSheet, "studentId"],
        ["Attendance Monthly Notes", getAttendanceNotesSheet, "studentId"],
      ];
      for (const [sheetName, getSheet, studentIdColumn] of relatedSheets) {
        if (!studentIds.size) break;
        if (!workbook.getWorksheet(sheetName)) continue;
        const { sheet: relatedSheet, columns } = getSheet(workbook);
        const rowsToDelete = [];
        for (let rowNumber = 2; rowNumber <= relatedSheet.rowCount; rowNumber += 1) {
          const row = relatedSheet.getRow(rowNumber);
          if (
            cellText(row.getCell(columns.classId).value) === session.classId
            && studentIds.has(cellText(row.getCell(columns[studentIdColumn]).value))
          ) {
            rowsToDelete.push(rowNumber);
          }
        }
        rowsToDelete.sort((left, right) => right - left)
          .forEach((rowNumber) => relatedSheet.spliceRows(rowNumber, 1));
      }
      studentRowsToDelete.sort((left, right) => right - left)
        .forEach((rowNumber) => studentSheet.spliceRows(rowNumber, 1));

      const { sheet: classSheet, columns: classColumns } = getClassSheet(workbook);
      for (let rowNumber = 2; rowNumber <= classSheet.rowCount; rowNumber += 1) {
        const row = classSheet.getRow(rowNumber);
        if (
          cellText(row.getCell(classColumns.id).value) === session.classId
          && cellText(row.getCell(classColumns.teacherId).value) === session.userId
        ) {
          row.getCell(classColumns.students).value = Math.max(
            0,
            (Number(row.getCell(classColumns.students).value) || 0) - studentRowsToDelete.length,
          );
          break;
        }
      }
      await saveDatabaseWorkbook(workbook);
      return { deleted: studentRowsToDelete.length };
    });
  }

  async function getParentPortal(session) {
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: userSheet, columns: userColumns } = getUserSheet(workbook);
      let parent = null;
      for (let rowNumber = 2; rowNumber <= userSheet.rowCount; rowNumber += 1) {
        const row = userSheet.getRow(rowNumber);
        if (
          cellText(row.getCell(userColumns.id).value) === session.userId
          && cellText(row.getCell(userColumns.role).value).toLowerCase() === "parent"
          && cellText(row.getCell(userColumns.classId).value) === session.classId
          && cellText(row.getCell(userColumns.status).value).toLowerCase() === "active"
        ) {
          parent = row;
          break;
        }
      }
      if (!parent) throw new HttpError(404, "The parent account could not be found.");

      const { sheet: classSheet, columns: classColumns } = getClassSheet(workbook);
      let classroom = null;
      for (let rowNumber = 2; rowNumber <= classSheet.rowCount; rowNumber += 1) {
        const row = classSheet.getRow(rowNumber);
        if (cellText(row.getCell(classColumns.id).value) === session.classId) {
          classroom = {
            number: cellText(row.getCell(classColumns.classroom).value),
            teacherName: cellText(row.getCell(classColumns.teacherName).value),
            teacherId: cellText(row.getCell(classColumns.teacherId).value),
          };
          break;
        }
      }
      if (!classroom) throw new HttpError(404, "The class linked to this parent account could not be found.");

      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      const learners = [];
      const learnersById = new Map();
      for (let rowNumber = 2; rowNumber <= studentSheet.rowCount; rowNumber += 1) {
        const row = studentSheet.getRow(rowNumber);
        if (
          cellText(row.getCell(studentColumns.classId).value) !== session.classId
          || cellText(row.getCell(studentColumns.teacherUserId).value) !== classroom.teacherId
          || cellText(row.getCell(studentColumns.parentUserId).value) !== session.userId
        ) continue;
        const learner = {
          id: cellText(row.getCell(studentColumns.id).value),
          name: cellText(row.getCell(studentColumns.name).value),
          profile: STUDENT_FIELDS.filter(([key]) => key !== "teacherUserId").map(([key, label]) => {
            const value = row.getCell(studentColumns[key]).value;
            return {
              label,
              value: key === "birthdate" && value instanceof Date
                ? value.toISOString().slice(0, 10)
                : cellText(value),
            };
          }),
          grades: Object.fromEntries(GRADE_TERMS.map((term) => [
            term,
            Object.fromEntries(GRADE_SUBJECTS.map(([key]) => [key, ""])),
          ])),
          reading: Object.fromEntries(READING_PERIODS.map((period) => [
            period,
            Object.fromEntries(READING_LANGUAGES.map((language) => [
              language,
              { level: "", assessedDate: "", attempts: [] },
            ])),
          ])),
          attendance: [],
        };
        learners.push(learner);
        learnersById.set(learner.id, learner);
      }

      const gradesSheet = workbook.getWorksheet("Term Grades");
      if (gradesSheet) {
        const { columns } = getGradeSheet(workbook);
        for (let rowNumber = 2; rowNumber <= gradesSheet.rowCount; rowNumber += 1) {
          const row = gradesSheet.getRow(rowNumber);
          if (
            cellText(row.getCell(columns.classId).value) !== session.classId
            || cellText(row.getCell(columns.teacherUserId).value) !== classroom.teacherId
          ) continue;
          const learner = learnersById.get(cellText(row.getCell(columns.studentId).value));
          const term = cellText(row.getCell(columns.term).value);
          const subject = cellText(row.getCell(columns.subject).value);
          if (!learner || !Object.hasOwn(learner.grades, term)
            || !Object.hasOwn(learner.grades[term], subject)) continue;
          const value = row.getCell(columns.grade).value;
          learner.grades[term][subject] = value === null || value === undefined || value === ""
            ? ""
            : Number(value);
        }
      }
      for (const learner of learners) {
        for (const term of GRADE_TERMS) {
          const music = learner.grades[term].musicAndArts;
          const pe = learner.grades[term].physicalEducationAndHealth;
          learner.grades[term].mapeh = typeof music === "number" && typeof pe === "number"
            ? Math.round(((music + pe) / 2) * 100) / 100
            : "";
        }
      }

      const readingSheet = workbook.getWorksheet("Reading Assessments");
      if (readingSheet) {
        const { columns } = getReadingAssessmentSheet(workbook);
        for (let rowNumber = 2; rowNumber <= readingSheet.rowCount; rowNumber += 1) {
          const row = readingSheet.getRow(rowNumber);
          if (
            cellText(row.getCell(columns.classId).value) !== session.classId
            || cellText(row.getCell(columns.teacherUserId).value) !== classroom.teacherId
          ) continue;
          const learner = learnersById.get(cellText(row.getCell(columns.studentId).value));
          const period = cellText(row.getCell(columns.period).value);
          const language = cellText(row.getCell(columns.language).value);
          if (!learner || !Object.hasOwn(learner.reading, period)
            || !Object.hasOwn(learner.reading[period], language)) continue;
          const assessed = row.getCell(columns.assessed).value;
          const rawAttemptNo = Number(row.getCell(columns.attempt).value);
          const attemptNo = Number.isSafeInteger(rawAttemptNo) && rawAttemptNo > 0 ? rawAttemptNo : 1;
          const attempt = {
            attemptNo,
            level: cellText(row.getCell(columns.level).value),
            assessedDate: assessed instanceof Date ? assessed.toISOString().slice(0, 10) : cellText(assessed),
          };
          const languageReading = learner.reading[period][language];
          languageReading.attempts.push(attempt);
          if (!languageReading.level || attemptNo >= languageReading.attemptNo) {
            languageReading.level = attempt.level;
            languageReading.assessedDate = attempt.assessedDate;
            languageReading.attemptNo = attemptNo;
          }
        }
      }
      for (const learner of learners) {
        for (const period of READING_PERIODS) {
          for (const language of READING_LANGUAGES) {
            learner.reading[period][language].attempts.sort((left, right) => left.attemptNo - right.attemptNo);
          }
        }
      }

      const attendanceSheet = workbook.getWorksheet("Attendance Records");
      if (attendanceSheet) {
        const { columns } = getAttendanceSheet(workbook);
        for (let rowNumber = 2; rowNumber <= attendanceSheet.rowCount; rowNumber += 1) {
          const row = attendanceSheet.getRow(rowNumber);
          if (
            cellText(row.getCell(columns.classId).value) !== session.classId
            || cellText(row.getCell(columns.teacherUserId).value) !== classroom.teacherId
          ) continue;
          const learner = learnersById.get(cellText(row.getCell(columns.studentId).value));
          if (!learner) continue;
          const recordedDate = row.getCell(columns.date).value;
          const attendanceRecord = {
            date: recordedDate instanceof Date
              ? recordedDate.toISOString().slice(0, 10)
              : cellText(recordedDate),
            status: normalizeAttendanceStatus(row.getCell(columns.status).value),
          };
          const reason = cellText(row.getCell(columns.reason).value);
          if (reason) attendanceRecord.reason = reason;
          learner.attendance.push(attendanceRecord);
        }
      }
      for (const learner of learners) {
        learner.attendance.sort((left, right) => right.date.localeCompare(left.date));
      }

      return {
        parentName: cellText(parent.getCell(userColumns.name).value),
        email: cellText(parent.getCell(userColumns.email).value),
        classroomNumber: classroom.number,
        teacherName: classroom.teacherName,
        learners,
        gradeTerms: GRADE_TERMS,
        gradeSubjects: [
          ...GRADE_SUBJECTS.slice(0, 7),
          ["mapeh", "MAPEH (average)"],
          ...GRADE_SUBJECTS.slice(7),
        ].map(([key, label]) => ({ key, label })),
        readingPeriods: READING_PERIODS,
        readingLanguages: READING_LANGUAGES,
      };
    });
  }

  async function unlinkStudentParent(session, studentId) {
    if (!studentId || studentId.length > 40) {
      throw new HttpError(400, "Choose a student from your class roster.");
    }
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      const student = findOwnedStudent(studentSheet, studentColumns, session, studentId);
      student.getCell(studentColumns.teacherUserId).value = session.userId;
      student.getCell(studentColumns.parentUserId).value = null;
      await saveDatabaseWorkbook(workbook);
      return { unlinked: true };
    });
  }

  async function linkStudentParent(session, studentId, data) {
    if (!studentId || studentId.length > 40) {
      throw new HttpError(400, "Choose a student from your class roster.");
    }
    const parentUserId = typeof data.parentUserId === "string" ? data.parentUserId.trim() : "";
    if (parentUserId.length > 40) throw new HttpError(400, "Choose a valid parent account.");
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      const student = findOwnedStudent(studentSheet, studentColumns, session, studentId);
      if (parentUserId) {
        const { sheet: userSheet, columns: userColumns } = getUserSheet(workbook);
        let parentFound = false;
        for (let rowNumber = 2; rowNumber <= userSheet.rowCount; rowNumber += 1) {
          const row = userSheet.getRow(rowNumber);
          if (
            cellText(row.getCell(userColumns.id).value) === parentUserId
            && cellText(row.getCell(userColumns.role).value).toLowerCase() === "parent"
            && cellText(row.getCell(userColumns.classId).value) === session.classId
            && cellText(row.getCell(userColumns.status).value).toLowerCase() === "active"
          ) {
            parentFound = true;
            break;
          }
        }
        if (!parentFound) throw new HttpError(400, "Choose an active parent account registered with your class.");
      }
      student.getCell(studentColumns.teacherUserId).value = session.userId;
      student.getCell(studentColumns.parentUserId).value = parentUserId || null;
      await saveDatabaseWorkbook(workbook);
      return { saved: true, parentUserId };
    });
  }

  function writeStudents(workbook, session, students) {
    const { sheet, columns } = getStudentSheet(workbook);
    const added = [];
    for (const student of students) {
      const values = [];
      values[columns.id] = nextStudentId(sheet, columns.id);
      values[columns.name] = [student.lastName, student.firstName, student.middleName]
        .filter(Boolean)
        .join(", ");
      values[columns.classId] = session.classId;
      values[columns.parentUserId] = null;
      for (const [key] of STUDENT_FIELDS) {
        if (key === "teacherUserId") {
          values[columns[key]] = session.userId;
        } else if (key === "birthdate") {
          values[columns[key]] = student.birthdate;
        } else {
          values[columns[key]] = student[key] || null;
        }
      }
      const row = writeRow(sheet, [columns.id, columns.lrn], values);

      added.push({
        id: cellText(row.getCell(columns.id).value),
        name: cellText(row.getCell(columns.name).value),
        ...student,
        birthdate: student.birthdate.toISOString().slice(0, 10),
        teacherUserId: session.userId,
      });
    }

    if (added.length) {
      const { sheet: classSheet, columns: classColumns } = getClassSheet(workbook);
      for (let rowNumber = 2; rowNumber <= classSheet.rowCount; rowNumber += 1) {
        const classRow = classSheet.getRow(rowNumber);
        if (cellText(classRow.getCell(classColumns.id).value) === session.classId) {
          for (let index = 0; index < added.length; index += 1) {
            incrementCell(classRow.getCell(classColumns.students));
          }
          break;
        }
      }
    }
    return added;
  }

  async function addStudent(session, data) {
    const student = validateStudent(data);
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet, columns } = getStudentSheet(workbook);
      for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
        if (cellText(sheet.getRow(rowNumber).getCell(columns.lrn).value) === student.lrn) {
          throw new HttpError(409, "A student with this LRN is already registered.");
        }
      }
      const [added] = writeStudents(workbook, session, [student]);
      await saveDatabaseWorkbook(workbook);
      return added;
    });
  }

  async function importStudents(session, buffer) {
    const imported = parseSF1Workbook(buffer);
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet, columns } = getStudentSheet(workbook);
      const existingLrns = new Set();
      for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const existingLrn = cellText(sheet.getRow(rowNumber).getCell(columns.lrn).value);
        if (existingLrn) existingLrns.add(existingLrn);
      }

      const skipped = [...imported.skipped];
      const studentsToAdd = [];
      for (const entry of imported.students) {
        if (existingLrns.has(entry.student.lrn)) {
          skipped.push({ row: entry.row, reason: "LRN is already in the student roster or this SF1 file." });
          continue;
        }
        existingLrns.add(entry.student.lrn);
        studentsToAdd.push(entry.student);
      }

      const students = writeStudents(workbook, session, studentsToAdd);
      if (students.length) await saveDatabaseWorkbook(workbook);
      skipped.sort((left, right) => left.row - right.row);
      return { imported: students.length, skipped };
    });
  }

  async function getStudentGrades(session, studentId, termValue) {
    const term = validateGradeTerm(termValue);
    if (!studentId || studentId.length > 40) {
      throw new HttpError(400, "Choose a student from your class roster.");
    }
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      const student = findOwnedStudent(studentSheet, studentColumns, session, studentId);
      const { sheet, columns } = getGradeSheet(workbook);
      const grades = Object.fromEntries(GRADE_SUBJECTS.map(([key]) => [key, ""]));
      for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const row = sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(columns.studentId).value) !== studentId
          || cellText(row.getCell(columns.classId).value) !== session.classId
          || cellText(row.getCell(columns.teacherUserId).value) !== session.userId
          || cellText(row.getCell(columns.term).value) !== term
        ) continue;
        const subject = cellText(row.getCell(columns.subject).value);
        if (Object.hasOwn(grades, subject)) {
          const value = row.getCell(columns.grade).value;
          grades[subject] = value === null || value === undefined ? "" : String(value);
        }
      }
      return {
        student: {
          id: cellText(student.getCell(studentColumns.id).value),
          name: cellText(student.getCell(studentColumns.name).value),
        },
        term,
        grades,
      };
    });
  }

  function retireLegacyTermGradeLocks(session, sheet, columns) {
    let changed = false;
    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      if (
        cellText(row.getCell(columns.classId).value) !== session.classId
        || cellText(row.getCell(columns.teacherUserId).value) !== session.userId
        || cellText(row.getCell(columns.studentId).value)
      ) continue;

      const term = cellText(row.getCell(columns.term).value);
      const legacyUpdated = row.getCell(columns.updated).value;
      const legacyTimestamp = legacyUpdated instanceof Date
        ? legacyUpdated.getTime()
        : typeof legacyUpdated === "string" ? Date.parse(legacyUpdated) : Number.NaN;
      for (let studentRowNumber = 2; studentRowNumber <= sheet.rowCount; studentRowNumber += 1) {
        const studentRow = sheet.getRow(studentRowNumber);
        const studentId = cellText(studentRow.getCell(columns.studentId).value);
        const individualUpdated = studentRow.getCell(columns.updated).value;
        const individualTimestamp = individualUpdated instanceof Date
          ? individualUpdated.getTime()
          : typeof individualUpdated === "string" ? Date.parse(individualUpdated) : Number.NaN;
        const individualLocked = studentRow.getCell(columns.locked).value;
        if (
          studentId
          && cellText(studentRow.getCell(columns.classId).value) === session.classId
          && cellText(studentRow.getCell(columns.teacherUserId).value) === session.userId
          && cellText(studentRow.getCell(columns.term).value) === term
          && (individualLocked === true || cellText(individualLocked).toLowerCase() === "true")
          && Number.isFinite(legacyTimestamp)
          && Number.isFinite(individualTimestamp)
          && individualTimestamp <= legacyTimestamp
        ) {
          studentRow.getCell(columns.locked).value = false;
          studentRow.getCell(columns.updated).value = new Date();
          changed = true;
        }
      }

      row.getCell(columns.locked).value = false;
      row.getCell(columns.classId).value = "";
      row.getCell(columns.teacherUserId).value = "";
      row.getCell(columns.term).value = "";
      row.getCell(columns.updated).value = new Date();
      changed = true;
    }
    return changed;
  }

  async function getStudentGradeLock(session, studentIdValue, termValue) {
    const term = validateGradeTerm(termValue);
    const studentId = typeof studentIdValue === "string" ? studentIdValue.trim() : "";
    if (!studentId || studentId.length > 40) {
      throw new HttpError(400, "Choose a student from your class roster.");
    }
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      findOwnedStudent(studentSheet, studentColumns, session, studentId);
      const existingLockSheet = workbook.getWorksheet("Term Grade Locks");
      if (!existingLockSheet) return { studentId, term, locked: false };
      const { sheet, columns } = getGradeLockSheet(workbook);
      const migrated = retireLegacyTermGradeLocks(session, sheet, columns);
      let locked = false;
      for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const row = sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(columns.classId).value) === session.classId
          && cellText(row.getCell(columns.teacherUserId).value) === session.userId
          && cellText(row.getCell(columns.term).value) === term
          && cellText(row.getCell(columns.studentId).value) === studentId
        ) {
          const value = row.getCell(columns.locked).value;
          locked = value === true || cellText(value).toLowerCase() === "true";
          break;
        }
      }
      if (migrated) await saveDatabaseWorkbook(workbook);
      return { studentId, term, locked };
    });
  }

  async function setStudentGradeLock(session, data) {
    const term = validateGradeTerm(data && data.term);
    const studentId = typeof (data && data.studentId) === "string" ? data.studentId.trim() : "";
    if (!studentId || studentId.length > 40) {
      throw new HttpError(400, "Choose a student from your class roster.");
    }
    if (!data || typeof data.locked !== "boolean") {
      throw new HttpError(400, "Choose whether to lock or unlock this student's grades.");
    }
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      findOwnedStudent(studentSheet, studentColumns, session, studentId);
      const { sheet, columns } = getGradeLockSheet(workbook);
      retireLegacyTermGradeLocks(session, sheet, columns);
      let lockRow = null;
      for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const row = sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(columns.classId).value) === session.classId
          && cellText(row.getCell(columns.teacherUserId).value) === session.userId
          && cellText(row.getCell(columns.term).value) === term
          && cellText(row.getCell(columns.studentId).value) === studentId
        ) {
          lockRow = row;
          break;
        }
      }
      if (!lockRow) {
        lockRow = sheet.getRow(sheet.rowCount + 1);
        lockRow.getCell(columns.classId).value = session.classId;
        lockRow.getCell(columns.teacherUserId).value = session.userId;
        lockRow.getCell(columns.term).value = term;
        lockRow.getCell(columns.studentId).value = studentId;
      }
      lockRow.getCell(columns.locked).value = data.locked;
      lockRow.getCell(columns.updated).value = new Date();
      await saveDatabaseWorkbook(workbook);
      return { studentId, term, locked: data.locked };
    });
  }

  async function getGradeSummary(session) {
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      const students = [];
      const studentsById = new Map();

      for (let rowNumber = 2; rowNumber <= studentSheet.rowCount; rowNumber += 1) {
        const row = studentSheet.getRow(rowNumber);
        if (
          cellText(row.getCell(studentColumns.classId).value) !== session.classId
          || cellText(row.getCell(studentColumns.teacherUserId).value) !== session.userId
        ) continue;
        const id = cellText(row.getCell(studentColumns.id).value);
        const student = {
          id,
          name: cellText(row.getCell(studentColumns.name).value),
          lrn: cellText(row.getCell(studentColumns.lrn).value),
          grades: Object.fromEntries(GRADE_TERMS.map((term) => [
            term,
            Object.fromEntries(GRADE_SUBJECTS.map(([key]) => [key, ""])),
          ])),
        };
        students.push(student);
        studentsById.set(id, student);
      }

      const gradeSheet = workbook.getWorksheet("Term Grades");
      if (gradeSheet) {
        const columns = getGradeSheet(workbook).columns;
        for (let rowNumber = 2; rowNumber <= gradeSheet.rowCount; rowNumber += 1) {
          const row = gradeSheet.getRow(rowNumber);
          if (
            cellText(row.getCell(columns.classId).value) !== session.classId
            || cellText(row.getCell(columns.teacherUserId).value) !== session.userId
          ) continue;
          const student = studentsById.get(cellText(row.getCell(columns.studentId).value));
          const term = cellText(row.getCell(columns.term).value);
          const subject = cellText(row.getCell(columns.subject).value);
          if (!student || !Object.hasOwn(student.grades, term)
            || !Object.hasOwn(student.grades[term], subject)) continue;
          const value = row.getCell(columns.grade).value;
          student.grades[term][subject] = value === null || value === undefined || value === ""
            ? ""
            : Number(value);
        }
      }

      for (const student of students) {
        for (const term of GRADE_TERMS) {
          const music = student.grades[term].musicAndArts;
          const physicalEducation = student.grades[term].physicalEducationAndHealth;
          student.grades[term].mapeh = typeof music === "number" && typeof physicalEducation === "number"
            ? Math.round(((music + physicalEducation) / 2) * 100) / 100
            : "";
        }
      }

      return {
        terms: GRADE_TERMS,
        subjects: [
          ...GRADE_SUBJECTS.slice(0, 7),
          ["mapeh", "MAPEH (average)"],
          ...GRADE_SUBJECTS.slice(7),
        ].map(([key, label]) => ({ key, label })),
        students,
      };
    });
  }

  async function saveStudentGrades(session, data) {
    const term = validateGradeTerm(data.term);
    const studentId = typeof data.studentId === "string" ? data.studentId.trim() : "";
    if (!studentId || studentId.length > 40) {
      throw new HttpError(400, "Choose a student from your class roster.");
    }
    const grades = validateStudentGrades(data.grades);
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      findOwnedStudent(studentSheet, studentColumns, session, studentId);
      const { sheet: lockSheet, columns: lockColumns } = getGradeLockSheet(workbook);
      retireLegacyTermGradeLocks(session, lockSheet, lockColumns);
      let studentLockRow = null;
      for (let rowNumber = 2; rowNumber <= lockSheet.rowCount; rowNumber += 1) {
        const row = lockSheet.getRow(rowNumber);
        if (
          cellText(row.getCell(lockColumns.classId).value) === session.classId
          && cellText(row.getCell(lockColumns.teacherUserId).value) === session.userId
          && cellText(row.getCell(lockColumns.term).value) === term
          && cellText(row.getCell(lockColumns.studentId).value) === studentId
        ) {
          studentLockRow = row;
          const value = row.getCell(lockColumns.locked).value;
          if (value === true || cellText(value).toLowerCase() === "true") {
            throw new HttpError(409, `${term} grades for this student are locked. Choose Edit Grades before saving changes.`);
          }
          break;
        }
      }
      if (!studentLockRow) {
        studentLockRow = lockSheet.getRow(lockSheet.rowCount + 1);
        studentLockRow.getCell(lockColumns.classId).value = session.classId;
        studentLockRow.getCell(lockColumns.teacherUserId).value = session.userId;
        studentLockRow.getCell(lockColumns.term).value = term;
        studentLockRow.getCell(lockColumns.studentId).value = studentId;
      }
      const { sheet, columns } = getGradeSheet(workbook);
      const rowsBySubject = new Map();
      for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const row = sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(columns.studentId).value) === studentId
          && cellText(row.getCell(columns.classId).value) === session.classId
          && cellText(row.getCell(columns.teacherUserId).value) === session.userId
          && cellText(row.getCell(columns.term).value) === term
        ) {
          rowsBySubject.set(cellText(row.getCell(columns.subject).value), row);
        }
      }

      for (const [subject] of GRADE_SUBJECTS) {
        let row = rowsBySubject.get(subject);
        if (!row) {
          row = writeRow(sheet, [columns.id, columns.subject], {
            [columns.id]: nextGradeId(sheet, columns.id),
            [columns.studentId]: studentId,
            [columns.classId]: session.classId,
            [columns.teacherUserId]: session.userId,
            [columns.term]: term,
            [columns.subject]: subject,
            [columns.grade]: grades[subject],
            [columns.updated]: new Date(),
          });
          rowsBySubject.set(subject, row);
        } else {
          row.getCell(columns.grade).value = grades[subject];
          row.getCell(columns.updated).value = new Date();
        }
      }
      studentLockRow.getCell(lockColumns.locked).value = true;
      studentLockRow.getCell(lockColumns.updated).value = new Date();
      await saveDatabaseWorkbook(workbook);
      return { saved: true, studentId, term, grades, locked: true };
    });
  }

  async function getReadingAssessmentSummary(session, period, attemptNo) {
    if (!READING_PERIODS.includes(period)) {
      throw new HttpError(400, "Choose BOSY, MOSY, or EOSY for the assessment period.");
    }
    if (!Number.isSafeInteger(attemptNo) || attemptNo < 1) {
      throw new HttpError(400, "Choose a valid positive assessment attempt number.");
    }
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      const students = getOwnedStudents(studentSheet, studentColumns, session).map((student) => ({
        id: student.id,
        name: student.name,
        attemptNo,
        levels: Object.fromEntries(READING_LANGUAGES.map((language) => [language, ""])),
        assessedDates: Object.fromEntries(READING_LANGUAGES.map((language) => [language, ""])),
      }));
      const studentsById = new Map(students.map((student) => [student.id, student]));

      const readingSheet = workbook.getWorksheet("Reading Assessments");
      if (readingSheet) {
        const { columns } = getReadingAssessmentSheet(workbook);
        for (let rowNumber = 2; rowNumber <= readingSheet.rowCount; rowNumber += 1) {
          const row = readingSheet.getRow(rowNumber);
          if (
            cellText(row.getCell(columns.classId).value) !== session.classId
            || cellText(row.getCell(columns.teacherUserId).value) !== session.userId
            || cellText(row.getCell(columns.period).value) !== period
          ) continue;
          const rawAttemptNo = Number(row.getCell(columns.attempt).value);
          const rowAttemptNo = Number.isSafeInteger(rawAttemptNo) && rawAttemptNo > 0 ? rawAttemptNo : 1;
          if (rowAttemptNo !== attemptNo) continue;
          const student = studentsById.get(cellText(row.getCell(columns.studentId).value));
          const language = cellText(row.getCell(columns.language).value);
          if (!student || !READING_LANGUAGES.includes(language)) continue;
          student.levels[language] = cellText(row.getCell(columns.level).value);
          const assessedDate = row.getCell(columns.assessed).value;
          student.assessedDates[language] = assessedDate instanceof Date
            ? assessedDate.toISOString().slice(0, 10)
            : cellText(assessedDate);
        }
      }

      return { period, attemptNo, languages: READING_LANGUAGES, levels: READING_LEVELS, students };
    });
  }

  async function saveReadingAssessment(session, data) {
    const period = typeof data.period === "string" ? data.period : "";
    if (!READING_PERIODS.includes(period)) {
      throw new HttpError(400, "Choose BOSY, MOSY, or EOSY for the assessment period.");
    }
    const attemptNo = data.attemptNo === undefined ? 1 : data.attemptNo;
    if (!Number.isSafeInteger(attemptNo) || attemptNo < 1) {
      throw new HttpError(400, "Choose a valid positive assessment attempt number.");
    }
    const studentId = typeof data.studentId === "string" ? data.studentId.trim() : "";
    if (!studentId || studentId.length > 40) {
      throw new HttpError(400, "Choose a student from your class roster.");
    }
    if (!data.levels || typeof data.levels !== "object" || Array.isArray(data.levels)) {
      throw new HttpError(400, "Enter a reading level for English and Filipino.");
    }
    const levels = {};
    for (const language of READING_LANGUAGES) {
      const value = data.levels[language];
      if (typeof value !== "string" || (value !== "" && !READING_LEVELS.includes(value))) {
        throw new HttpError(400, `Choose a valid ${language} reading level.`);
      }
      levels[language] = value;
    }

    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      findOwnedStudent(studentSheet, studentColumns, session, studentId);
      const { sheet, columns } = getReadingAssessmentSheet(workbook);
      const rowsByLanguage = new Map();
      for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const row = sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(columns.studentId).value) === studentId
          && cellText(row.getCell(columns.classId).value) === session.classId
          && cellText(row.getCell(columns.teacherUserId).value) === session.userId
          && cellText(row.getCell(columns.period).value) === period
          && (Number(row.getCell(columns.attempt).value) || 1) === attemptNo
        ) {
          rowsByLanguage.set(cellText(row.getCell(columns.language).value), row);
        }
      }

      const rowsToDelete = [];
      const assessmentDate = new Date();
      const assessedDates = {};
      for (const language of READING_LANGUAGES) {
        const existing = rowsByLanguage.get(language);
        if (!levels[language]) {
          if (existing) rowsToDelete.push(existing.number);
          assessedDates[language] = "";
          continue;
        }
        if (existing) {
          existing.getCell(columns.level).value = levels[language];
          existing.getCell(columns.assessed).value = assessmentDate;
          existing.getCell(columns.updated).value = assessmentDate;
        } else {
          writeRow(sheet, [columns.id], {
            [columns.id]: nextReadingAssessmentId(sheet, columns.id),
            [columns.studentId]: studentId,
            [columns.classId]: session.classId,
            [columns.teacherUserId]: session.userId,
            [columns.period]: period,
            [columns.language]: language,
            [columns.level]: levels[language],
            [columns.assessed]: assessmentDate,
            [columns.attempt]: attemptNo,
            [columns.updated]: assessmentDate,
          });
        }
        if (existing) existing.getCell(columns.attempt).value = attemptNo;
        assessedDates[language] = assessmentDate.toISOString().slice(0, 10);
      }
      rowsToDelete.sort((left, right) => right - left).forEach((rowNumber) => sheet.spliceRows(rowNumber, 1));
      await saveDatabaseWorkbook(workbook);
      return { saved: true, attemptNo, assessedDates };
    });
  }

  function validateAttendanceDate(value) {
    if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(value)) {
      throw new HttpError(400, "Choose a valid attendance date.");
    }
    const date = new Date(`${value}T00:00:00.000Z`);
    if (Number.isNaN(date.getTime()) || date.toISOString().slice(0, 10) !== value) {
      throw new HttpError(400, "Choose a valid attendance date.");
    }
    return value;
  }

  function validateAttendanceMonth(value) {
    if (typeof value !== "string" || !/^(?:19|[2-9]\d)\d{2}-(0[1-9]|1[0-2])$/.test(value)) {
      throw new HttpError(400, "Choose a valid attendance month.");
    }
    return value;
  }

  function attendanceDates(month) {
    const [year, monthNumber] = month.split("-").map(Number);
    const dayCount = new Date(Date.UTC(year, monthNumber, 0)).getUTCDate();
    return Array.from({ length: dayCount }, (_, index) =>
      `${month}-${String(index + 1).padStart(2, "0")}`).filter((date) => {
      const weekday = new Date(`${date}T00:00:00.000Z`).getUTCDay();
      return weekday !== 0 && weekday !== 6;
    });
  }

  function attendanceSchoolDays(month) {
    return attendanceDates(month);
  }

  function attendanceDateText(value) {
    return value instanceof Date ? value.toISOString().slice(0, 10) : cellText(value);
  }

  function sf2WorksheetName(month) {
    const monthNames = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
    const [year, monthNumber] = month.split("-").map(Number);
    return `${monthNames[monthNumber - 1]} ${year}`;
  }

  function sf2DisplayedMonth(month) {
    const monthNames = [
      "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
      "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER",
    ];
    const monthNumber = Number(month.slice(5, 7));
    return monthNames[monthNumber - 1];
  }

  function normalizeSF2LearnerName(value) {
    return cellText(value).normalize("NFKD").replace(/[\u0300-\u036f]/g, "")
      .toLowerCase().replace(/[^a-z0-9]/g, "");
  }

  async function parseSF2Upload(buffer, month) {
    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.load(buffer);
    } catch {
      throw new HttpError(400, "The selected file is not a readable SF2 .xlsx workbook.");
    }
    const sheetName = sf2WorksheetName(month);
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      throw new HttpError(400, `The uploaded SF2 workbook must contain a "${sheetName}" worksheet for the selected month.`);
    }
    if (
      !normalizeHeader(cellText(sheet.getCell(9, 2).value)).startsWith("learnersname")
      || !normalizeHeader(cellText(sheet.getCell(9, SF2_FIRST_DAY_COLUMN).value)).startsWith("1strowfordate")
      || !normalizeHeader(cellText(sheet.getCell(9, 32).value)).includes("totalforthemonth")
      || !normalizeHeader(cellText(sheet.getCell(9, 33).value)).includes("totalforthemonth")
      || !normalizeHeader(cellText(sheet.getCell(9, SF2_REMARKS_COLUMN).value)).startsWith("remarks")
      || !normalizeHeader(cellText(sheet.getCell(63, 2).value)).includes("male")
      || !normalizeHeader(cellText(sheet.getCell(63, 2).value)).includes("total")
      || !normalizeHeader(cellText(sheet.getCell(114, 2).value)).includes("female")
      || !normalizeHeader(cellText(sheet.getCell(114, 2).value)).includes("total")
    ) {
      throw new HttpError(400, "The selected worksheet does not match the SF2 template layout.");
    }

    const schoolDays = attendanceSchoolDays(month);
    const datesByColumn = new Map();
    for (let column = SF2_FIRST_DAY_COLUMN; column <= SF2_LAST_DAY_COLUMN; column += 1) {
      const rawDay = sheet.getCell(10, column).value;
      const schoolDayIndex = column - SF2_FIRST_DAY_COLUMN;
      if (schoolDayIndex >= schoolDays.length) {
        if (rawDay !== null && rawDay !== undefined && cellText(rawDay).trim() !== "") {
          throw new HttpError(400, `The date heading in column ${column} does not match ${sheetName}.`);
        }
        continue;
      }
      const expectedDate = schoolDays[schoolDayIndex];
      const expectedDay = Number(expectedDate.slice(-2));
      const weekday = new Date(`${expectedDate}T00:00:00.000Z`).getUTCDay();
      if (
        Number(rawDay) !== expectedDay
        || cellText(sheet.getCell(11, column).value).trim().toUpperCase() !== SF2_WEEKDAY_NAMES[(weekday + 6) % 7]
      ) {
        throw new HttpError(400, `The date or weekday heading in column ${column} does not match ${sheetName}.`);
      }
      datesByColumn.set(column, expectedDate);
    }

    const records = [];
    const namesSeen = new Set();
    for (const rowNumber of SF2_STUDENT_ROWS) {
      const row = sheet.getRow(rowNumber);
      const name = cellText(row.getCell(2).value).trim();
      if (!name) {
        if (cellText(row.getCell(SF2_REMARKS_COLUMN).value).trim()) {
          throw new HttpError(400, `Remarks in row ${rowNumber} have no learner name.`);
        }
        for (let column = SF2_FIRST_DAY_COLUMN; column <= SF2_LAST_DAY_COLUMN; column += 1) {
          if (cellText(row.getCell(column).value).trim()) {
            throw new HttpError(400, `Attendance marks in row ${rowNumber} have no learner name.`);
          }
        }
        continue;
      }
      const normalizedName = normalizeSF2LearnerName(name);
      if (!normalizedName || namesSeen.has(normalizedName)) {
        throw new HttpError(400, `The learner name in row ${rowNumber} is blank or duplicated.`);
      }
      namesSeen.add(normalizedName);
      const attendance = [];
      for (let column = SF2_FIRST_DAY_COLUMN; column <= SF2_LAST_DAY_COLUMN; column += 1) {
        const rawMark = cellText(row.getCell(column).value).trim().toLowerCase();
        if (!datesByColumn.has(column)) {
          if (rawMark) {
            throw new HttpError(400, `Attendance marks in row ${rowNumber}, column ${column} have no date heading.`);
          }
          continue;
        }
        const status = SF2_STATUS_BY_MARK.get(rawMark);
        if (status === undefined) {
          throw new HttpError(
            400,
            `Unsupported SF2 mark "${rawMark}" in row ${rowNumber}, column ${column}. Use blank, x, h, t, e, NC, or N/A for a locked learner.`,
          );
        }
        attendance.push({
          date: datesByColumn.get(column),
          status,
        });
      }
      const rawRemarks = cellText(row.getCell(SF2_REMARKS_COLUMN).value).trim();
      if (rawRemarks.length > 500) {
        throw new HttpError(400, `Remarks in row ${rowNumber} exceed the 500-character limit.`);
      }
      const statusDateMatches = [...rawRemarks.matchAll(/\n\[Grade Portal status date: (\d{4}-\d{2}-\d{2})\]/g)];
      if (statusDateMatches.length > 1) {
        throw new HttpError(400, `The learner status date in row ${rowNumber} is duplicated.`);
      }
      const statusDate = statusDateMatches[0]?.[1] || "";
      const remarksWithoutStatusDate = rawRemarks.replace(/\n\[Grade Portal status date: \d{4}-\d{2}-\d{2}\]/g, "");
      const markerIndex = remarksWithoutStatusDate.lastIndexOf(SF2_ABSENCE_REASONS_MARKER);
      const remarks = markerIndex < 0
        ? remarksWithoutStatusDate.trim()
        : remarksWithoutStatusDate.slice(0, markerIndex).trim();
      const absenceReasons = new Map();
      if (markerIndex >= 0) {
        const absentDates = new Set(
          attendance.filter(({ status }) => status === "Absent").map(({ date }) => date),
        );
        const reasonLines = remarksWithoutStatusDate.slice(markerIndex + SF2_ABSENCE_REASONS_MARKER.length)
          .split(/\r?\n/).filter(Boolean);
        if (!reasonLines.length) {
          throw new HttpError(400, `The absence reasons in row ${rowNumber} are empty.`);
        }
        for (const line of reasonLines) {
          const match = line.match(/^(\d{4}-\d{2}-\d{2}):\s(.+)$/);
          const reason = match && match[2].trim();
          if (!match || !reason || !absentDates.has(match[1]) || absenceReasons.has(match[1])) {
            throw new HttpError(400, `The date-specific absence reason in row ${rowNumber} is invalid.`);
          }
          absenceReasons.set(match[1], reason);
        }
      }
      records.push({
        name,
        normalizedName,
        gender: rowNumber <= SF2_LAST_MALE_STUDENT_ROW ? "Male" : "Female",
        attendance,
        remarks,
        statusDate,
        absenceReasons,
        rowNumber,
      });
    }
    if (!records.length) {
      throw new HttpError(400, "No learner names were found in the selected SF2 worksheet.");
    }
    return records;
  }

  function normalizeAttendanceStatus(value) {
    const status = cellText(value);
    return status === "Late" ? "Tardy" : status;
  }

  function normalizeMonthlyAttendanceRemark(value) {
    const remark = cellText(value).trim().toLowerCase().replace(/\s+/g, " ");
    if (remark === "nls" || remark === "nls - no longer in school") return "NLS";
    if (remark === "transferred out" || remark === "transfered out") return "Transferred Out";
    if (remark === "transferred in" || remark === "transfered in") return "Transferred In";
    return "";
  }

  function isMonthlyAttendanceRemark(value) {
    return Boolean(normalizeMonthlyAttendanceRemark(value));
  }

  function attendanceMonthEnd(month) {
    const [year, monthNumber] = month.split("-").map(Number);
    return `${month}-${String(new Date(Date.UTC(year, monthNumber, 0)).getUTCDate()).padStart(2, "0")}`;
  }

  function monthlyAttendanceStatusEvents(workbook, session, studentIds) {
    const eventsByStudent = new Map([...studentIds].map((studentId) => [studentId, []]));
    const notes = workbook.getWorksheet("Attendance Monthly Notes");
    if (!notes) return eventsByStudent;
    const { columns } = getAttendanceNotesSheet(workbook);
    for (let rowNumber = 2; rowNumber <= notes.rowCount; rowNumber += 1) {
      const row = notes.getRow(rowNumber);
      const studentId = cellText(row.getCell(columns.studentId).value);
      if (
        !eventsByStudent.has(studentId)
        || cellText(row.getCell(columns.classId).value) !== session.classId
        || cellText(row.getCell(columns.teacherUserId).value) !== session.userId
      ) continue;
      const status = normalizeMonthlyAttendanceRemark(row.getCell(columns.remarks).value);
      const eventMonth = cellText(row.getCell(columns.month).value);
      if (!status || !/^\d{4}-(0[1-9]|1[0-2])$/.test(eventMonth)) continue;
      const storedDate = attendanceDateText(row.getCell(columns.statusDate).value);
      const statusDate = storedDate || attendanceMonthEnd(eventMonth);
      const parsedDate = new Date(`${statusDate}T00:00:00.000Z`);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(statusDate)
        || Number.isNaN(parsedDate.getTime())
        || parsedDate.toISOString().slice(0, 10) !== statusDate) continue;
      eventsByStudent.get(studentId).push({ status, date: statusDate });
    }
    for (const events of eventsByStudent.values()) {
      events.sort((left, right) => left.date.localeCompare(right.date));
    }
    return eventsByStudent;
  }

  function attendanceStatusEffectiveDate(event) {
    if (event.status === "Transferred In") return event.date;
    const [year, month] = event.date.slice(0, 7).split("-").map(Number);
    const nextMonth = new Date(Date.UTC(year, month, 1));
    return nextMonth.toISOString().slice(0, 10);
  }

  function isAttendanceDateLocked(events, date) {
    let active = true;
    if (events[0]?.status === "Transferred In" && events[0].date > date) active = false;
    for (const event of events) {
      if (attendanceStatusEffectiveDate(event) > date) break;
      active = event.status === "Transferred In";
    }
    return !active;
  }

  function latestMonthlyStatus(events, month) {
    const monthEnd = attendanceMonthEnd(month);
    return events.filter((event) => event.date <= monthEnd).slice(-1)[0] || null;
  }

  async function getAttendanceMonth(session, monthValue) {
    const month = validateAttendanceMonth(monthValue);
    const dates = attendanceDates(month);
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      const students = getOwnedStudents(studentSheet, studentColumns, session).map((student) => ({
        id: student.id,
        name: student.name,
        gender: student.gender,
        attendance: Array(dates.length).fill(""),
        reasons: Array(dates.length).fill(""),
        remarks: "",
        statusDate: "",
        statusInherited: false,
        attendanceLocked: false,
        lockedDates: [],
        statusBaselineLockedDates: [],
      }));
      const noClassDates = new Set();
      const calendar = getAttendanceCalendarSheet(workbook);
      for (let rowNumber = 2; rowNumber <= calendar.sheet.rowCount; rowNumber += 1) {
        const row = calendar.sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(calendar.columns.classId).value) === session.classId
          && cellText(row.getCell(calendar.columns.teacherUserId).value) === session.userId
          && cellText(row.getCell(calendar.columns.dayType).value) === "Holiday/No Class"
        ) {
          const date = attendanceDateText(row.getCell(calendar.columns.date).value);
          if (dates.includes(date)) noClassDates.add(date);
        }
      }
      const studentsById = new Map(students.map((student) => [student.id, student]));
      const sheet = workbook.getWorksheet("Attendance Records");
      if (sheet) {
        const { columns } = getAttendanceSheet(workbook);
        for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
          const row = sheet.getRow(rowNumber);
          if (
            cellText(row.getCell(columns.classId).value) !== session.classId
            || cellText(row.getCell(columns.teacherUserId).value) !== session.userId
          ) continue;
          const student = studentsById.get(cellText(row.getCell(columns.studentId).value));
          const date = attendanceDateText(row.getCell(columns.date).value);
          if (!student || !date.startsWith(`${month}-`)) continue;
          const dayIndex = dates.indexOf(date);
          if (dayIndex >= 0) {
            student.attendance[dayIndex] = normalizeAttendanceStatus(row.getCell(columns.status).value);
            student.reasons[dayIndex] = cellText(row.getCell(columns.reason).value);
          }
        }
      }
      const notesSheet = workbook.getWorksheet("Attendance Monthly Notes");
      if (notesSheet) {
        const { columns } = getAttendanceNotesSheet(workbook);
        const statusEvents = monthlyAttendanceStatusEvents(workbook, session, studentsById.keys());
        for (let rowNumber = 2; rowNumber <= notesSheet.rowCount; rowNumber += 1) {
          const row = notesSheet.getRow(rowNumber);
          if (
            studentsById.has(cellText(row.getCell(columns.studentId).value))
            && cellText(row.getCell(columns.classId).value) === session.classId
            && cellText(row.getCell(columns.teacherUserId).value) === session.userId
            && cellText(row.getCell(columns.month).value) === month
          ) {
            const student = studentsById.get(cellText(row.getCell(columns.studentId).value));
            const remark = cellText(row.getCell(columns.remarks).value);
            if (!normalizeMonthlyAttendanceRemark(remark)) student.remarks = remark;
          }
        }
        for (const student of students) {
          const event = latestMonthlyStatus(statusEvents.get(student.id) || [], month);
          if (event && (event.status !== "Transferred In" || event.date.startsWith(`${month}-`))) {
            student.remarks = event.status;
            student.statusDate = event.date;
            student.statusInherited = !event.date.startsWith(`${month}-`);
          }
          const events = statusEvents.get(student.id) || [];
          student.lockedDates = dates.filter((date) => isAttendanceDateLocked(events, date));
          student.attendanceLocked = student.lockedDates.length === dates.length;
          const baselineEvents = events.filter((statusEvent) => !statusEvent.date.startsWith(`${month}-`));
          student.statusBaselineLockedDates = dates.filter((date) => isAttendanceDateLocked(baselineEvents, date));
          dates.forEach((date, index) => {
            if (student.lockedDates.includes(date)) {
              student.attendance[index] = "";
              student.reasons[index] = "";
            }
          });
        }
      }
      return { month, dates, statuses: ATTENDANCE_STATUSES, noClassDates: [...noClassDates], students };
    });
  }

  async function getAttendanceSummary(session) {
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      const students = getOwnedStudents(studentSheet, studentColumns, session).map((student) => ({
        id: student.id,
        name: student.name,
        gender: student.gender,
        present: 0,
        absent: 0,
        months: {},
      }));
      const studentsById = new Map(students.map((student) => [student.id, student]));
      const statusEvents = monthlyAttendanceStatusEvents(workbook, session, studentsById.keys());
      const noClassDates = new Set();
      const calendar = getAttendanceCalendarSheet(workbook);
      for (let rowNumber = 2; rowNumber <= calendar.sheet.rowCount; rowNumber += 1) {
        const row = calendar.sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(calendar.columns.classId).value) === session.classId
          && cellText(row.getCell(calendar.columns.teacherUserId).value) === session.userId
          && cellText(row.getCell(calendar.columns.dayType).value) === "Holiday/No Class"
        ) {
          noClassDates.add(attendanceDateText(row.getCell(calendar.columns.date).value));
        }
      }
      const recordsByLearnerDate = new Map();
      const sheet = workbook.getWorksheet("Attendance Records");
      if (sheet) {
        const { columns } = getAttendanceSheet(workbook);
        for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
          const row = sheet.getRow(rowNumber);
          if (
            cellText(row.getCell(columns.classId).value) !== session.classId
            || cellText(row.getCell(columns.teacherUserId).value) !== session.userId
          ) continue;
          const studentId = cellText(row.getCell(columns.studentId).value);
          const date = attendanceDateText(row.getCell(columns.date).value);
          if (!studentsById.has(studentId) || !date) continue;
          recordsByLearnerDate.set(`${studentId}\u0000${date}`, {
            studentId,
            date,
            status: normalizeAttendanceStatus(row.getCell(columns.status).value),
          });
        }
      }
      let firstDate = "";
      let lastDate = "";
      const months = new Set();
      for (const record of recordsByLearnerDate.values()) {
        const month = record.date.slice(0, 7);
        months.add(month);
        const student = studentsById.get(record.studentId);
        const monthSummary = student.months[month] || { present: 0, absent: 0 };
        if (noClassDates.has(record.date)
          || isAttendanceDateLocked(statusEvents.get(record.studentId) || [], record.date)) {
          student.months[month] = monthSummary;
          continue;
        }
        if (record.status === "Present" || record.status === "Tardy") {
          student.present += 1;
          monthSummary.present += 1;
        } else if (record.status === "Half-day") {
          student.present += 0.5;
          monthSummary.present += 0.5;
        } else if (record.status === "Absent") {
          student.absent += 1;
          monthSummary.absent += 1;
        }
        student.months[month] = monthSummary;
        if (record.status === "Present" || record.status === "Tardy"
          || record.status === "Half-day" || record.status === "Absent") {
          if (!firstDate || record.date < firstDate) firstDate = record.date;
          if (!lastDate || record.date > lastDate) lastDate = record.date;
        }
      }
      const recordedMonths = [...months].sort();
      for (const student of students) {
        for (const month of recordedMonths) {
          if (!student.months[month]) student.months[month] = { present: 0, absent: 0 };
        }
      }
      return { students, months: recordedMonths, firstDate, lastDate };
    });
  }

  async function saveAttendanceMonth(session, data) {
    const month = validateAttendanceMonth(data.month);
    const dates = attendanceDates(month);
    if (data.noClassDates !== undefined && !Array.isArray(data.noClassDates)) {
      throw new HttpError(400, "Choose valid dates to mark as Holiday/No Class.");
    }
    const requestedNoClassDates = data.noClassDates === undefined
      ? null
      : new Set(data.noClassDates);
    if (
      requestedNoClassDates
      && (requestedNoClassDates.size !== data.noClassDates.length
        || [...requestedNoClassDates].some((date) => typeof date !== "string" || !dates.includes(date)))
    ) {
      throw new HttpError(400, "Choose valid dates to mark as Holiday/No Class.");
    }
    if (!Array.isArray(data.records)) {
      throw new HttpError(400, "Submit one attendance row for every student in your class.");
    }
    const recordsByStudentId = new Map();
    for (const record of data.records) {
      const studentId = record && typeof record.studentId === "string" ? record.studentId.trim() : "";
      if (!studentId || studentId.length > 40 || recordsByStudentId.has(studentId)
        || !Array.isArray(record.attendance) || record.attendance.length !== dates.length
        || (record.reasons !== undefined && (!Array.isArray(record.reasons)
          || record.reasons.length !== dates.length
          || record.reasons.some((reason) => typeof reason !== "string" || reason.length > 500)))
        || (record.remarks !== undefined && typeof record.remarks !== "string")
        || (record.statusDate !== undefined && typeof record.statusDate !== "string")
        || (record.remarks || "").length > 500) {
        throw new HttpError(400, "Each student must have exactly one complete attendance row for the selected month.");
      }
      const remarks = (record.remarks || "").trim();
      const statusDate = (record.statusDate || "").trim();
      const monthlyStatus = normalizeMonthlyAttendanceRemark(remarks);
      if (statusDate) {
        validateAttendanceDate(statusDate);
        if (statusDate > attendanceMonthEnd(month)) {
          throw new HttpError(400, "A monthly learner status date cannot be after the selected month.");
        }
      }
      if (monthlyStatus && !statusDate) {
        throw new HttpError(400, "Choose the date the learner status took effect.");
      }
      const statuses = record.attendance.map((status) => {
        const normalized = status === "Late" ? "Tardy" : status;
        if (typeof normalized !== "string" || (normalized !== "" && !ATTENDANCE_STATUSES.includes(normalized))) {
          throw new HttpError(400, "Choose Present, Absent, Tardy, Half-day, or Excused for each attendance day.");
        }
        return normalized;
      });
      const reasons = record.reasons || Array(dates.length).fill("");
      if (statuses.some((status, index) => status === "Absent" && !reasons[index].trim())) {
        throw new HttpError(400, "Add a reason for every learner marked Absent.");
      }
      recordsByStudentId.set(studentId, {
        attendance: statuses,
        reasons: reasons.map((reason) => reason.trim()),
        remarks,
        statusDate,
      });
    }

    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      const students = getOwnedStudents(studentSheet, studentColumns, session);
      if (recordsByStudentId.size !== students.length
        || students.some((student) => !recordsByStudentId.has(student.id))) {
        throw new HttpError(400, "Submit an attendance row for every student in your class.");
      }
      const notes = getAttendanceNotesSheet(workbook);
      const statusEvents = monthlyAttendanceStatusEvents(workbook, session, students.map((student) => student.id));
      for (const student of students) {
        const record = recordsByStudentId.get(student.id);
        const status = normalizeMonthlyAttendanceRemark(record.remarks);
        const previousStatus = latestMonthlyStatus(statusEvents.get(student.id) || [], month);
        if (status && record.statusDate.slice(0, 7) !== month
          && (record.statusDate !== previousStatus?.date || status !== previousStatus?.status)) {
          throw new HttpError(400, "Choose a status date in the selected month, or keep the existing status and its date.");
        }
        const events = (statusEvents.get(student.id) || []).filter((event) => !event.date.startsWith(`${month}-`));
        if (status && record.statusDate.startsWith(`${month}-`)) {
          events.push({ status, date: record.statusDate });
          events.sort((left, right) => left.date.localeCompare(right.date));
        }
        for (let index = 0; index < dates.length; index += 1) {
          if (isAttendanceDateLocked(events, dates[index])
            && (record.attendance[index] || record.reasons[index])) {
            throw new HttpError(400, `${student.name}'s attendance is locked from ${dates[index]} because the learner is marked NLS or Transferred Out.`);
          }
        }
      }
      const calendar = getAttendanceCalendarSheet(workbook);
      const existingNoClassDates = new Set();
      for (let rowNumber = 2; rowNumber <= calendar.sheet.rowCount; rowNumber += 1) {
        const row = calendar.sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(calendar.columns.classId).value) === session.classId
          && cellText(row.getCell(calendar.columns.teacherUserId).value) === session.userId
          && cellText(row.getCell(calendar.columns.dayType).value) === "Holiday/No Class"
        ) {
          const date = attendanceDateText(row.getCell(calendar.columns.date).value);
          if (dates.includes(date)) existingNoClassDates.add(date);
        }
      }
      const noClassDates = requestedNoClassDates || existingNoClassDates;
      for (const date of noClassDates) {
        const dayIndex = dates.indexOf(date);
        if ([...recordsByStudentId.values()].some((record) =>
          record.attendance[dayIndex] || record.reasons[dayIndex])) {
          throw new HttpError(400, `${date} is marked Holiday/No Class. Clear its attendance before saving.`);
        }
      }
      const calendarRowsByDate = new Map();
      for (let rowNumber = 2; rowNumber <= calendar.sheet.rowCount; rowNumber += 1) {
        const row = calendar.sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(calendar.columns.classId).value) === session.classId
          && cellText(row.getCell(calendar.columns.teacherUserId).value) === session.userId
        ) {
          const date = attendanceDateText(row.getCell(calendar.columns.date).value);
          if (dates.includes(date)) calendarRowsByDate.set(date, row);
        }
      }
      const calendarRowsToDelete = [];
      for (const [date, row] of calendarRowsByDate) {
        if (noClassDates.has(date)) {
          row.getCell(calendar.columns.dayType).value = "Holiday/No Class";
          row.getCell(calendar.columns.updated).value = new Date();
        } else {
          calendarRowsToDelete.push(row.number);
        }
      }
      for (const date of noClassDates) {
        if (calendarRowsByDate.has(date)) continue;
        writeRow(calendar.sheet, [calendar.columns.id], {
          [calendar.columns.id]: `AC${crypto.randomBytes(6).toString("hex").toUpperCase()}`,
          [calendar.columns.classId]: session.classId,
          [calendar.columns.teacherUserId]: session.userId,
          [calendar.columns.date]: date,
          [calendar.columns.dayType]: "Holiday/No Class",
          [calendar.columns.updated]: new Date(),
          [calendar.columns.key]: `${session.classId}${session.userId}${date}`,
        });
      }
      calendarRowsToDelete.sort((left, right) => right - left)
        .forEach((rowNumber) => calendar.sheet.spliceRows(rowNumber, 1));
      const { sheet, columns } = getAttendanceSheet(workbook);
      const rowsByKey = new Map();
      for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const row = sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(columns.classId).value) !== session.classId
          || cellText(row.getCell(columns.teacherUserId).value) !== session.userId
        ) continue;
        const date = attendanceDateText(row.getCell(columns.date).value);
        if (date.startsWith(`${month}-`)) {
          rowsByKey.set(`${cellText(row.getCell(columns.studentId).value)}|${date}`, row);
        }
      }
      const rowsToDelete = [];
      for (const [studentId, record] of recordsByStudentId) {
        record.attendance.forEach((status, index) => {
          const key = `${studentId}|${dates[index]}`;
          const existing = rowsByKey.get(key);
          if (!status || noClassDates.has(dates[index])) {
            if (existing) rowsToDelete.push(existing.number);
          } else if (existing) {
            existing.getCell(columns.status).value = status;
            existing.getCell(columns.reason).value = record.reasons[index];
            existing.getCell(columns.updated).value = new Date();
            existing.getCell(columns.key).value = `${studentId}${dates[index]}`;
          } else {
            writeRow(sheet, [columns.id], {
              [columns.id]: `A${crypto.randomBytes(6).toString("hex").toUpperCase()}`,
              [columns.studentId]: studentId,
              [columns.classId]: session.classId,
              [columns.teacherUserId]: session.userId,
              [columns.date]: dates[index],
              [columns.status]: status,
              [columns.reason]: record.reasons[index],
              [columns.updated]: new Date(),
              [columns.key]: `${studentId}${dates[index]}`,
            });
          }
        });
      }
      rowsToDelete.sort((left, right) => right - left)
        .forEach((rowNumber) => sheet.spliceRows(rowNumber, 1));
      const rowsByStudentId = new Map();
      for (let rowNumber = 2; rowNumber <= notes.sheet.rowCount; rowNumber += 1) {
        const row = notes.sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(notes.columns.classId).value) === session.classId
          && cellText(row.getCell(notes.columns.teacherUserId).value) === session.userId
          && cellText(row.getCell(notes.columns.month).value) === month
        ) {
          rowsByStudentId.set(cellText(row.getCell(notes.columns.studentId).value), row);
        }
      }
      const noteRowsToDelete = [];
      for (const [studentId, record] of recordsByStudentId) {
        const existing = rowsByStudentId.get(studentId);
        const status = normalizeMonthlyAttendanceRemark(record.remarks);
        const persistRemark = status && !record.statusDate.startsWith(`${month}-`)
          ? ""
          : record.remarks;
        if (!persistRemark) {
          if (existing) noteRowsToDelete.push(existing.number);
        } else if (existing) {
          existing.getCell(notes.columns.remarks).value = persistRemark;
          existing.getCell(notes.columns.statusDate).value = status ? record.statusDate : null;
          existing.getCell(notes.columns.updated).value = new Date();
          existing.getCell(notes.columns.key).value = `${studentId}${month}`;
        } else {
          writeRow(notes.sheet, [notes.columns.id], {
            [notes.columns.id]: `AN${crypto.randomBytes(6).toString("hex").toUpperCase()}`,
            [notes.columns.studentId]: studentId,
            [notes.columns.classId]: session.classId,
            [notes.columns.teacherUserId]: session.userId,
            [notes.columns.month]: month,
            [notes.columns.remarks]: persistRemark,
            [notes.columns.statusDate]: status ? record.statusDate : null,
            [notes.columns.updated]: new Date(),
            [notes.columns.key]: `${studentId}${month}`,
          });
        }
      }
      noteRowsToDelete.sort((left, right) => right - left)
        .forEach((rowNumber) => notes.sheet.spliceRows(rowNumber, 1));
      await saveDatabaseWorkbook(workbook);
      return { saved: true, month };
    });
  }

  async function importAttendanceSF2(session, monthValue, buffer) {
    const month = validateAttendanceMonth(monthValue);
    const importedRecords = await parseSF2Upload(buffer, month);
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      const students = getOwnedStudents(studentSheet, studentColumns, session);
      const maleCount = students.filter((student) => student.gender === "Male").length;
      const femaleCount = students.filter((student) => student.gender === "Female").length;
      if (
        maleCount + femaleCount !== students.length
        || maleCount > SF2_MAX_STUDENTS_PER_GENDER
        || femaleCount > SF2_MAX_STUDENTS_PER_GENDER
      ) {
        throw new HttpError(400, "This class roster exceeds the SF2 template's 50-learner limit per gender section.");
      }

      const studentsByName = new Map();
      for (const student of students) {
        const key = normalizeSF2LearnerName(student.name);
        if (!studentsByName.has(key)) studentsByName.set(key, []);
        studentsByName.get(key).push(student);
      }
      const matchedRecords = [];
      const unmatchedRows = [];
      const importedNoClassDates = new Set();
      for (const record of importedRecords) {
        const matches = studentsByName.get(record.normalizedName) || [];
        if (matches.length !== 1 || matches[0].gender !== record.gender) {
          unmatchedRows.push(record.rowNumber);
          continue;
        }
        for (const entry of record.attendance) {
          if (entry.status === "No Class") importedNoClassDates.add(entry.date);
        }
        if (record.attendance.some(({ date, status }) =>
          status === "Absent"
          && (!record.remarks || isMonthlyAttendanceRemark(record.remarks))
          && !record.absenceReasons.has(date))) {
          throw new HttpError(
            400,
            `Row ${record.rowNumber} marks a learner absent but has no SF2 remarks to use as the required absence reason.`,
          );
        }
        matchedRecords.push({
          ...record,
          remarks: normalizeMonthlyAttendanceRemark(record.remarks) || record.remarks,
          statusDate: record.statusDate
            || (isMonthlyAttendanceRemark(record.remarks) ? attendanceMonthEnd(month) : ""),
          student: matches[0],
        });
      }
      if (unmatchedRows.length) {
        throw new HttpError(
          400,
          `No unique learner matching the SF2 name and gender section was found in your class roster for row${unmatchedRows.join(", ")}. No attendance was changed.`,
        );
      }
      const statusEvents = monthlyAttendanceStatusEvents(workbook, session, students.map((student) => student.id));
      for (const record of matchedRecords) {
        const status = normalizeMonthlyAttendanceRemark(record.remarks);
        if (status && record.statusDate) {
          validateAttendanceDate(record.statusDate);
          if (record.statusDate > attendanceMonthEnd(month)) {
            throw new HttpError(400, `The learner status date in row ${record.rowNumber} is after the selected month.`);
          }
        }
        const previousStatus = latestMonthlyStatus(statusEvents.get(record.student.id) || [], month);
        if (status && record.statusDate.slice(0, 7) !== month
          && (record.statusDate !== previousStatus?.date || status !== previousStatus?.status)) {
          throw new HttpError(400, `The learner status date in row ${record.rowNumber} does not match the saved status history.`);
        }
        const events = (statusEvents.get(record.student.id) || [])
          .filter((event) => !event.date.startsWith(`${month}-`));
        if (status && record.statusDate.startsWith(`${month}-`)) {
          events.push({ status, date: record.statusDate });
          events.sort((left, right) => left.date.localeCompare(right.date));
        }
        for (const entry of record.attendance) {
          const locked = isAttendanceDateLocked(events, entry.date);
          if ((locked && entry.status !== "Locked" && entry.status !== "No Class")
            || (!locked && entry.status === "Locked")) {
            throw new HttpError(
              400,
              `${record.student.name}'s SF2 attendance does not match the NLS/Transferred Out lock for ${entry.date}.`,
            );
          }
        }
      }
      for (const date of importedNoClassDates) {
        if (matchedRecords.some((record) =>
          record.attendance.find((entry) => entry.date === date)?.status !== "No Class")) {
          throw new HttpError(400, `${date} must be marked NC for every learner in the SF2 worksheet.`);
        }
      }

      const importedStudentIds = new Set(matchedRecords.map(({ student }) => student.id));
      const calendar = getAttendanceCalendarSheet(workbook);
      const noClassDates = new Set();
      for (let rowNumber = 2; rowNumber <= calendar.sheet.rowCount; rowNumber += 1) {
        const row = calendar.sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(calendar.columns.classId).value) === session.classId
          && cellText(row.getCell(calendar.columns.teacherUserId).value) === session.userId
          && cellText(row.getCell(calendar.columns.dayType).value) === "Holiday/No Class"
        ) {
          const date = attendanceDateText(row.getCell(calendar.columns.date).value);
          if (date.startsWith(`${month}-`)) noClassDates.add(date);
        }
      }
      for (const date of importedNoClassDates) noClassDates.add(date);
      const calendarRowsByDate = new Map();
      for (let rowNumber = 2; rowNumber <= calendar.sheet.rowCount; rowNumber += 1) {
        const row = calendar.sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(calendar.columns.classId).value) === session.classId
          && cellText(row.getCell(calendar.columns.teacherUserId).value) === session.userId
        ) {
          const date = attendanceDateText(row.getCell(calendar.columns.date).value);
          if (date.startsWith(`${month}-`)) calendarRowsByDate.set(date, row);
        }
      }
      for (const date of importedNoClassDates) {
        const existing = calendarRowsByDate.get(date);
        if (existing) {
          existing.getCell(calendar.columns.dayType).value = "Holiday/No Class";
          existing.getCell(calendar.columns.updated).value = new Date();
          continue;
        }
        const row = writeRow(calendar.sheet, [calendar.columns.id], {
          [calendar.columns.id]: `AC${crypto.randomBytes(6).toString("hex").toUpperCase()}`,
          [calendar.columns.classId]: session.classId,
          [calendar.columns.teacherUserId]: session.userId,
          [calendar.columns.date]: date,
          [calendar.columns.dayType]: "Holiday/No Class",
          [calendar.columns.updated]: new Date(),
          [calendar.columns.key]: `${session.classId}${session.userId}${date}`,
        });
        calendarRowsByDate.set(date, row);
      }
      const { sheet: attendanceSheet, columns: attendanceColumns } = getAttendanceSheet(workbook);
      const attendanceRowsToDelete = [];
      for (let rowNumber = 2; rowNumber <= attendanceSheet.rowCount; rowNumber += 1) {
        const row = attendanceSheet.getRow(rowNumber);
        const date = attendanceDateText(row.getCell(attendanceColumns.date).value);
        if (
          importedStudentIds.has(cellText(row.getCell(attendanceColumns.studentId).value))
          && cellText(row.getCell(attendanceColumns.classId).value) === session.classId
          && cellText(row.getCell(attendanceColumns.teacherUserId).value) === session.userId
          && date.startsWith(`${month}-`)
        ) {
          attendanceRowsToDelete.push(rowNumber);
        }
      }
      attendanceRowsToDelete.sort((left, right) => right - left)
        .forEach((rowNumber) => attendanceSheet.spliceRows(rowNumber, 1));

      let importedDays = 0;
      for (const record of matchedRecords) {
        for (const entry of record.attendance) {
          if (noClassDates.has(entry.date) || entry.status === "Locked" || entry.status === "No Class") continue;
          const reason = entry.status === "Absent"
            ? record.absenceReasons.get(entry.date)
              || (isMonthlyAttendanceRemark(record.remarks) ? "" : record.remarks)
            : "";
          writeRow(attendanceSheet, [attendanceColumns.id], {
            [attendanceColumns.id]: `A${crypto.randomBytes(6).toString("hex").toUpperCase()}`,
            [attendanceColumns.studentId]: record.student.id,
            [attendanceColumns.classId]: session.classId,
            [attendanceColumns.teacherUserId]: session.userId,
            [attendanceColumns.date]: entry.date,
            [attendanceColumns.status]: entry.status,
            [attendanceColumns.reason]: reason,
            [attendanceColumns.updated]: new Date(),
            [attendanceColumns.key]: `${record.student.id}${entry.date}`,
          });
          importedDays += 1;
        }
      }

      const notes = getAttendanceNotesSheet(workbook);
      const noteRowsToDelete = [];
      for (let rowNumber = 2; rowNumber <= notes.sheet.rowCount; rowNumber += 1) {
        const row = notes.sheet.getRow(rowNumber);
        if (
          importedStudentIds.has(cellText(row.getCell(notes.columns.studentId).value))
          && cellText(row.getCell(notes.columns.classId).value) === session.classId
          && cellText(row.getCell(notes.columns.teacherUserId).value) === session.userId
          && cellText(row.getCell(notes.columns.month).value) === month
        ) {
          noteRowsToDelete.push(rowNumber);
        }
      }
      noteRowsToDelete.sort((left, right) => right - left)
        .forEach((rowNumber) => notes.sheet.spliceRows(rowNumber, 1));
      for (const record of matchedRecords) {
        if (!record.remarks) continue;
        if (normalizeMonthlyAttendanceRemark(record.remarks)
          && !record.statusDate.startsWith(`${month}-`)) continue;
        writeRow(notes.sheet, [notes.columns.id], {
          [notes.columns.id]: `AN${crypto.randomBytes(6).toString("hex").toUpperCase()}`,
          [notes.columns.studentId]: record.student.id,
          [notes.columns.classId]: session.classId,
          [notes.columns.teacherUserId]: session.userId,
          [notes.columns.month]: month,
          [notes.columns.remarks]: record.remarks,
          [notes.columns.statusDate]: record.statusDate || null,
          [notes.columns.updated]: new Date(),
          [notes.columns.key]: `${record.student.id}${month}`,
        });
      }
      await saveDatabaseWorkbook(workbook);
      return { imported: matchedRecords.length, importedDays, month };
    });
  }

  async function exportAttendanceSF2(session, monthValue) {
    const month = validateAttendanceMonth(monthValue);
    const attendance = await getAttendanceMonth(session, month);
    const teacherProfile = await getTeacherProfile(session);
    const maleStudents = attendance.students.filter((student) => student.gender === "Male");
    const femaleStudents = attendance.students.filter((student) => student.gender === "Female");
    if (
      maleStudents.length + femaleStudents.length !== attendance.students.length
      || maleStudents.length > SF2_MAX_STUDENTS_PER_GENDER
      || femaleStudents.length > SF2_MAX_STUDENTS_PER_GENDER
    ) {
      throw new HttpError(400, "This class roster exceeds the SF2 template's 50-learner limit per gender section.");
    }

    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.readFile(SF2_TEMPLATE_PATH);
    } catch (error) {
      throw new Error(`The SF2 template could not be read: ${error.message}`);
    }
    const sheet = workbook.getWorksheet("OCT 2026") || workbook.worksheets[0];
    if (
      !sheet
      || !normalizeHeader(cellText(sheet.getCell(9, 2).value)).startsWith("learnersname")
      || !normalizeHeader(cellText(sheet.getCell(9, SF2_FIRST_DAY_COLUMN).value)).startsWith("1strowfordate")
      || !normalizeHeader(cellText(sheet.getCell(9, 32).value)).includes("totalforthemonth")
      || !normalizeHeader(cellText(sheet.getCell(9, 33).value)).includes("totalforthemonth")
      || !normalizeHeader(cellText(sheet.getCell(9, SF2_REMARKS_COLUMN).value)).startsWith("remarks")
      || !normalizeHeader(cellText(sheet.getCell(63, 2).value)).includes("male")
      || !normalizeHeader(cellText(sheet.getCell(63, 2).value)).includes("total")
      || !normalizeHeader(cellText(sheet.getCell(114, 2).value)).includes("female")
      || !normalizeHeader(cellText(sheet.getCell(114, 2).value)).includes("total")
    ) {
      throw new Error("The project SF2 template does not match the expected worksheet layout.");
    }
    for (const otherSheet of workbook.worksheets.slice()) {
      if (otherSheet.id !== sheet.id) workbook.removeWorksheet(otherSheet.id);
    }
    sheet.name = sf2WorksheetName(month);
    const displayedMonth = sf2DisplayedMonth(month);
    sheet.getCell("N6").value = teacherProfile.schoolYear || null;
    sheet.getCell("G6").value = teacherProfile.schoolId || null;
    sheet.getCell("G7").value = teacherProfile.schoolName || null;
    sheet.getCell("AA7").value = teacherProfile.grade || null;
    sheet.getCell("AF7").value = teacherProfile.section || null;
    sheet.getCell("AA6").value = displayedMonth;
    sheet.getCell("AG117").value = displayedMonth;
    sheet.getCell("AG118").value = displayedMonth;
    sheet.getCell("AE140").value = teacherProfile.name;
    sheet.getCell("AE145").value = teacherProfile.schoolHead || null;
    const weekdayDates = attendanceSchoolDays(month);
    const classDayCount = weekdayDates.filter((date) =>
      !attendance.noClassDates.includes(date)).length;
    sheet.getCell("AQ9").value = classDayCount;
    sheet.getCell("AJ117").value = { formula: "AQ9", result: classDayCount };
    sheet.getCell("AJ118").value = { formula: "AQ9", result: classDayCount };
    for (let column = SF2_FIRST_DAY_COLUMN; column <= SF2_LAST_DAY_COLUMN; column += 1) {
      sheet.getCell(10, column).value = null;
      sheet.getCell(11, column).value = null;
    }
    const dateColumns = new Map();
    weekdayDates.forEach((date, index) => {
      const column = SF2_FIRST_DAY_COLUMN + index;
      const day = Number(date.slice(-2));
      const weekday = new Date(`${date}T00:00:00.000Z`).getUTCDay();
      sheet.getCell(10, column).value = day;
      sheet.getCell(11, column).value = SF2_WEEKDAY_NAMES[(weekday + 6) % 7];
      dateColumns.set(date, column);
    });
    for (const rowNumber of SF2_STUDENT_ROWS) {
      const row = sheet.getRow(rowNumber);
      row.hidden = true;
      row.getCell(1).value = null;
      for (let column = 2; column <= 6; column += 1) row.getCell(column).value = null;
      for (let column = SF2_FIRST_DAY_COLUMN; column <= SF2_LAST_DAY_COLUMN; column += 1) {
        row.getCell(column).value = null;
      }
      row.getCell(SF2_REMARKS_COLUMN).value = null;
      row.getCell(SF2_REMARKS_COLUMN + 1).value = null;
    }
    const outputRows = [
      ...maleStudents.map((student, index) => ({
        student,
        rowNumber: SF2_FIRST_MALE_STUDENT_ROW + index,
        sequence: index + 1,
      })),
      ...femaleStudents.map((student, index) => ({
        student,
        rowNumber: SF2_FIRST_FEMALE_STUDENT_ROW + index,
        sequence: index + 1,
      })),
    ];
    const weekdayIndices = new Map(weekdayDates.map((date, index) => [date, index]));
    for (const { student, rowNumber, sequence } of outputRows) {
      const row = sheet.getRow(rowNumber);
      row.hidden = false;
      row.getCell(1).value = sequence;
      row.getCell(2).value = student.name;
      for (const date of weekdayDates) {
        const dayIndex = weekdayIndices.get(date);
        const cell = row.getCell(dateColumns.get(date));
        if (attendance.noClassDates.includes(date)) {
          cell.value = "NC";
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFE6E6E6" },
          };
          continue;
        }
        if (student.lockedDates.includes(date)) {
          cell.value = "N/A";
          continue;
        }
        const status = normalizeAttendanceStatus(student.attendance[dayIndex]);
        if (!status) continue;
        const mark = SF2_MARK_BY_STATUS.get(status);
        if (mark === undefined) {
          throw new HttpError(400, `The attendance status "${status}" cannot be represented in SF2.`);
        }
        cell.value = mark || null;
      }
      let remarks = normalizeMonthlyAttendanceRemark(student.remarks) || student.remarks || "";
      const dateSpecificReasons = [];
      for (const date of weekdayDates) {
        const dayIndex = weekdayIndices.get(date);
        if (normalizeAttendanceStatus(student.attendance[dayIndex]) !== "Absent") continue;
        const reason = String(student.reasons[dayIndex] || "").trim();
        if (!reason) {
          throw new HttpError(400, `Add an absence reason for ${student.name} on ${date} before downloading SF2.`);
        }
        if (reason !== remarks) dateSpecificReasons.push(`${date}: ${reason}`);
      }
      if (dateSpecificReasons.length) {
        remarks += `${SF2_ABSENCE_REASONS_MARKER}${dateSpecificReasons.join("\n")}`;
      }
      if (normalizeMonthlyAttendanceRemark(student.remarks) && student.statusDate) {
        remarks += `${SF2_STATUS_DATE_MARKER}${student.statusDate}]`;
      }
      if (remarks.length > 500) {
        throw new HttpError(400, `Remarks and absence reasons for ${student.name} exceed the SF2 500-character limit.`);
      }
      row.getCell(SF2_REMARKS_COLUMN).value = remarks || null;
    }

    const studentRowsByNumber = new Map(outputRows.map(({ student, rowNumber }) => [rowNumber, student]));
    const maleCount = maleStudents.length;
    const femaleCount = femaleStudents.length;
    const formulaValue = (cell, formula, result) => {
      cell.value = { formula, result };
    };
    const registeredLearnerCount = (students) => students.reduce((count, student) => {
      const monthlyRemark = normalizeMonthlyAttendanceRemark(student.remarks);
      return count + (monthlyRemark === "NLS" || monthlyRemark === "Transferred Out" ? -1
        : monthlyRemark === "Transferred In" ? 1 : 0);
    }, students.length);
    const registeredCountFormula = (firstRow, lastRow) => {
      const names = `$B$${firstRow}:$B$${lastRow}`;
      const remarks = `$AH$${firstRow}:$AH$${lastRow}`;
      return `COUNTA(${names})-COUNTIF(${remarks},"NLS*")-COUNTIF(${remarks},"Transferred Out*")+COUNTIF(${remarks},"Transferred In*")`;
    };
    const maleRange = `$B$${SF2_FIRST_MALE_STUDENT_ROW}:$B$${SF2_LAST_MALE_STUDENT_ROW}`;
    const femaleRange = `$B$${SF2_FIRST_FEMALE_STUDENT_ROW}:$B$${SF2_LAST_FEMALE_STUDENT_ROW}`;

    for (const [firstRow, lastRow, students] of [
      [SF2_FIRST_MALE_STUDENT_ROW, SF2_LAST_MALE_STUDENT_ROW, maleStudents],
      [SF2_FIRST_FEMALE_STUDENT_ROW, SF2_LAST_FEMALE_STUDENT_ROW, femaleStudents],
    ]) {
      for (let rowNumber = firstRow; rowNumber <= lastRow; rowNumber += 1) {
        const student = studentRowsByNumber.get(rowNumber);
        let absences = 0;
        let countedDays = 0;
        if (student) {
          for (const date of weekdayDates) {
            if (attendance.noClassDates.includes(date)) continue;
            if (student.lockedDates.includes(date)) continue;
            countedDays += 1;
            const dayIndex = weekdayIndices.get(date);
            const status = normalizeAttendanceStatus(student.attendance[dayIndex]);
            if (status === "Absent") absences += 1;
            else if (status === "Half-day") absences += 0.5;
          }
        }
        formulaValue(
          sheet.getCell(`AF${rowNumber}`),
          `IF(B${rowNumber}="","",COUNTIF(G${rowNumber}:AE${rowNumber},"x")+COUNTIF(G${rowNumber}:AE${rowNumber},"h")*0.5)`,
          student ? absences : "",
        );
        formulaValue(
          sheet.getCell(`AG${rowNumber}`),
          `IF(B${rowNumber}="","",$AJ$117-COUNTIF(G${rowNumber}:AE${rowNumber},"N/A")-AF${rowNumber})`,
          student ? countedDays - absences : "",
        );
      }
      const totalRow = firstRow === SF2_FIRST_MALE_STUDENT_ROW ? 63 : 114;
      const registeredStudents = students.filter((student) => {
        const status = normalizeMonthlyAttendanceRemark(student.remarks);
        return status !== "NLS" && status !== "Transferred Out";
      });
      const registeredAbsentTotal = registeredStudents.reduce((total, student) => {
        return total + weekdayDates.reduce((studentTotal, date) => {
          if (attendance.noClassDates.includes(date) || student.lockedDates.includes(date)) return studentTotal;
          const status = normalizeAttendanceStatus(student.attendance[weekdayIndices.get(date)]);
          return studentTotal + (status === "Absent" ? 1 : status === "Half-day" ? 0.5 : 0);
        }, 0);
      }, 0);
      formulaValue(
        sheet.getCell(`AF${totalRow}`),
        `SUM(AF${firstRow}:AF${lastRow})-SUMIF(AH${firstRow}:AH${lastRow},"NLS*",AF${firstRow}:AF${lastRow})-SUMIF(AH${firstRow}:AH${lastRow},"Transferred Out*",AF${firstRow}:AF${lastRow})`,
        registeredAbsentTotal,
      );
      formulaValue(
        sheet.getCell(`AG${totalRow}`),
        `SUM(AG${firstRow}:AG${lastRow})-SUMIF(AH${firstRow}:AH${lastRow},"NLS*",AG${firstRow}:AG${lastRow})-SUMIF(AH${firstRow}:AH${lastRow},"Transferred Out*",AG${firstRow}:AG${lastRow})`,
        registeredStudents.reduce((total, student) => {
          const studentAbsences = weekdayDates.reduce((count, date) => {
            if (attendance.noClassDates.includes(date) || student.lockedDates.includes(date)) return count;
            const status = normalizeAttendanceStatus(student.attendance[weekdayIndices.get(date)]);
            return count + (status === "Absent" ? 1 : status === "Half-day" ? 0.5 : 0);
          }, 0);
          const studentClassDays = weekdayDates.filter((date) =>
            !attendance.noClassDates.includes(date) && !student.lockedDates.includes(date)).length;
          return total + studentClassDays - studentAbsences;
        }, 0),
      );
      for (let column = SF2_FIRST_DAY_COLUMN; column <= SF2_LAST_DAY_COLUMN; column += 1) {
        const letter = sheet.getColumn(column).letter;
        const date = weekdayDates[column - SF2_FIRST_DAY_COLUMN];
        const presentTotal = date && !attendance.noClassDates.includes(date)
          ? registeredStudents.reduce((count, student) => {
            if (student.lockedDates.includes(date)) return count;
            const status = normalizeAttendanceStatus(student.attendance[weekdayIndices.get(date)]);
            return count + (status === "Absent" ? 0 : status === "Half-day" ? 0.5 : 1);
          }, 0)
          : 0;
        const dailyFormula = date && !attendance.noClassDates.includes(date)
          ? `IF(${letter}10="","",COUNTA($B$${firstRow}:$B$${lastRow})-COUNTIF(${letter}${firstRow}:${letter}${lastRow},"x")-COUNTIF(${letter}${firstRow}:${letter}${lastRow},"h")*0.5-COUNTIF(${letter}${firstRow}:${letter}${lastRow},"N/A")-COUNTIF(AH${firstRow}:AH${lastRow},"NLS*")-COUNTIF(AH${firstRow}:AH${lastRow},"Transferred Out*"))`
          : `IF(${letter}10="","",0)`;
        formulaValue(sheet.getCell(`${letter}${totalRow}`), dailyFormula, date ? presentTotal : "");
      }
    }
    const combinedAbsent = (sheet.getCell("AF63").value.result || 0) + (sheet.getCell("AF114").value.result || 0);
    const combinedPresent = (sheet.getCell("AG63").value.result || 0) + (sheet.getCell("AG114").value.result || 0);
    formulaValue(sheet.getCell("AF115"), "AF63+AF114", combinedAbsent);
    formulaValue(sheet.getCell("AG115"), "AG63+AG114", combinedPresent);
    for (let column = SF2_FIRST_DAY_COLUMN; column <= SF2_LAST_DAY_COLUMN; column += 1) {
      const letter = sheet.getColumn(column).letter;
      const date = weekdayDates[column - SF2_FIRST_DAY_COLUMN];
      const combinedDaily = date && !attendance.noClassDates.includes(date)
        ? (sheet.getCell(`${letter}63`).value.result || 0) + (sheet.getCell(`${letter}114`).value.result || 0)
        : 0;
      formulaValue(sheet.getCell(`${letter}115`), `${letter}63+${letter}114`,
        date ? combinedDaily : "");
    }

    const summaryCounts = [
      ["AK", SF2_FIRST_MALE_STUDENT_ROW, SF2_LAST_MALE_STUDENT_ROW, maleStudents],
      ["AL", SF2_FIRST_FEMALE_STUDENT_ROW, SF2_LAST_FEMALE_STUDENT_ROW, femaleStudents],
    ];
    const registeredByColumn = new Map();
    for (const [column, firstRow, lastRow, students] of summaryCounts) {
      const count = students.length;
      const registeredCount = registeredLearnerCount(students);
      registeredByColumn.set(column, registeredCount);
      for (const rowNumber of [119, 120]) sheet.getCell(`${column}${rowNumber}`).value = count;
      formulaValue(
        sheet.getCell(`${column}123`),
        registeredCountFormula(firstRow, lastRow),
        registeredCount,
      );
    }
    sheet.getCell("AM119").value = { formula: "AK119+AL119", result: maleCount + femaleCount };
    sheet.getCell("AM120").value = { formula: "AK119+AL119", result: maleCount + femaleCount };
    const totalRegisteredCount = registeredByColumn.get("AK") + registeredByColumn.get("AL");
    formulaValue(sheet.getCell("AM123"), "AK123+AL123", totalRegisteredCount);
    for (const rowNumber of [125, 126]) {
      for (const [column, enrolled, enrollment] of [
        ["AK", registeredByColumn.get("AK"), maleCount],
        ["AL", registeredByColumn.get("AL"), femaleCount],
        ["AM", totalRegisteredCount, maleCount + femaleCount],
      ]) {
        formulaValue(
          sheet.getCell(`${column}${rowNumber}`),
          `IFERROR(${column}123/${column}119,0)`,
          enrollment ? enrolled / enrollment : 0,
        );
      }
    }
    for (const [column, students, presentTotal] of [
      ["AK", maleStudents, sheet.getCell("AG63").value.result],
      ["AL", femaleStudents, sheet.getCell("AG114").value.result],
      ["AM", attendance.students, combinedPresent],
    ]) {
      const averageDailyAttendance = classDayCount ? presentTotal / classDayCount : 0;
      formulaValue(
        sheet.getCell(`${column}127`),
        `IFERROR(AG${column === "AK" ? 63 : column === "AL" ? 114 : 115}/AJ117,0)`,
        averageDailyAttendance,
      );
      for (const rowNumber of [128, 129]) {
        formulaValue(
          sheet.getCell(`${column}${rowNumber}`),
          `IFERROR(${column}127/${column}123,0)`,
          registeredByColumn.get(column)
            ? averageDailyAttendance / registeredByColumn.get(column)
            : 0,
        );
      }
    }
    for (const [rowNumber, label] of [[132, "NLS"], [134, "Transferred Out"], [136, "Transferred In"]]) {
      const countLabel = (student) => normalizeMonthlyAttendanceRemark(student.remarks) === label ? 1 : 0;
      const maleTotal = maleStudents.reduce((total, student) => total + countLabel(student), 0);
      const femaleTotal = femaleStudents.reduce((total, student) => total + countLabel(student), 0);
      for (const row of [rowNumber, rowNumber + 1]) {
        formulaValue(
          sheet.getCell(`AK${row}`),
          `COUNTIF(AH${SF2_FIRST_MALE_STUDENT_ROW}:AH${SF2_LAST_MALE_STUDENT_ROW},"${label}")`,
          maleTotal,
        );
        formulaValue(
          sheet.getCell(`AL${row}`),
          `COUNTIF(AH${SF2_FIRST_FEMALE_STUDENT_ROW}:AH${SF2_LAST_FEMALE_STUDENT_ROW},"${label}")`,
          femaleTotal,
        );
        formulaValue(sheet.getCell(`AM${row}`), `AK${row}+AL${row}`, maleTotal + femaleTotal);
      }
    }
    workbook.calcProperties = { fullCalcOnLoad: true };

    return {
      buffer: Buffer.from(await workbook.xlsx.writeBuffer()),
      filename: `SF2_${new Date(`${month}-01T00:00:00.000Z`).toLocaleString("en-US", {
        month: "long",
        timeZone: "UTC",
      })}_${month.slice(0, 4)}.xlsx`,
    };
  }

  async function getAttendance(session, attendanceDate) {
    const date = validateAttendanceDate(attendanceDate);
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      const students = getOwnedStudents(studentSheet, studentColumns, session).map((student) => ({
        id: student.id,
        name: student.name,
        status: "",
        reason: "",
      }));
      const studentsById = new Map(students.map((student) => [student.id, student]));
      const statusEvents = monthlyAttendanceStatusEvents(workbook, session, studentsById.keys());
      const attendanceSheet = workbook.getWorksheet("Attendance Records");
      if (attendanceSheet) {
        const { columns } = getAttendanceSheet(workbook);
        for (let rowNumber = 2; rowNumber <= attendanceSheet.rowCount; rowNumber += 1) {
          const row = attendanceSheet.getRow(rowNumber);
          if (
            cellText(row.getCell(columns.classId).value) !== session.classId
            || cellText(row.getCell(columns.teacherUserId).value) !== session.userId
            || cellText(row.getCell(columns.date).value) !== date
          ) continue;
          const student = studentsById.get(cellText(row.getCell(columns.studentId).value));
          if (student && !isAttendanceDateLocked(statusEvents.get(student.id) || [], date)) {
            student.status = normalizeAttendanceStatus(row.getCell(columns.status).value);
            student.reason = cellText(row.getCell(columns.reason).value);
          }
        }
      }
      return { date, statuses: ATTENDANCE_STATUSES, students };
    });
  }

  async function saveAttendance(session, data) {
    const date = validateAttendanceDate(data.date);
    if (!Array.isArray(data.records)) {
      throw new HttpError(400, "Submit an attendance status for every student.");
    }
    const recordsByStudentId = new Map();
    for (const record of data.records) {
      const studentId = record && typeof record.studentId === "string" ? record.studentId.trim() : "";
      const status = record && typeof record.status === "string"
        ? normalizeAttendanceStatus(record.status)
        : "";
      if (!studentId || studentId.length > 40 || recordsByStudentId.has(studentId)) {
        throw new HttpError(400, "Each student must appear exactly once in the attendance list.");
      }
      if (!ATTENDANCE_STATUSES.includes(status)) {
        throw new HttpError(400, "Choose Present, Absent, Tardy, Half-day, or Excused for each student.");
      }
      recordsByStudentId.set(studentId, status);
    }

    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: studentSheet, columns: studentColumns } = getStudentSheet(workbook);
      const students = getOwnedStudents(studentSheet, studentColumns, session);
      if (recordsByStudentId.size !== students.length
        || students.some((student) => !recordsByStudentId.has(student.id))) {
        throw new HttpError(400, "Submit an attendance status for every student in your class.");
      }
      const statusEvents = monthlyAttendanceStatusEvents(workbook, session, students.map((student) => student.id));
      const lockedStudent = students.find((student) =>
        isAttendanceDateLocked(statusEvents.get(student.id) || [], date));
      if (lockedStudent) {
        throw new HttpError(400, `${lockedStudent.name}'s attendance is locked for ${date} because the learner is marked NLS or Transferred Out.`);
      }
      const calendar = getAttendanceCalendarSheet(workbook);
      for (let rowNumber = 2; rowNumber <= calendar.sheet.rowCount; rowNumber += 1) {
        const row = calendar.sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(calendar.columns.classId).value) === session.classId
          && cellText(row.getCell(calendar.columns.teacherUserId).value) === session.userId
          && attendanceDateText(row.getCell(calendar.columns.date).value) === date
          && cellText(row.getCell(calendar.columns.dayType).value) === "Holiday/No Class"
        ) {
          throw new HttpError(400, `${date} is marked Holiday/No Class. Attendance is locked.`);
        }
      }
      const { sheet, columns } = getAttendanceSheet(workbook);
      const rowsByStudentId = new Map();
      for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const row = sheet.getRow(rowNumber);
        if (
          cellText(row.getCell(columns.classId).value) === session.classId
          && cellText(row.getCell(columns.teacherUserId).value) === session.userId
          && cellText(row.getCell(columns.date).value) === date
        ) {
          rowsByStudentId.set(cellText(row.getCell(columns.studentId).value), row);
        }
      }
      for (const [studentId, status] of recordsByStudentId) {
        const existing = rowsByStudentId.get(studentId);
        if (existing) {
          existing.getCell(columns.status).value = status;
          existing.getCell(columns.updated).value = new Date();
          existing.getCell(columns.key).value = `${studentId}${date}`;
        } else {
          writeRow(sheet, [columns.id], {
            [columns.id]: `A${crypto.randomBytes(6).toString("hex").toUpperCase()}`,
            [columns.studentId]: studentId,
            [columns.classId]: session.classId,
            [columns.teacherUserId]: session.userId,
            [columns.date]: date,
            [columns.status]: status,
            [columns.updated]: new Date(),
            [columns.key]: `${studentId}${date}`,
          });
        }
      }
      await saveDatabaseWorkbook(workbook);
      return { saved: true, date };
    });
  }

  async function getTeacherProfile(session) {
    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: userSheet, columns: userColumns } = getUserSheet(workbook);
      let userRow = null;
      for (let rowNumber = 2; rowNumber <= userSheet.rowCount; rowNumber += 1) {
        const row = userSheet.getRow(rowNumber);
        if (
          cellText(row.getCell(userColumns.id).value) === session.userId
          && cellText(row.getCell(userColumns.role).value).toLowerCase() === "teacher"
        ) {
          userRow = row;
          break;
        }
      }
      if (!userRow) throw new HttpError(404, "The teacher profile could not be found.");

      const { sheet: classSheet, columns: classColumns } = getClassSheet(workbook);
      let classRow = null;
      for (let rowNumber = 2; rowNumber <= classSheet.rowCount; rowNumber += 1) {
        const row = classSheet.getRow(rowNumber);
        if (
          cellText(row.getCell(classColumns.id).value) === session.classId
          && cellText(row.getCell(classColumns.teacherId).value) === session.userId
        ) {
          classRow = row;
          break;
        }
      }
      if (!classRow) throw new HttpError(404, "The teacher's classroom could not be found.");
      const hasPhoto = await teacherPhotoExists(session.userId);
      return {
        name: cellText(userRow.getCell(userColumns.name).value),
        email: cellText(userRow.getCell(userColumns.email).value),
        classroomNumber: cellText(classRow.getCell(classColumns.classroom).value),
        grade: cellText(classRow.getCell(classColumns.grade).value),
        section: cellText(classRow.getCell(classColumns.section).value),
        schoolYear: cellText(classRow.getCell(classColumns.schoolYear).value),
        schoolName: cellText(classRow.getCell(classColumns.schoolName).value),
        schoolId: cellText(classRow.getCell(classColumns.schoolId).value),
        schoolHead: cellText(classRow.getCell(classColumns.schoolHead).value),
        hasPhoto,
      };
    });
  }

  async function updateTeacherProfile(session, data) {
    const name = typeof data.name === "string" ? data.name.trim() : "";
    const email = typeof data.email === "string" ? data.email.trim().toLowerCase() : "";
    const classroomNumber = typeof data.classroomNumber === "string" ? data.classroomNumber.trim() : "";
    const profileFields = [
      ["grade", "Grade", 100],
      ["section", "Section", 100],
      ["schoolYear", "School Year", 50],
      ["schoolName", "Name of School", 200],
      ["schoolId", "School ID", 20],
      ["schoolHead", "School Head", 120],
    ];
    const suppliedProfileFields = new Map();
    for (const [key, label, maxLength] of profileFields) {
      if (data[key] === undefined) continue;
      if (typeof data[key] !== "string" || data[key].trim().length > maxLength) {
        throw new HttpError(400, `Enter ${label} using no more than ${maxLength} characters.`);
      }
      suppliedProfileFields.set(key, data[key].trim());
    }
    if (!name || name.length > 120) {
      throw new HttpError(400, "Enter a name between 1 and 120 characters.");
    }
    if (!/^\S+@\S+\.\S+$/.test(email) || email.length > 254) {
      throw new HttpError(400, "Enter a valid email address.");
    }
    if (!classroomNumber || classroomNumber.length > 100) {
      throw new HttpError(400, "Enter a classroom number between 1 and 100 characters.");
    }

    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet: userSheet, columns: userColumns } = getUserSheet(workbook);
      let userRow = null;
      for (let rowNumber = 2; rowNumber <= userSheet.rowCount; rowNumber += 1) {
        const row = userSheet.getRow(rowNumber);
        if (cellText(row.getCell(userColumns.id).value) === session.userId) {
          userRow = row;
          continue;
        }
        if (cellText(row.getCell(userColumns.email).value).toLowerCase() === email) {
          throw new HttpError(409, "That email already belongs to another account.");
        }
      }
      if (!userRow || cellText(userRow.getCell(userColumns.role).value).toLowerCase() !== "teacher") {
        throw new HttpError(404, "The teacher profile could not be found.");
      }

      const { sheet: classSheet, columns: classColumns } = getClassSheet(workbook);
      let classRow = null;
      for (let rowNumber = 2; rowNumber <= classSheet.rowCount; rowNumber += 1) {
        const row = classSheet.getRow(rowNumber);
        if (
          cellText(row.getCell(classColumns.id).value) === session.classId
          && cellText(row.getCell(classColumns.teacherId).value) === session.userId
        ) {
          classRow = row;
          continue;
        }
        if (cellText(row.getCell(classColumns.classroom).value).toLowerCase() === classroomNumber.toLowerCase()) {
          throw new HttpError(409, "That classroom number is already registered.");
        }
      }
      if (!classRow) throw new HttpError(404, "The teacher's classroom could not be found.");

      userRow.getCell(userColumns.name).value = name;
      userRow.getCell(userColumns.email).value = email;
      classRow.getCell(classColumns.classroom).value = classroomNumber;
      classRow.getCell(classColumns.teacherName).value = name;
      for (const [key, value] of suppliedProfileFields) {
        classRow.getCell(classColumns[key]).value = value || null;
      }
      await saveDatabaseWorkbook(workbook);
      const activeSession = sessions.get(session.sessionId);
      if (activeSession) activeSession.name = name;
      return {
        name,
        email,
        classroomNumber,
        grade: cellText(classRow.getCell(classColumns.grade).value),
        section: cellText(classRow.getCell(classColumns.section).value),
        schoolYear: cellText(classRow.getCell(classColumns.schoolYear).value),
        schoolName: cellText(classRow.getCell(classColumns.schoolName).value),
        schoolId: cellText(classRow.getCell(classColumns.schoolId).value),
        schoolHead: cellText(classRow.getCell(classColumns.schoolHead).value),
      };
    });
  }

  async function updateTeacherPassword(session, data) {
    const currentPassword = typeof data.currentPassword === "string" ? data.currentPassword : "";
    const newPassword = typeof data.newPassword === "string" ? data.newPassword : "";
    if (currentPassword.length < 6 || currentPassword.length > 256) {
      throw new HttpError(400, "Enter your current password.");
    }
    if (newPassword.length < 6 || newPassword.length > 256) {
      throw new HttpError(400, "Use a new password between 6 and 256 characters.");
    }

    return serializeWorkbook(async () => {
      const workbook = await readDatabaseWorkbook();
      const { sheet, columns } = getUserSheet(workbook);
      let userRow = null;
      for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const row = sheet.getRow(rowNumber);
        if (cellText(row.getCell(columns.id).value) === session.userId) {
          userRow = row;
          break;
        }
      }
      if (!userRow || cellText(userRow.getCell(columns.role).value).toLowerCase() !== "teacher") {
        throw new HttpError(404, "The teacher profile could not be found.");
      }
      const matches = await verifyPassword(currentPassword, cellText(userRow.getCell(columns.password).value));
      if (!matches) throw new HttpError(400, "Your current password doesn't match.");
      userRow.getCell(columns.password).value = await hashPassword(newPassword);
      setPasswordDocumentation(workbook, columns.password);
      await saveDatabaseWorkbook(workbook);
      return { saved: true };
    });
  }

  function teacherPhotoPath(userId, extension) {
    if (!/^T\d{3,}$/i.test(userId)) throw new HttpError(404, "The profile picture was not found.");
    return path.join(profilePhotoDirectory, `${userId}.${extension}`);
  }

  async function teacherPhotoExists(userId) {
    for (const extension of ["jpg", "png", "webp"]) {
      try {
        const photo = await fs.stat(teacherPhotoPath(userId, extension));
        if (photo.isFile()) return true;
      } catch (error) {
        if (error.code !== "ENOENT") throw error;
      }
    }
    return false;
  }

  async function findTeacherPhoto(userId) {
    for (const extension of ["jpg", "png", "webp"]) {
      const filePath = teacherPhotoPath(userId, extension);
      try {
        const buffer = await fs.readFile(filePath);
        return { filePath, buffer, ...getProfilePhotoFormat(buffer) };
      } catch (error) {
        if (error.code !== "ENOENT") throw error;
      }
    }
    throw new HttpError(404, "No profile picture has been uploaded.");
  }

  async function saveTeacherPhoto(session, request) {
    const photo = await readProfilePhoto(request);
    const filePath = teacherPhotoPath(session.userId, photo.extension);
    await fs.mkdir(profilePhotoDirectory, { recursive: true });
    await fs.writeFile(filePath, photo.buffer);
    for (const extension of ["jpg", "png", "webp"]) {
      if (extension === photo.extension) continue;
      try {
        await fs.unlink(teacherPhotoPath(session.userId, extension));
      } catch (error) {
        if (error.code !== "ENOENT") throw error;
      }
    }
  }

  async function removeTeacherPhoto(session) {
    for (const extension of ["jpg", "png", "webp"]) {
      try {
        await fs.unlink(teacherPhotoPath(session.userId, extension));
      } catch (error) {
        if (error.code !== "ENOENT") throw error;
      }
    }
  }

  const server = http.createServer(async (request, response) => {
    try {
      const url = new URL(request.url, "http://localhost");
      if (request.method === "GET" && (url.pathname === "/" || url.pathname === "/index.html")) {
        const html = await fs.readFile(INDEX_PATH);
        response.writeHead(200, {
          "Content-Type": "text/html; charset=utf-8",
          "Cache-Control": "no-store",
          "X-Content-Type-Options": "nosniff",
        });
        response.end(html);
        return;
      }

      if (request.method === "GET" && url.pathname === "/api/health") {
        try {
          await serializeWorkbook(checkWorkbook);
          sendJson(response, 200, { status: "ok", storage: "sqlite" });
        } catch (error) {
          console.error("Portal database health check failed:", error);
          sendJson(response, 503, {
            error: "The portal is running, but its SQLite database is unavailable. Check that the database file exists and is accessible.",
          });
        }
        return;
      }

      if (request.method === "POST" && url.pathname === "/api/auth/register") {
        const result = await register(await readJson(request));
        openSession(response, request, result.user);
        const { id, classId, ...user } = result.user;
        sendJson(response, 201, { user });
        return;
      }

      if (request.method === "POST" && url.pathname === "/api/auth/login") {
        const result = await login(await readJson(request));
        openSession(response, request, result.user);
        const { id, classId, ...user } = result.user;
        sendJson(response, 200, { user });
        return;
      }

      if (request.method === "GET" && url.pathname === "/api/auth/session") {
        const session = getSession(request);
        if (!session) {
          sendJson(response, 401, { error: "Your session has expired. Sign in again." });
          return;
        }
        sendJson(response, 200, {
          user: {
            name: session.name,
            role: session.role,
            classCode: session.classCode,
          },
        });
        return;
      }

      if (request.method === "POST" && url.pathname === "/api/auth/logout") {
        const session = getSession(request);
        if (session) sessions.delete(session.sessionId);
        response.setHeader("Set-Cookie", `${SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0`);
        sendJson(response, 200, { ok: true });
        return;
      }

      if (request.method === "GET" && url.pathname === "/api/teacher/profile") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await getTeacherProfile(session));
        return;
      }

      if (request.method === "GET" && url.pathname === "/api/parent/portal") {
        const session = requireParentSession(request);
        sendJson(response, 200, await getParentPortal(session));
        return;
      }

      const parentLinkMatch = url.pathname.match(/^\/api\/students\/([^/]+)\/parent$/);
      if (request.method === "PUT" && parentLinkMatch) {
        const session = requireTeacherSession(request);
        const result = await linkStudentParent(
          session,
          decodeURIComponent(parentLinkMatch[1]),
          await readJson(request),
        );
        sendJson(response, 200, result);
        return;
      }

      if (request.method === "DELETE" && parentLinkMatch) {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await unlinkStudentParent(
          session,
          decodeURIComponent(parentLinkMatch[1]),
        ));
        return;
      }

      if (request.method === "PUT" && url.pathname === "/api/teacher/profile") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await updateTeacherProfile(session, await readJson(request)));
        return;
      }

      if (request.method === "PUT" && url.pathname === "/api/teacher/profile/password") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await updateTeacherPassword(session, await readJson(request)));
        return;
      }

      if (request.method === "GET" && url.pathname === "/api/teacher/profile/photo") {
        const session = requireTeacherSession(request);
        const photo = await findTeacherPhoto(session.userId);
        response.writeHead(200, {
          "Content-Type": photo.contentType,
          "Content-Length": photo.buffer.length,
          "Cache-Control": "private, no-store",
          "X-Content-Type-Options": "nosniff",
        });
        response.end(photo.buffer);
        return;
      }

      if (request.method === "PUT" && url.pathname === "/api/teacher/profile/photo") {
        const session = requireTeacherSession(request);
        await saveTeacherPhoto(session, request);
        sendJson(response, 200, { saved: true });
        return;
      }

      if (request.method === "DELETE" && url.pathname === "/api/teacher/profile/photo") {
        const session = requireTeacherSession(request);
        await removeTeacherPhoto(session);
        sendJson(response, 200, { removed: true });
        return;
      }

      if (request.method === "GET" && url.pathname === "/api/students") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await listStudents(session));
        return;
      }

      if (request.method === "POST" && url.pathname === "/api/students") {
        const session = requireTeacherSession(request);
        const student = await addStudent(session, await readJson(request));
        sendJson(response, 201, { student });
        return;
      }

      if (request.method === "POST" && url.pathname === "/api/students/clear/verify") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await verifyRosterClearPassword(session, await readJson(request)));
        return;
      }

      if (request.method === "DELETE" && url.pathname === "/api/students/clear") {
        const session = requireTeacherSession(request);
        const data = await readJson(request);
        if (data.confirmed !== true) throw new HttpError(400, "Confirm that you want to clear the student roster.");
        sendJson(response, 200, await clearStudentRoster(session, data));
        return;
      }

      if (request.method === "POST" && url.pathname === "/api/students/import") {
        const session = requireTeacherSession(request);
        const upload = await readSF1Upload(request);
        const result = await importStudents(session, upload);
        sendJson(response, 200, result);
        return;
      }

      if (request.method === "GET" && url.pathname === "/api/grades") {
        const session = requireTeacherSession(request);
        const result = await getStudentGrades(
          session,
          url.searchParams.get("studentId") || "",
          url.searchParams.get("term") || "",
        );
        sendJson(response, 200, result);
        return;
      }

      if (request.method === "GET" && url.pathname === "/api/grades/lock") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await getStudentGradeLock(
          session,
          url.searchParams.get("studentId") || "",
          url.searchParams.get("term") || "",
        ));
        return;
      }

      if (request.method === "POST" && url.pathname === "/api/grades/lock") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await setStudentGradeLock(session, await readJson(request)));
        return;
      }

      if (request.method === "GET" && url.pathname === "/api/grades/summary") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await getGradeSummary(session));
        return;
      }

      if (request.method === "POST" && url.pathname === "/api/grades") {
        const session = requireTeacherSession(request);
        const result = await saveStudentGrades(session, await readJson(request));
        sendJson(response, 200, result);
        return;
      }

      if (request.method === "GET" && url.pathname === "/api/reading-assessments") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await getReadingAssessmentSummary(
          session,
          url.searchParams.get("period") || "BOSY",
          url.searchParams.has("attemptNo") ? Number(url.searchParams.get("attemptNo")) : 1,
        ));
        return;
      }

      if (request.method === "POST" && url.pathname === "/api/reading-assessments") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await saveReadingAssessment(session, await readJson(request)));
        return;
      }

      if (request.method === "GET" && url.pathname === "/api/attendance") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await getAttendance(session, url.searchParams.get("date") || ""));
        return;
      }

      if (request.method === "POST" && url.pathname === "/api/attendance") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await saveAttendance(session, await readJson(request)));
        return;
      }

      if (request.method === "POST" && url.pathname === "/api/attendance/month/import") {
        const session = requireTeacherSession(request);
        const month = validateAttendanceMonth(url.searchParams.get("month") || "");
        const buffer = await readSF2Upload(request);
        sendJson(response, 200, await importAttendanceSF2(session, month, buffer));
        return;
      }

      if (request.method === "GET" && url.pathname === "/api/attendance/month/export") {
        const session = requireTeacherSession(request);
        const { buffer, filename } = await exportAttendanceSF2(
          session,
          url.searchParams.get("month") || "",
        );
        response.writeHead(200, {
          "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "Content-Disposition": `attachment; filename="${filename}"`,
          "Cache-Control": "no-store",
          "X-Content-Type-Options": "nosniff",
        });
        response.end(buffer);
        return;
      }

      if (request.method === "GET" && url.pathname === "/api/attendance/month") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await getAttendanceMonth(session, url.searchParams.get("month") || ""));
        return;
      }

      if (request.method === "GET" && url.pathname === "/api/attendance/summary") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await getAttendanceSummary(session));
        return;
      }

      if (request.method === "POST" && url.pathname === "/api/attendance/month") {
        const session = requireTeacherSession(request);
        sendJson(response, 200, await saveAttendanceMonth(session, await readJson(request, 2 * 1024 * 1024)));
        return;
      }

      sendJson(response, 404, { error: "Page not found." });
    } catch (error) {
      if (error instanceof HttpError) {
        sendJson(response, error.status, { error: error.message });
        return;
      }
      console.error("Portal request failed:", error);
      sendJson(response, 500, {
        error: "The portal database could not be read or updated. Check the server log for details, then try again.",
      });
    }
  });
  server.on("close", () => database?.close());
  return server;
}

if (require.main === module) {
  const configuredPort = process.env.PORT ? Number(process.env.PORT) : 3003;
  const host = process.env.HOST || "127.0.0.1";
  const maxPort = process.env.PORT ? configuredPort : configuredPort + 20;

  function listen(port) {
    const server = createServer();
    server.once("error", (error) => {
      if (!process.env.PORT && error.code === "EADDRINUSE" && port < maxPort) {
        console.warn(`Port ${port} is already in use; trying ${port + 1}.`);
        listen(port + 1);
        return;
      }
      console.error(`Could not start the portal on ${host}:${port}:`, error.message);
      process.exitCode = 1;
    });
    server.listen(port, host, () => {
      console.log(`Using SQLite backend: ${path.resolve(DATABASE_PATH)}`);
      console.log(`Class Adviser's Portal is running at http://${host}:${port}`);
    });
  }

  listen(configuredPort);
}

module.exports = {
  createServer,
  readPortalDatabaseWorkbook,
  writePortalDatabaseWorkbook,
};
