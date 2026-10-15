"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");
const ExcelJS = require("exceljs");
const XLSX = require("@e965/xlsx");
const { createServer, readPortalDatabaseWorkbook, writePortalDatabaseWorkbook } = require("../server");

const sourceWorkbook = path.join(__dirname, "..", "grade5-portal-database.xlsx");
const sourceSF2Template = path.join(__dirname, "..", "SF2.xlsx");
const sf2StudentRows = [
  ...Array.from({ length: 50 }, (_, index) => 13 + index),
  ...Array.from({ length: 50 }, (_, index) => 64 + index),
];

function attendanceWeekdays(month) {
  const [year, monthNumber] = month.split("-").map(Number);
  const dayCount = new Date(Date.UTC(year, monthNumber, 0)).getUTCDate();
  return Array.from({ length: dayCount }, (_, index) => {
    const day = index + 1;
    const date = `${month}-${String(day).padStart(2, "0")}`;
    const weekday = new Date(`${date}T00:00:00.000Z`).getUTCDay();
    return weekday !== 0 && weekday !== 6 ? date : null;
  }).filter(Boolean);
}

async function createSF2Upload(month, learners) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(sourceSF2Template);
  const sheet = workbook.getWorksheet("OCT 2026");
  assert.ok(sheet, "SF2 template worksheet exists");
  const [year, monthNumber] = month.split("-").map(Number);
  const monthNames = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  sheet.name = `${monthNames[monthNumber - 1]} ${year}`;
  for (let column = 7; column <= 31; column += 1) {
    sheet.getCell(10, column).value = null;
    sheet.getCell(11, column).value = ["M", "T", "W", "TH", "F"][(column - 7) % 5];
  }
  const schoolDays = [];
  const dayCount = new Date(Date.UTC(year, monthNumber, 0)).getUTCDate();
  for (let day = 1; day <= dayCount; day += 1) {
    const weekday = new Date(Date.UTC(year, monthNumber - 1, day)).getUTCDay();
    if (weekday === 0 || weekday === 6) continue;
    const column = 7 + schoolDays.length;
    sheet.getCell(10, column).value = day;
    sheet.getCell(11, column).value = ["M", "T", "W", "TH", "F"][(weekday + 6) % 7];
    schoolDays.push({ day, column });
  }
  for (const rowNumber of sf2StudentRows) {
    sheet.getCell(rowNumber, 2).value = null;
    for (let column = 7; column <= 31; column += 1) {
      sheet.getCell(rowNumber, column).value = null;
    }
    sheet.getCell(rowNumber, 34).value = null;
  }
  const nextStudentRow = { Male: 13, Female: 64 };
  learners.forEach((learner) => {
    const gender = learner.gender || "Male";
    const row = sheet.getRow(nextStudentRow[gender]);
    nextStudentRow[gender] += 1;
    row.getCell(2).value = learner.name;
    row.getCell(34).value = learner.remarks || null;
    for (const { day, column } of schoolDays) {
      row.getCell(column).value = learner.marks && learner.marks[day] ? learner.marks[day] : null;
    }
  });
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

async function createEmptyTestWorkbook(databasePath) {
  await fs.copyFile(sourceWorkbook, databasePath);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(databasePath);
  const clearColumns = (sheetName, columns) => {
    const sheet = workbook.getWorksheet(sheetName);
    assert.ok(sheet, `${sheetName} worksheet exists`);
    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
      for (const column of columns) {
        const cell = sheet.getRow(rowNumber).getCell(column);
        if (!cell.value || typeof cell.value !== "object" || !("formula" in cell.value || "sharedFormula" in cell.value)) {
          cell.value = null;
        }
      }
    }
  };

  const users = workbook.getWorksheet("Users");
  const passwordHeader = users.getRow(1).values.findIndex((value) =>
    typeof value === "string" && value.toLowerCase().startsWith("password"),
  );
  assert.ok(passwordHeader > 0, "Users worksheet has a password column");
  clearColumns("Users", Array.from({ length: 8 }, (_, index) => index + 1));

  const classes = workbook.getWorksheet("Classes");
  clearColumns("Classes", [1, 2, 3, 4]);
  const students = workbook.getWorksheet("Students");
  clearColumns("Students", [...Array.from({ length: 4 }, (_, index) => index + 1),
    ...Array.from({ length: Math.max(0, students.columnCount - 6) }, (_, index) => index + 7)]);
  const termGrades = workbook.getWorksheet("Term Grades");
  if (termGrades) clearColumns("Term Grades", Array.from({ length: 8 }, (_, index) => index + 1));
  clearColumns("SignInLog", [1, 2, 3, 4, 6]);

  await workbook.xlsx.writeFile(databasePath);
}

test("registrations persist in the workbook and can sign in after a restart", async (context) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "grade5-portal-"));
  const databasePath = path.join(directory, "database.sqlite");
  const migrationWorkbookPath = path.join(directory, "migration.xlsx");
  await createEmptyTestWorkbook(migrationWorkbookPath);
  const emptyWorkbook = new ExcelJS.Workbook();
  await emptyWorkbook.xlsx.readFile(migrationWorkbookPath);
  const emptyUsers = emptyWorkbook.getWorksheet("Users");
  const emptyUserHeaders = new Map();
  emptyUsers.getRow(1).eachCell((cell, column) => {
    emptyUserHeaders.set(String(cell.value).toLowerCase(), column);
  });
  for (let rowNumber = 2; rowNumber <= emptyUsers.rowCount; rowNumber += 1) {
    const row = emptyUsers.getRow(rowNumber);
    for (let column = 1; column <= 8; column += 1) {
      assert.equal(row.getCell(column).value || "", "", `user data is empty in row ${rowNumber}, column ${column}`);
    }
  }

  let server = createServer({ databasePath, migrationWorkbookPath });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  context.after(async () => {
    if (server.listening) await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
    await fs.rm(directory, { recursive: true, force: true });
  });

  const baseUrl = `http://127.0.0.1:${server.address().port}`;
  const health = await fetch(`${baseUrl}/api/health`);
  assert.equal(health.status, 200);
  assert.deepEqual(await health.json(), { status: "ok", storage: "sqlite" });
  const portalPageResponse = await fetch(baseUrl);
  assert.equal(portalPageResponse.status, 200);
  const portalPage = await portalPageResponse.text();
  assert.match(portalPage, /sidebar-collapsed \.sidebar\{flex-basis:0;width:0/);
  assert.match(portalPage, /window\.matchMedia\("\(max-width: 760px\)"\)/);
  assert.match(portalPage, /toggle\("sidebar-collapsed"\)/);
  assert.match(portalPage, /<fieldset class="grade-fieldset" id="gradeFields">/);
  assert.match(portalPage, /id="teacherProfileSchoolName"/);
  assert.match(portalPage, /id="teacherProfileSchoolId"/);
  assert.match(portalPage, /id="clearStudentRoster"[^>]*>Clear List<\/button>/);
  assert.match(portalPage, /id="refreshStudentRoster"[^>]*>Refresh roster<\/button>/);
  assert.match(portalPage, /id="clearRosterPassword" type="password"/);
  assert.match(portalPage, /id="confirmClearRosterForm" hidden/);
  assert.match(portalPage, /id="editGrades"[^>]*>Edit Grades<\/button>/);
  assert.match(portalPage, /\$\("gradeFields"\)\.disabled=disabled\|\|gradeSaving/);
  assert.match(portalPage, /new Option\("Holiday\/No Class","no-class"\)/);
  assert.match(portalPage, /new Option\("NLS - No longer in school","NLS"\)/);
  assert.match(portalPage, /new Option\("Transferred Out","Transferred Out"\)/);
  assert.match(portalPage, /new Option\("Transferred In","Transferred In"\)/);
  assert.match(portalPage, /className="attendance-status-date"/);
  assert.match(portalPage, /statusDate:student\.statusDate\|\|""/);
  assert.match(portalPage, /id="attendanceSummaryTab"[^>]*>Summary Of Attendance<\/button>/);
  assert.match(portalPage, /id="attendanceSummaryRows"/);
  assert.match(portalPage, /\.attendance-summary-table th:nth-child\(2\),\.attendance-summary-table td:nth-child\(2\)\{position:sticky;left:48px/);
  assert.match(portalPage, /api\("\/api\/attendance\/summary"\)/);
  assert.match(portalPage, /Registered Learner as of End of the Month count/);
  assert.match(portalPage, /noClassMark\.textContent="NC"/);
  assert.match(portalPage, /\.attendance-table td\.attendance-no-class-cell\{background:#e6e6e6/);
  assert.match(portalPage, /control\.disabled=isNoClass/);
  assert.match(portalPage, /fetch\("\/api\/attendance\/month\/export\?month="\+encodeURIComponent\(month\)\)/);
  assert.match(portalPage, /link\.download=filename;[\s\S]*?link\.click\(\);/);
  assert.ok(portalPage.indexOf('id="navSettings"') > portalPage.indexOf('id="navTeacherProfile"'));
  assert.match(portalPage, /id="settingsPage"/);
  assert.match(portalPage, /id="appearanceSetting"/);
  assert.match(portalPage, /<nav id="parentSidebarNav" aria-label="Learner information navigation" hidden>[\s\S]*?data-parent-target="parentStudentInfo" aria-current="page"[\s\S]*?data-parent-target="parentAttendanceSection"[\s\S]*?data-parent-target="parentReadingSection"[\s\S]*?data-parent-target="parentGradesSection"[\s\S]*?<\/nav>/);
  assert.match(portalPage, /\$\("parentSidebarNav"\)\.hidden=user\.role!=="parent"/);
  assert.match(portalPage, /function showParentSection\(targetId\)\{[\s\S]*?section\.hidden=!selected[\s\S]*?button\.setAttribute\("aria-current","page"\)/);
  assert.match(portalPage, /showParentSection\("parentStudentInfo"\)/);
  assert.match(portalPage, /id="parentAttendanceSection" hidden/);
  assert.match(portalPage, /id="parentReadingSection" hidden/);
  assert.match(portalPage, /id="parentGradesSection" hidden/);
  assert.match(portalPage, /id="menuToggle"[^>]*aria-controls="sidebar"/);
  assert.doesNotMatch(portalPage, /parentSectionMenuToggle|parentSectionNav|parent-section-menu|parent-quick-nav/);
  assert.doesNotMatch(portalPage, /class="parent-record-layout"/);
  assert.match(portalPage, /data-parent-target="parentStudentInfo" aria-current="page">[\s\S]*?Student's Information/);
  assert.match(portalPage, /data-parent-target="parentAttendanceSection">[\s\S]*?Attendance/);
  assert.match(portalPage, /data-parent-target="parentReadingSection">[\s\S]*?Reading Progress/);
  assert.match(portalPage, /data-parent-target="parentGradesSection">[\s\S]*?Term Grades/);
  assert.match(portalPage, /BOSY means beginning of school year, MOSY means middle of school year, and EOSY means end of school year/);
  assert.match(portalPage, /parent-reading-grid\{display:grid;width:100%;grid-template-columns:minmax\(0,1fr\);grid-auto-rows:1fr;align-items:stretch/);
  assert.doesNotMatch(portalPage, /parent-reading-grid\{grid-template-columns:repeat\([23],minmax\(0,1fr\)\)/);
  assert.match(portalPage, /parent-reading-period\{display:flex;min-width:0;height:100%;flex-direction:column/);
  assert.match(portalPage, /parent-reading-period\.period-bosy\{--period-color:/);
  assert.match(portalPage, /parent-reading-period\.period-mosy\{--period-color:/);
  assert.match(portalPage, /parent-reading-period\.period-eosy\{--period-color:/);
  assert.match(portalPage, /card\.className="parent-reading-period period-"\+period\.toLowerCase\(\)/);
  assert.match(portalPage, /parent-reading-period h3\{display:inline-flex;align-self:flex-start/);
  assert.match(portalPage, /parent-reading-period-caption/);
  assert.match(portalPage, /id="parentReadingTrendChart" class="parent-reading-trend-chart"/);
  assert.match(portalPage, /id="parentReadingDateTrendChart" class="parent-reading-date-trend-chart"/);
  assert.match(portalPage, /id="parentReadingDateTrendEmpty"/);
  assert.match(portalPage, /id="parentReadingTrendEmpty"/);
  assert.match(portalPage, /function renderParentReadingTrend\(learner\)\{[\s\S]*?class:"parent-reading-trend-series"[\s\S]*?class:"parent-reading-trend-point"/);
  assert.match(portalPage, /parent-reading-date-trend-series/);
  assert.match(portalPage, /parent-reading-date-trend-point/);
  assert.match(portalPage, /function parseReadingDate\(value\)/);
  assert.match(portalPage, /formatReadingDate\(result\.assessedDate\)/);
  assert.match(portalPage, /if\(!point\|\|index===0\|\|!points\[index-1\]\)return/);
  assert.match(portalPage, /chart\.setAttribute\("aria-label","Reading progress by assessment period\. "\+accessibleSummary\.join\("; "\)\)/);
  assert.match(portalPage, /steps\.style\.setProperty\("--reading-progress",Math\.max\(0,levelIndex\)\/\(levels\.length-1\)\*100\+"%"\)/);
  assert.match(portalPage, /if\(index<=levelIndex\)step\.classList\.add\("is-complete"\)/);
  const appearanceOptions = portalPage.match(/id="appearanceSetting">([\s\S]*?)<\/select>/)[1];
  const appearanceValues = Array.from(appearanceOptions.matchAll(/<option value="([^"]+)"/g), (match) => match[1]);
  assert.deepEqual(appearanceValues.slice(0, 3), ["system", "light", "dark"]);
  assert.equal(appearanceValues.length - 3, 20, "Appearance includes 20 custom themes in addition to the standard choices");
  for (const theme of appearanceValues.slice(3)) {
    assert.ok(portalPage.includes(`:root[data-theme="${theme}"]`), `${theme} has a theme palette`);
  }

  const request = (route, body) => fetch(`${baseUrl}${route}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const teacherCredentials = {
    name: "Taylor Teacher",
    email: "taylor.teacher@example.com",
    password: "secure-pass-123",
    role: "teacher",
    extra: "12",
  };
  const teacherRegistration = await request("/api/auth/register", teacherCredentials);
  assert.equal(teacherRegistration.status, 201);
  const teacherCookie = teacherRegistration.headers.get("set-cookie").split(";")[0];
  const teacherResult = await teacherRegistration.json();
  assert.equal(teacherResult.user.name, teacherCredentials.name);
  assert.equal(teacherResult.user.role, "teacher");
  assert.match(teacherResult.user.classCode, /^G5-[A-F0-9]{8}$/);

  const profileResponse = await fetch(`${baseUrl}/api/teacher/profile`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(profileResponse.status, 200);
  assert.deepEqual(await profileResponse.json(), {
    name: teacherCredentials.name,
    email: teacherCredentials.email,
    classroomNumber: "12",
    grade: "",
    section: "",
    schoolYear: "",
    schoolName: "",
    schoolId: "",
    schoolHead: "",
    hasPhoto: false,
  });

  const updateProfile = (cookie, data) => fetch(`${baseUrl}/api/teacher/profile`, {
    method: "PUT",
    headers: { "Content-Type": "application/json", Cookie: cookie },
    body: JSON.stringify(data),
  });
  const updatedProfileResponse = await updateProfile(teacherCookie, {
    name: "Taylor Updated",
    email: "taylor.updated@example.com",
    classroomNumber: "12-A",
    grade: "Grade 5",
    section: "Mabini",
    schoolYear: "2026-2027",
    schoolName: "Sample Elementary School",
    schoolId: "123456",
    schoolHead: "Alex Principal",
  });
  assert.equal(updatedProfileResponse.status, 200);
  assert.deepEqual(await updatedProfileResponse.json(), {
    name: "Taylor Updated",
    email: "taylor.updated@example.com",
    classroomNumber: "12-A",
    grade: "Grade 5",
    section: "Mabini",
    schoolYear: "2026-2027",
    schoolName: "Sample Elementary School",
    schoolId: "123456",
    schoolHead: "Alex Principal",
  });
  const updatedTeacherLogin = await request("/api/auth/login", {
    email: "taylor.updated@example.com",
    password: teacherCredentials.password,
    role: "teacher",
  });
  assert.equal(updatedTeacherLogin.status, 200);
  const updatedTeacherCookie = updatedTeacherLogin.headers.get("set-cookie").split(";")[0];
  const restoreProfileResponse = await updateProfile(updatedTeacherCookie, {
    name: teacherCredentials.name,
    email: teacherCredentials.email,
    classroomNumber: "12",
    grade: "Grade 5",
    section: "Mabini",
    schoolYear: "2026-2027",
    schoolName: "Sample Elementary School",
    schoolId: "123456",
    schoolHead: "Alex Principal",
  });
  assert.equal(restoreProfileResponse.status, 200);

  const unauthorizedPhoto = await fetch(`${baseUrl}/api/teacher/profile/photo`);
  assert.equal(unauthorizedPhoto.status, 401);
  const missingPhoto = await fetch(`${baseUrl}/api/teacher/profile/photo`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(missingPhoto.status, 404);
  const profilePng = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j5l8AAAAASUVORK5CYII=",
    "base64",
  );
  const savePhotoResponse = await fetch(`${baseUrl}/api/teacher/profile/photo`, {
    method: "PUT",
    headers: { "Content-Type": "image/png", Cookie: teacherCookie },
    body: profilePng,
  });
  assert.equal(savePhotoResponse.status, 200);
  const getPhotoResponse = await fetch(`${baseUrl}/api/teacher/profile/photo`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(getPhotoResponse.status, 200);
  assert.equal(getPhotoResponse.headers.get("content-type"), "image/png");
  assert.deepEqual(Buffer.from(await getPhotoResponse.arrayBuffer()), profilePng);
  const profileWithPhoto = await fetch(`${baseUrl}/api/teacher/profile`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal((await profileWithPhoto.json()).hasPhoto, true);
  const invalidPhotoResponse = await fetch(`${baseUrl}/api/teacher/profile/photo`, {
    method: "PUT",
    headers: { "Content-Type": "image/png", Cookie: teacherCookie },
    body: Buffer.from("not a png"),
  });
  assert.equal(invalidPhotoResponse.status, 400);
  const removePhotoResponse = await fetch(`${baseUrl}/api/teacher/profile/photo`, {
    method: "DELETE",
    headers: { Cookie: teacherCookie },
  });
  assert.equal(removePhotoResponse.status, 200);
  const removedPhotoResponse = await fetch(`${baseUrl}/api/teacher/profile/photo`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(removedPhotoResponse.status, 404);
  const profileWithoutPhoto = await fetch(`${baseUrl}/api/teacher/profile`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal((await profileWithoutPhoto.json()).hasPhoto, false);
  const incorrectCurrentPassword = await fetch(`${baseUrl}/api/teacher/profile/password`, {
    method: "PUT",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      currentPassword: "incorrect-password",
      newPassword: "updated-secure-pass-456",
    }),
  });
  assert.equal(incorrectCurrentPassword.status, 400);
  const updatedPassword = "updated-secure-pass-456";
  const updatePasswordResponse = await fetch(`${baseUrl}/api/teacher/profile/password`, {
    method: "PUT",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      currentPassword: teacherCredentials.password,
      newPassword: updatedPassword,
    }),
  });
  assert.equal(updatePasswordResponse.status, 200);
  const oldPasswordLogin = await request("/api/auth/login", {
    email: teacherCredentials.email,
    password: teacherCredentials.password,
    role: "teacher",
  });
  assert.equal(oldPasswordLogin.status, 401);
  const newPasswordLogin = await request("/api/auth/login", {
    email: teacherCredentials.email,
    password: updatedPassword,
    role: "teacher",
  });
  assert.equal(newPasswordLogin.status, 200);
  teacherCredentials.password = updatedPassword;

  const studentDetails = {
    lrn: "123456789012",
    lastName: "Santos",
    firstName: "Alex",
    middleName: "Reyes",
    gender: "Female",
    birthdate: "2015-04-21",
    mothersMaidenName: "Cruz",
    guardianRelationship: "Aunt",
    father: "Jordan Santos",
    currentResidence: "Barangay Central",
    religion: "Catholic",
    motherTongue: "Cebuano",
    dialects: "Chavacano",
    ethnicities: "Filipino",
  };
  const unauthenticatedStudents = await fetch(`${baseUrl}/api/students`);
  assert.equal(unauthenticatedStudents.status, 401);

  const addStudent = await fetch(`${baseUrl}/api/students`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify(studentDetails),
  });
  assert.equal(addStudent.status, 201);
  const addedStudent = await addStudent.json();
  assert.equal(addedStudent.student.firstName, studentDetails.firstName);
  assert.equal(addedStudent.student.birthdate, studentDetails.birthdate);

  const legacyRosterWorkbook = readPortalDatabaseWorkbook(databasePath);
  const legacyStudentsSheet = legacyRosterWorkbook.getWorksheet("Students");
  const usersSheet = legacyRosterWorkbook.getWorksheet("Users");
  const userHeaders = new Map();
  usersSheet.getRow(1).eachCell((cell, column) => {
    userHeaders.set(String(cell.value).toLowerCase(), column);
  });
  let registeredTeacherId = "";
  for (let rowNumber = 2; rowNumber <= usersSheet.rowCount; rowNumber += 1) {
    const row = usersSheet.getRow(rowNumber);
    if (String(row.getCell(userHeaders.get("email")).value).toLowerCase() === teacherCredentials.email) {
      registeredTeacherId = String(row.getCell(userHeaders.get("user id")).value);
      break;
    }
  }
  assert.ok(registeredTeacherId, "registered teacher has a workbook user id");
  const legacyHeaders = new Map();
  legacyStudentsSheet.getRow(1).eachCell((cell, column) => {
    legacyHeaders.set(String(cell.value).toLowerCase(), column);
  });
  const teacherUserIdColumn = legacyHeaders.get("teacher user id") || legacyHeaders.get("teacher id");
  assert.ok(teacherUserIdColumn, "student worksheet has a teacher ownership column");
  for (let rowNumber = 2; rowNumber <= legacyStudentsSheet.rowCount; rowNumber += 1) {
    const row = legacyStudentsSheet.getRow(rowNumber);
    if (String(row.getCell(1).value) === addedStudent.student.id) {
      row.getCell(teacherUserIdColumn).value = null;
      break;
    }
  }
  writePortalDatabaseWorkbook(databasePath, legacyRosterWorkbook);
  const legacyRosterResponse = await fetch(`${baseUrl}/api/students`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(legacyRosterResponse.status, 200);
  assert.deepEqual((await legacyRosterResponse.json()).students.map(({ id }) => id), [addedStudent.student.id]);
  for (let rowNumber = 2; rowNumber <= legacyStudentsSheet.rowCount; rowNumber += 1) {
    const row = legacyStudentsSheet.getRow(rowNumber);
    if (String(row.getCell(1).value) === addedStudent.student.id) {
      row.getCell(teacherUserIdColumn).value = registeredTeacherId;
      break;
    }
  }
  writePortalDatabaseWorkbook(databasePath, legacyRosterWorkbook);

  const unauthenticatedGrades = await fetch(
    `${baseUrl}/api/grades?studentId=${encodeURIComponent(addedStudent.student.id)}&term=Term%201`,
  );
  assert.equal(unauthenticatedGrades.status, 401);
  const unauthenticatedGradeLock = await fetch(`${baseUrl}/api/grades/lock?term=Term%201`);
  assert.equal(unauthenticatedGradeLock.status, 401);
  const unauthenticatedSummary = await fetch(`${baseUrl}/api/grades/summary`);
  assert.equal(unauthenticatedSummary.status, 401);
  const unauthenticatedReading = await fetch(`${baseUrl}/api/reading-assessments`);
  assert.equal(unauthenticatedReading.status, 401);
  const unauthenticatedAttendance = await fetch(`${baseUrl}/api/attendance?date=2026-10-14`);
  assert.equal(unauthenticatedAttendance.status, 401);
  const unauthenticatedSF2Export = await fetch(`${baseUrl}/api/attendance/month/export?month=2026-10`);
  assert.equal(unauthenticatedSF2Export.status, 401);
  const unauthenticatedSF2Import = await fetch(`${baseUrl}/api/attendance/month/import?month=2026-10`, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream" },
    body: Buffer.from("Not a workbook"),
  });
  assert.equal(unauthenticatedSF2Import.status, 401);
  const unauthenticatedMonthlyAttendance = await fetch(`${baseUrl}/api/attendance/month?month=2026-10`);
  assert.equal(unauthenticatedMonthlyAttendance.status, 401);

  const emptyReadingResponse = await fetch(`${baseUrl}/api/reading-assessments`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(emptyReadingResponse.status, 200);
  const emptyReading = await emptyReadingResponse.json();
  assert.deepEqual(emptyReading.languages, ["English", "Filipino"]);
  assert.deepEqual(emptyReading.levels, ["Non-reader", "Frustration", "Instructional", "Independent"]);
  assert.equal(emptyReading.attemptNo, 1);
  assert.equal(emptyReading.students[0].levels.English, "");
  const initialRosterResponse = await fetch(`${baseUrl}/api/students`, {
    headers: { Cookie: teacherCookie },
  });
  const initialRoster = (await initialRosterResponse.json()).students;
  assert.deepEqual(
    emptyReading.students.map(({ id, name }) => ({ id, name })),
    initialRoster.map(({ id, name }) => ({ id, name })),
    "reading report learners match the student list",
  );

  const saveReadingResponse = await fetch(`${baseUrl}/api/reading-assessments`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      studentId: addedStudent.student.id,
      period: "BOSY",
      levels: { English: "Independent", Filipino: "Instructional" },
    }),
  });
  assert.equal(saveReadingResponse.status, 200);
  const savedReadingResult = await saveReadingResponse.json();
  assert.equal(savedReadingResult.saved, true);
  assert.equal(savedReadingResult.attemptNo, 1);
  assert.match(savedReadingResult.assessedDates.English, /^\d{4}-\d{2}-\d{2}$/);
  assert.match(savedReadingResult.assessedDates.Filipino, /^\d{4}-\d{2}-\d{2}$/);
  const savedReadingResponse = await fetch(`${baseUrl}/api/reading-assessments`, {
    headers: { Cookie: teacherCookie },
  });
  const savedReading = await savedReadingResponse.json();
  assert.equal(savedReading.students[0].levels.English, "Independent");
  assert.equal(savedReading.students[0].levels.Filipino, "Instructional");
  assert.match(savedReading.students[0].assessedDates.English, /^\d{4}-\d{2}-\d{2}$/);

  const secondReadingAttempt = await fetch(`${baseUrl}/api/reading-assessments`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      studentId: addedStudent.student.id,
      period: "BOSY",
      attemptNo: 2,
      levels: { English: "Instructional", Filipino: "Independent" },
    }),
  });
  assert.equal(secondReadingAttempt.status, 200);
  const secondReadingAttemptResult = await secondReadingAttempt.json();
  assert.equal(secondReadingAttemptResult.attemptNo, 2);
  assert.match(secondReadingAttemptResult.assessedDates.English, /^\d{4}-\d{2}-\d{2}$/);
  const savedSecondAttempt = await fetch(`${baseUrl}/api/reading-assessments?period=BOSY&attemptNo=2`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(savedSecondAttempt.status, 200);
  const secondAttemptReport = await savedSecondAttempt.json();
  assert.equal(secondAttemptReport.students[0].levels.English, "Instructional");
  assert.equal(secondAttemptReport.students[0].levels.Filipino, "Independent");
  const stillSavedFirstAttempt = await fetch(`${baseUrl}/api/reading-assessments?period=BOSY&attemptNo=1`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal((await stillSavedFirstAttempt.json()).students[0].levels.English, "Independent");

  for (const assessment of [
    { period: "MOSY", levels: { English: "Instructional", Filipino: "Frustration" } },
    { period: "EOSY", levels: { English: "Independent", Filipino: "Independent" } },
  ]) {
    const response = await fetch(`${baseUrl}/api/reading-assessments`, {
      method: "POST",
      headers: { "Content-Type": "application/json", Cookie: teacherCookie },
      body: JSON.stringify({
        studentId: addedStudent.student.id,
        ...assessment,
      }),
    });
    assert.equal(response.status, 200);
  }
  for (const assessment of [
    { period: "BOSY", levels: { English: "Independent", Filipino: "Instructional" } },
    { period: "MOSY", levels: { English: "Instructional", Filipino: "Frustration" } },
    { period: "EOSY", levels: { English: "Independent", Filipino: "Independent" } },
  ]) {
    const response = await fetch(`${baseUrl}/api/reading-assessments?period=${assessment.period}`, {
      headers: { Cookie: teacherCookie },
    });
    assert.equal(response.status, 200);
    const report = await response.json();
    assert.equal(report.period, assessment.period);
    assert.equal(report.students[0].levels.English, assessment.levels.English);
    assert.equal(report.students[0].levels.Filipino, assessment.levels.Filipino);
  }
  const invalidReadingPeriod = await fetch(`${baseUrl}/api/reading-assessments?period=MOY`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(invalidReadingPeriod.status, 400);
  const invalidReadingAttempt = await fetch(`${baseUrl}/api/reading-assessments?period=BOSY&attemptNo=0`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(invalidReadingAttempt.status, 400);
  const invalidReadingAttemptSave = await fetch(`${baseUrl}/api/reading-assessments`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      studentId: addedStudent.student.id,
      period: "BOSY",
      attemptNo: 1.5,
      levels: { English: "Independent", Filipino: "Instructional" },
    }),
  });
  assert.equal(invalidReadingAttemptSave.status, 400);
  const missingReadingPeriod = await fetch(`${baseUrl}/api/reading-assessments`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      studentId: addedStudent.student.id,
      levels: { English: "Independent", Filipino: "Instructional" },
    }),
  });
  assert.equal(missingReadingPeriod.status, 400);

  const invalidReadingResponse = await fetch(`${baseUrl}/api/reading-assessments`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      studentId: addedStudent.student.id,
      period: "BOSY",
      levels: { English: "Unknown level", Filipino: "" },
    }),
  });
  assert.equal(invalidReadingResponse.status, 400);
  assert.match((await invalidReadingResponse.json()).error, /valid English reading level/i);

  const gradesUrl = `${baseUrl}/api/grades?studentId=${encodeURIComponent(addedStudent.student.id)}&term=Term%201`;
  const emptyGradesResponse = await fetch(gradesUrl, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(emptyGradesResponse.status, 200);
  const emptyGrades = await emptyGradesResponse.json();
  assert.equal(emptyGrades.student.id, addedStudent.student.id);
  assert.equal(emptyGrades.term, "Term 1");
  assert.equal(emptyGrades.grades.filipino, "");
  assert.equal(emptyGrades.grades.musicAndArts, "");
  const initialGradeLockResponse = await fetch(
    `${baseUrl}/api/grades/lock?studentId=${encodeURIComponent(addedStudent.student.id)}&term=Term%201`,
    {
      headers: { Cookie: teacherCookie },
    },
  );
  assert.equal(initialGradeLockResponse.status, 200);
  assert.deepEqual(await initialGradeLockResponse.json(), {
    studentId: addedStudent.student.id,
    term: "Term 1",
    locked: false,
  });

  const termGrades = {
    filipino: "88",
    english: "91",
    mathematics: "86.5",
    science: "90",
    aralingPanlipunan: "89",
    gmrc: "95",
    eppTle: "87",
    musicAndArts: "92",
    physicalEducationAndHealth: "84",
  };
  const saveGradesResponse = await fetch(`${baseUrl}/api/grades`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      studentId: addedStudent.student.id,
      term: "Term 1",
      grades: termGrades,
    }),
  });
  assert.equal(saveGradesResponse.status, 200);
  const savedGradesResult = await saveGradesResponse.json();
  assert.equal(savedGradesResult.locked, true, "saving grades locks this student's grades");
  assert.deepEqual(savedGradesResult.grades, {
    filipino: 88,
    english: 91,
    mathematics: 86.5,
    science: 90,
    aralingPanlipunan: 89,
    gmrc: 95,
    eppTle: 87,
    musicAndArts: 92,
    physicalEducationAndHealth: 84,
  });
  const lockAfterSaveResponse = await fetch(
    `${baseUrl}/api/grades/lock?studentId=${encodeURIComponent(addedStudent.student.id)}&term=Term%201`,
    { headers: { Cookie: teacherCookie } },
  );
  assert.equal(lockAfterSaveResponse.status, 200);
  assert.equal((await lockAfterSaveResponse.json()).locked, true);

  const lockGradesResponse = await fetch(`${baseUrl}/api/grades/lock`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({ studentId: addedStudent.student.id, term: "Term 1", locked: true }),
  });
  assert.equal(lockGradesResponse.status, 200);
  assert.deepEqual(await lockGradesResponse.json(), {
    studentId: addedStudent.student.id,
    term: "Term 1",
    locked: true,
  });
  const lockedGradeSaveResponse = await fetch(`${baseUrl}/api/grades`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      studentId: addedStudent.student.id,
      term: "Term 1",
      grades: { ...termGrades, filipino: "89" },
    }),
  });
  assert.equal(lockedGradeSaveResponse.status, 409);
  assert.match((await lockedGradeSaveResponse.json()).error, /Term 1 grades for this student are locked/i);
  const unlockGradesResponse = await fetch(`${baseUrl}/api/grades/lock`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({ studentId: addedStudent.student.id, term: "Term 1", locked: false }),
  });
  assert.equal(unlockGradesResponse.status, 200);
  assert.deepEqual(await unlockGradesResponse.json(), {
    studentId: addedStudent.student.id,
    term: "Term 1",
    locked: false,
  });
  const resaveAfterEditResponse = await fetch(`${baseUrl}/api/grades`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      studentId: addedStudent.student.id,
      term: "Term 1",
      grades: termGrades,
    }),
  });
  assert.equal(resaveAfterEditResponse.status, 200);
  assert.equal((await resaveAfterEditResponse.json()).locked, true, "saving edited grades locks them again");

  const termTwoGradesResponse = await fetch(
    `${baseUrl}/api/grades?studentId=${encodeURIComponent(addedStudent.student.id)}&term=Term%202`,
    { headers: { Cookie: teacherCookie } },
  );
  assert.equal(termTwoGradesResponse.status, 200);
  assert.equal((await termTwoGradesResponse.json()).grades.filipino, "");

  const termGradesByTerm = {
    "Term 1": Object.fromEntries(Object.entries(termGrades).map(([subject, grade]) => [subject, Number(grade)])),
    "Term 2": Object.fromEntries(Object.entries(termGrades).map(([subject, grade]) => [subject, Number(grade) - 1])),
    "Term 3": Object.fromEntries(Object.entries(termGrades).map(([subject, grade]) => [subject, Number(grade) - 2])),
  };
  for (const term of ["Term 2", "Term 3"]) {
    const response = await fetch(`${baseUrl}/api/grades`, {
      method: "POST",
      headers: { "Content-Type": "application/json", Cookie: teacherCookie },
      body: JSON.stringify({
        studentId: addedStudent.student.id,
        term,
        grades: termGradesByTerm[term],
      }),
    });
    assert.equal(response.status, 200);
  }

  const gradeSummaryResponse = await fetch(`${baseUrl}/api/grades/summary`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(gradeSummaryResponse.status, 200);
  const gradeSummary = await gradeSummaryResponse.json();
  assert.deepEqual(gradeSummary.terms, ["Term 1", "Term 2", "Term 3"]);
  assert.deepEqual(gradeSummary.students.map((student) => student.id), [addedStudent.student.id]);
  assert.deepEqual(gradeSummary.subjects.map((subject) => subject.key), [
    "filipino",
    "english",
    "mathematics",
    "science",
    "aralingPanlipunan",
    "gmrc",
    "eppTle",
    "mapeh",
    "musicAndArts",
    "physicalEducationAndHealth",
  ]);
  for (const term of gradeSummary.terms) {
    const summaryGrades = gradeSummary.students[0].grades[term];
    for (const [subject, grade] of Object.entries(termGradesByTerm[term])) {
      assert.equal(summaryGrades[subject], grade, `${term} ${subject} appears in the summary`);
    }
    assert.equal(
      summaryGrades.mapeh,
      Math.round(((termGradesByTerm[term].musicAndArts + termGradesByTerm[term].physicalEducationAndHealth) / 2) * 100) / 100,
      `${term} MAPEH average appears in the summary`,
    );
  }

  const invalidGradeResponse = await fetch(`${baseUrl}/api/grades`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      studentId: addedStudent.student.id,
      term: "Term 1",
      grades: { ...termGrades, filipino: 101 },
    }),
  });
  assert.equal(invalidGradeResponse.status, 400);
  assert.match((await invalidGradeResponse.json()).error, /0 to 100/i);

  const invalidTermResponse = await fetch(`${baseUrl}/api/grades`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      studentId: addedStudent.student.id,
      term: "Term 5",
      grades: termGrades,
    }),
  });
  assert.equal(invalidTermResponse.status, 400);
  assert.match((await invalidTermResponse.json()).error, /Term 1 through Term 3/i);

  const studentListResponse = await fetch(`${baseUrl}/api/students`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(studentListResponse.status, 200);
  assert.equal((await studentListResponse.json()).students.length, 1);

  const importedSf1Student = Array(41).fill("");
  importedSf1Student[0] = 123456789013;
  importedSf1Student[2] = "Mendoza, Sam, Cruz";
  importedSf1Student[6] = "M";
  importedSf1Student[7] = "2016-02-03";
  importedSf1Student[11] = "Cebuano";
  importedSf1Student[13] = "Filipino";
  importedSf1Student[14] = "Catholic";
  importedSf1Student[15] = "Unit 1";
  importedSf1Student[17] = "Barangay Central";
  importedSf1Student[20] = "Margosatubig";
  importedSf1Student[22] = "Zamboanga del Sur";
  importedSf1Student[27] = "Santos, Jordan, Cruz";
  importedSf1Student[31] = "Reyes, Ana, Maria";
  importedSf1Student[40] = "Uncle";
  const sf1Sheet = XLSX.utils.aoa_to_sheet([
    [],
    [],
    [],
    [],
    ["LRN", "", "NAME (Last Name, First Name, Middle Name)", "", "", "", "Sex (M/F)", "BIRTH DATE (mm/dd/yyyy)", "", "", "", "MOTHER TONGUE", "", "IP (Ethnic Group)", "RELIGION", "ADDRESS"],
    ["123456789012", "", "Duplicate, Student, Name", "", "", "", "F", "2015-04-21"],
    importedSf1Student,
    ["123456789014", "", "Invalid, Date, Student", "", "", "", "F", "not a date"],
    [16, "", "<=== COMBINED"],
    ["List and Code of Indicators under REMARKS column"],
    ["Transfered Out"],
  ]);
  const sf1Workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(sf1Workbook, sf1Sheet, "SF1");
  const sf1File = XLSX.write(sf1Workbook, { bookType: "biff8", type: "buffer" });
  assert.deepEqual([...sf1File.subarray(0, 8)], [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  const importedResponse = await fetch(
    `${baseUrl}/api/students/import?filename=Renamed-SF1-file.backup`,
    {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream", Cookie: teacherCookie },
      body: sf1File,
    },
  );
  assert.equal(importedResponse.status, 200);
  const importedResult = await importedResponse.json();
  assert.equal(importedResult.imported, 1);
  assert.deepEqual(importedResult.skipped, [
    { row: 6, reason: "LRN is already in the student roster or this SF1 file." },
    { row: 8, reason: "Enter a valid birthdate." },
  ]);
  const importedStudentList = await fetch(`${baseUrl}/api/students`, {
    headers: { Cookie: teacherCookie },
  });
  const importedStudents = (await importedStudentList.json()).students;
  assert.equal(importedStudents.length, 2);
  assert.equal(importedStudents[1].lrn, "123456789013");
  assert.equal(importedStudents[1].firstName, "Sam");
  assert.equal(importedStudents[1].gender, "Male");
  assert.equal(importedStudents[1].birthdate, "2016-02-03");
  assert.equal(importedStudents[1].motherTongue, "Cebuano");
  assert.equal(importedStudents[1].ethnicities, "Filipino");
  assert.equal(importedStudents[1].religion, "Catholic");
  assert.equal(importedStudents[1].currentResidence, "Unit 1, Barangay Central, Margosatubig, Zamboanga del Sur");
  const lockFirstStudentResponse = await fetch(`${baseUrl}/api/grades/lock`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({ studentId: addedStudent.student.id, term: "Term 1", locked: true }),
  });
  assert.equal(lockFirstStudentResponse.status, 200);
  const importedStudentLockResponse = await fetch(
    `${baseUrl}/api/grades/lock?studentId=${encodeURIComponent(importedStudents[1].id)}&term=Term%201`,
    { headers: { Cookie: teacherCookie } },
  );
  assert.equal(importedStudentLockResponse.status, 200);
  assert.deepEqual(await importedStudentLockResponse.json(), {
    studentId: importedStudents[1].id,
    term: "Term 1",
    locked: false,
  });
  const saveUnlockedStudentGradesResponse = await fetch(`${baseUrl}/api/grades`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      studentId: importedStudents[1].id,
      term: "Term 1",
      grades: termGrades,
    }),
  });
  assert.equal(saveUnlockedStudentGradesResponse.status, 200);
  const preserveUnlockedStudentLockResponse = await fetch(`${baseUrl}/api/grades/lock`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({ studentId: importedStudents[1].id, term: "Term 1", locked: false }),
  });
  assert.equal(preserveUnlockedStudentLockResponse.status, 200);
  const legacyLockWorkbook = readPortalDatabaseWorkbook(databasePath);
  const legacyLockSheet = legacyLockWorkbook.getWorksheet("Term Grade Locks");
  const lockHeaders = new Map();
  legacyLockSheet.getRow(1).eachCell((cell, column) => lockHeaders.set(String(cell.value), column));
  const existingStudentLock = legacyLockSheet.getRow(2);
  const legacyLock = legacyLockSheet.getRow(legacyLockSheet.rowCount + 1);
  legacyLock.getCell(lockHeaders.get("Class ID")).value = existingStudentLock.getCell(lockHeaders.get("Class ID")).value;
  legacyLock.getCell(lockHeaders.get("Teacher User ID")).value = existingStudentLock.getCell(lockHeaders.get("Teacher User ID")).value;
  legacyLock.getCell(lockHeaders.get("Term")).value = "Term 1";
  legacyLock.getCell(lockHeaders.get("Student ID")).value = "";
  legacyLock.getCell(lockHeaders.get("Locked")).value = true;
  const legacyTermThreeLock = legacyLockSheet.getRow(legacyLockSheet.rowCount + 1);
  legacyTermThreeLock.getCell(lockHeaders.get("Class ID")).value = legacyLock.getCell(lockHeaders.get("Class ID")).value;
  legacyTermThreeLock.getCell(lockHeaders.get("Teacher User ID")).value = legacyLock.getCell(lockHeaders.get("Teacher User ID")).value;
  legacyTermThreeLock.getCell(lockHeaders.get("Term")).value = "Term 3";
  legacyTermThreeLock.getCell(lockHeaders.get("Student ID")).value = "";
  legacyTermThreeLock.getCell(lockHeaders.get("Locked")).value = false;
  legacyTermThreeLock.getCell(lockHeaders.get("Updated")).value = new Date();
  const migratedLockTimestamp = new Date(Date.now() - 60_000);
  for (const studentId of [addedStudent.student.id, importedStudents[1].id]) {
    const migratedLock = legacyLockSheet.getRow(legacyLockSheet.rowCount + 1);
    migratedLock.getCell(lockHeaders.get("Class ID")).value = legacyTermThreeLock.getCell(lockHeaders.get("Class ID")).value;
    migratedLock.getCell(lockHeaders.get("Teacher User ID")).value = legacyTermThreeLock.getCell(lockHeaders.get("Teacher User ID")).value;
    migratedLock.getCell(lockHeaders.get("Term")).value = "Term 3";
    migratedLock.getCell(lockHeaders.get("Student ID")).value = studentId;
    migratedLock.getCell(lockHeaders.get("Locked")).value = true;
    migratedLock.getCell(lockHeaders.get("Updated")).value = migratedLockTimestamp;
  }
  writePortalDatabaseWorkbook(databasePath, legacyLockWorkbook);
  const migratedUnlockedStudentLockResponse = await fetch(
    `${baseUrl}/api/grades/lock?studentId=${encodeURIComponent(importedStudents[1].id)}&term=Term%201`,
    { headers: { Cookie: teacherCookie } },
  );
  assert.equal(migratedUnlockedStudentLockResponse.status, 200);
  assert.equal((await migratedUnlockedStudentLockResponse.json()).locked, false);
  const migratedLegacyLockWorkbook = readPortalDatabaseWorkbook(databasePath);
  const migratedLegacyLockSheet = migratedLegacyLockWorkbook.getWorksheet("Term Grade Locks");
  const migratedLegacyLock = migratedLegacyLockSheet.getRow(legacyLock.number);
  assert.equal(migratedLegacyLock.getCell(lockHeaders.get("Locked")).value, false);
  const migratedTermThreeLockResponse = await fetch(
    `${baseUrl}/api/grades/lock?studentId=${encodeURIComponent(importedStudents[1].id)}&term=Term%203`,
    { headers: { Cookie: teacherCookie } },
  );
  assert.equal(migratedTermThreeLockResponse.status, 200);
  assert.equal((await migratedTermThreeLockResponse.json()).locked, false);
  const originalStudentTermThreeLockResponse = await fetch(
    `${baseUrl}/api/grades/lock?studentId=${encodeURIComponent(addedStudent.student.id)}&term=Term%203`,
    { headers: { Cookie: teacherCookie } },
  );
  assert.equal(originalStudentTermThreeLockResponse.status, 200);
  assert.equal((await originalStudentTermThreeLockResponse.json()).locked, false);
  const persistedMigratedLockWorkbook = readPortalDatabaseWorkbook(databasePath);
  const persistedMigratedLockSheet = persistedMigratedLockWorkbook.getWorksheet("Term Grade Locks");
  for (let rowNumber = 2; rowNumber <= persistedMigratedLockSheet.rowCount; rowNumber += 1) {
    const row = persistedMigratedLockSheet.getRow(rowNumber);
    if (String(row.getCell(lockHeaders.get("Term")).value) === "Term 3"
      && String(row.getCell(lockHeaders.get("Student ID")).value)) {
      assert.equal(row.getCell(lockHeaders.get("Locked")).value, false);
    }
  }
  const lockedStudentGradesResponse = await fetch(`${baseUrl}/api/grades`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      studentId: addedStudent.student.id,
      term: "Term 1",
      grades: termGrades,
    }),
  });
  assert.equal(lockedStudentGradesResponse.status, 409);

  const updatedReadingResponse = await fetch(`${baseUrl}/api/reading-assessments?period=BOSY`, {
    headers: { Cookie: teacherCookie },
  });
  const updatedReading = await updatedReadingResponse.json();
  assert.deepEqual(
    updatedReading.students.map(({ id, name }) => ({ id, name })),
    importedStudents.map(({ id, name }) => ({ id, name })),
    "reading report reflects newly imported learners from the student list",
  );
  assert.equal(importedStudents[1].father, "Santos, Jordan, Cruz");
  assert.equal(importedStudents[1].mothersMaidenName, "Reyes, Ana, Maria");
  assert.equal(importedStudents[1].guardianRelationship, "Uncle");

  const renamedXlsxStudent = Array(8).fill("");
  renamedXlsxStudent[0] = "123456789015";
  renamedXlsxStudent[2] = "Garcia, Jamie, Lee";
  renamedXlsxStudent[6] = "F";
  renamedXlsxStudent[7] = "2015-09-17";
  const xlsxSheet = XLSX.utils.aoa_to_sheet([
    [],
    [],
    [],
    [],
    ["LRN", "", "NAME (Last Name, First Name, Middle Name)", "", "", "", "Sex (M/F)", "BIRTH DATE (mm/dd/yyyy)"],
    renamedXlsxStudent,
  ]);
  const xlsxWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(xlsxWorkbook, xlsxSheet, "SF1 Register");
  const renamedXlsxFile = XLSX.write(xlsxWorkbook, { bookType: "xlsx", type: "buffer" });
  assert.deepEqual([...renamedXlsxFile.subarray(0, 4)], [0x50, 0x4b, 0x03, 0x04]);
  const renamedXlsxImport = await fetch(`${baseUrl}/api/students/import`, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream", Cookie: teacherCookie },
    body: renamedXlsxFile,
  });
  assert.equal(renamedXlsxImport.status, 200);
  assert.deepEqual(await renamedXlsxImport.json(), { imported: 1, skipped: [] });
  const finalStudentList = await fetch(`${baseUrl}/api/students`, {
    headers: { Cookie: teacherCookie },
  });
  const finalStudents = (await finalStudentList.json()).students;
  assert.equal(finalStudents.length, 3);
  assert.equal(finalStudents[2].lrn, "123456789015");
  assert.equal(finalStudents[2].firstName, "Jamie");
  const attendanceDate = new Date().toISOString().slice(0, 10);
  const attendanceResponse = await fetch(`${baseUrl}/api/attendance?date=${attendanceDate}`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(attendanceResponse.status, 200);
  const emptyAttendance = await attendanceResponse.json();
  assert.equal(emptyAttendance.students.length, 3);
  assert.ok(emptyAttendance.students.every(({ status }) => status === ""));
  const savedAttendanceResponse = await fetch(`${baseUrl}/api/attendance`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      date: attendanceDate,
      records: [
        { studentId: addedStudent.student.id, status: "Present" },
        { studentId: importedStudents[1].id, status: "Absent" },
        { studentId: finalStudents[2].id, status: "Tardy" },
      ],
    }),
  });
  assert.equal(savedAttendanceResponse.status, 200);
  assert.deepEqual(await savedAttendanceResponse.json(), { saved: true, date: attendanceDate });
  const savedAttendance = await fetch(`${baseUrl}/api/attendance?date=${attendanceDate}`, {
    headers: { Cookie: teacherCookie },
  });
  assert.deepEqual((await savedAttendance.json()).students.map(({ id, status }) => ({ id, status })), [
    { id: addedStudent.student.id, status: "Present" },
    { id: importedStudents[1].id, status: "Absent" },
    { id: finalStudents[2].id, status: "Tardy" },
  ]);
  const attendanceMonth = `${Number(attendanceDate.slice(0, 4)) - 1}-01`;
  const emptyMonthResponse = await fetch(`${baseUrl}/api/attendance/month?month=${attendanceMonth}`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(emptyMonthResponse.status, 200);
  const emptyMonth = await emptyMonthResponse.json();
  const attendanceDates = attendanceWeekdays(attendanceMonth);
  assert.deepEqual(emptyMonth.dates, attendanceDates);
  assert.equal(emptyMonth.dates.every((date) => {
    const weekday = new Date(`${date}T00:00:00.000Z`).getUTCDay();
    return weekday !== 0 && weekday !== 6;
  }), true);
  assert.deepEqual(emptyMonth.students.map(({ id, attendance }) => ({
    id,
    days: attendance.length,
    recorded: attendance.filter(Boolean).length,
  })), [
    { id: addedStudent.student.id, days: attendanceDates.length, recorded: 0 },
    { id: importedStudents[1].id, days: attendanceDates.length, recorded: 0 },
    { id: finalStudents[2].id, days: attendanceDates.length, recorded: 0 },
  ]);
  const monthlyRecords = emptyMonth.students.map(({ id }) => ({
    studentId: id,
    attendance: Array(attendanceDates.length).fill(""),
    reasons: Array(attendanceDates.length).fill(""),
    remarks: "",
  }));
  monthlyRecords[0].attendance[0] = "Present";
  monthlyRecords[0].attendance[attendanceDates.length - 1] = "Absent";
  monthlyRecords[0].reasons[attendanceDates.length - 1] = "Medical appointment";
  monthlyRecords[0].remarks = "Transfer information";
  const savedMonthResponse = await fetch(`${baseUrl}/api/attendance/month`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({ month: attendanceMonth, records: monthlyRecords }),
  });
  assert.equal(savedMonthResponse.status, 200);
  assert.deepEqual(await savedMonthResponse.json(), { saved: true, month: attendanceMonth });
  monthlyRecords[0].attendance[0] = "Tardy";
  monthlyRecords[0].reasons[attendanceDates.length - 1] = "Flu";
  monthlyRecords[0].remarks = "Follow-up required";
  const updatedMonthResponse = await fetch(`${baseUrl}/api/attendance/month`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({ month: attendanceMonth, records: monthlyRecords }),
  });
  assert.equal(updatedMonthResponse.status, 200);
  const updatedMonth = await fetch(`${baseUrl}/api/attendance/month?month=${attendanceMonth}`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(updatedMonth.status, 200);
  const updatedMonthResult = await updatedMonth.json();
  assert.equal(updatedMonthResult.students[0].attendance[0], "Tardy");
  assert.equal(updatedMonthResult.students[0].attendance[attendanceDates.length - 1], "Absent");
  assert.equal(updatedMonthResult.students[0].reasons[attendanceDates.length - 1], "Flu");
  assert.equal(updatedMonthResult.students[1].attendance.every((status) => status === ""), true);
  assert.equal(updatedMonthResult.students[0].remarks, "Follow-up required");
  assert.equal(updatedMonthResult.students[0].gender, "Female");
  const sf2WorkbookBuffer = await createSF2Upload(attendanceMonth, [
    {
      name: addedStudent.student.name,
      gender: addedStudent.student.gender,
      marks: { 1: "x", 2: "h" },
      remarks: "Medical appointment",
    },
    {
      name: importedStudents[1].name,
      gender: importedStudents[1].gender,
      marks: { 1: "t", 2: "e" },
      remarks: "Early dismissal",
    },
  ]);
  const importedSF2Response = await fetch(
    `${baseUrl}/api/attendance/month/import?month=${attendanceMonth}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream", Cookie: teacherCookie },
      body: sf2WorkbookBuffer,
    },
  );
  assert.equal(importedSF2Response.status, 200);
  const importedSF2 = await importedSF2Response.json();
  const importedSchoolDays = attendanceDates.map((date) => Number(date.slice(-2)));
  assert.deepEqual(importedSF2, {
    imported: 2,
    importedDays: importedSchoolDays.length * 2,
    month: attendanceMonth,
  });
  const importedSF2MonthResponse = await fetch(
    `${baseUrl}/api/attendance/month?month=${attendanceMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  const importedSF2Month = await importedSF2MonthResponse.json();
  assert.deepEqual(importedSF2Month.students[0].attendance.slice(0, 3), ["Absent", "Half-day", "Present"]);
  assert.equal(importedSF2Month.students[0].reasons[0], "Medical appointment");
  assert.equal(importedSF2Month.students[0].remarks, "Medical appointment");
  assert.deepEqual(importedSF2Month.students[1].attendance.slice(0, 2), ["Tardy", "Excused"]);
  assert.equal(importedSF2Month.students[1].reasons[0], "");
  assert.equal(importedSF2Month.students[1].remarks, "Early dismissal");
  assert.equal(importedSF2Month.students[2].attendance.every((status) => status === ""), true);
  const attendanceSummaryResponse = await fetch(`${baseUrl}/api/attendance/summary`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(attendanceSummaryResponse.status, 200);
  const attendanceSummary = await attendanceSummaryResponse.json();
  const summaryDailyStatuses = new Map([
    [addedStudent.student.id, "Present"],
    [importedStudents[1].id, "Absent"],
    [finalStudents[2].id, "Tardy"],
  ]);
  assert.deepEqual(attendanceSummary.months, [...new Set([
    attendanceMonth,
    attendanceDate.slice(0, 7),
  ])].sort());
  assert.deepEqual(attendanceSummary.students.map((student, index) => {
    const monthStudent = importedSF2Month.students[index];
    const monthPresent = monthStudent.attendance.reduce((total, status) =>
      total + (status === "Present" || status === "Tardy" ? 1 : status === "Half-day" ? 0.5 : 0), 0);
    const monthAbsent = monthStudent.attendance.filter((status) => status === "Absent").length;
    const dailyStatus = summaryDailyStatuses.get(student.id);
    const dailyPresent = dailyStatus === "Present" || dailyStatus === "Tardy" ? 1 : 0;
    const dailyAbsent = dailyStatus === "Absent" ? 1 : 0;
    return {
      id: student.id,
      present: monthPresent + (dailyStatus === "Present" || dailyStatus === "Tardy" ? 1 : 0),
      absent: monthAbsent + (dailyStatus === "Absent" ? 1 : 0),
      months: {
        [attendanceMonth]: { present: monthPresent, absent: monthAbsent },
        [attendanceDate.slice(0, 7)]: { present: dailyPresent, absent: dailyAbsent },
      },
    };
  }), attendanceSummary.students.map(({ id, present, absent, months }) => ({ id, present, absent, months })));
  assert.match(attendanceSummary.firstDate, /^\d{4}-\d{2}-\d{2}$/);
  assert.equal(attendanceSummary.lastDate, attendanceDate);
  const sf2ExportRecords = importedSF2Month.students.map((student, index) => ({
    studentId: student.id,
    attendance: student.attendance,
    reasons: student.reasons,
    remarks: index === 0 ? "Monthly follow-up" : student.remarks,
  }));
  const sf2ExportNotesResponse = await fetch(`${baseUrl}/api/attendance/month`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({ month: attendanceMonth, records: sf2ExportRecords }),
  });
  assert.equal(sf2ExportNotesResponse.status, 200);

  const sf2ExportResponse = await fetch(
    `${baseUrl}/api/attendance/month/export?month=${attendanceMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  assert.equal(sf2ExportResponse.status, 200, await sf2ExportResponse.clone().text());
  assert.equal(
    sf2ExportResponse.headers.get("content-disposition"),
    `attachment; filename="SF2_January_${attendanceMonth.slice(0, 4)}.xlsx"`,
  );
  const exportedSF2Workbook = new ExcelJS.Workbook();
  const exportedSF2Buffer = Buffer.from(await sf2ExportResponse.arrayBuffer());
  await exportedSF2Workbook.xlsx.load(exportedSF2Buffer);
  const monthNames = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  const exportedMonthName = `${monthNames[Number(attendanceMonth.slice(5, 7)) - 1]} ${attendanceMonth.slice(0, 4)}`;
  const exportedSF2Sheet = exportedSF2Workbook.getWorksheet(exportedMonthName);
  assert.ok(exportedSF2Sheet, "exported workbook has the selected month worksheet");
  assert.equal(exportedSF2Sheet.getCell("AA6").value, "JANUARY");
  assert.equal(exportedSF2Sheet.getCell("AG117").value, "JANUARY");
  assert.equal(exportedSF2Sheet.getCell("AG118").value, "JANUARY");
  assert.equal(exportedSF2Sheet.getCell("N6").value, "2026-2027");
  assert.equal(exportedSF2Sheet.getCell("G6").value, "123456");
  assert.equal(exportedSF2Sheet.getCell("G7").value, "Sample Elementary School");
  assert.equal(exportedSF2Sheet.getCell("AA7").value, "Grade 5");
  assert.equal(exportedSF2Sheet.getCell("AF7").value, "Mabini");
  assert.equal(exportedSF2Sheet.getCell("AE140").value, teacherCredentials.name);
  assert.equal(exportedSF2Sheet.getCell("AE145").value, "Alex Principal");
  const alternateMonth = `${attendanceMonth.slice(0, 4)}-02`;
  const alternateExportResponse = await fetch(
    `${baseUrl}/api/attendance/month/export?month=${alternateMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  assert.equal(alternateExportResponse.status, 200, await alternateExportResponse.clone().text());
  assert.equal(
    alternateExportResponse.headers.get("content-disposition"),
    `attachment; filename="SF2_February_${alternateMonth.slice(0, 4)}.xlsx"`,
  );
  const alternateExportWorkbook = new ExcelJS.Workbook();
  await alternateExportWorkbook.xlsx.load(Buffer.from(await alternateExportResponse.arrayBuffer()));
  const alternateExportSheet = alternateExportWorkbook.getWorksheet(`FEB ${alternateMonth.slice(0, 4)}`);
  assert.ok(alternateExportSheet, "export uses the requested alternate month worksheet");
  assert.equal(alternateExportSheet.getCell("AA6").value, "FEBRUARY");
  assert.equal(alternateExportSheet.getCell("AG117").value, "FEBRUARY");
  assert.equal(alternateExportSheet.getCell("AG118").value, "FEBRUARY");
  const alternateSchoolDays = Array.from({ length: new Date(
    Date.UTC(Number(alternateMonth.slice(0, 4)), 2, 0),
  ).getUTCDate() }, (_, index) => index + 1).filter((day) => {
    const weekday = new Date(
      `${alternateMonth}-${String(day).padStart(2, "0")}T00:00:00.000Z`,
    ).getUTCDay();
    return weekday !== 0 && weekday !== 6;
  });
  assert.deepEqual(
    alternateSchoolDays.map((_, index) => alternateExportSheet.getCell(10, 7 + index).value),
    alternateSchoolDays,
    "alternate selected month dates replace the template's example dates",
  );
  assert.equal(alternateExportSheet.getCell("AQ9").value, alternateSchoolDays.length);
  assert.equal(alternateExportSheet.getCell("AJ117").value.result, alternateSchoolDays.length);
  assert.equal(alternateExportSheet.getCell("AJ118").value.result, alternateSchoolDays.length);
  assert.equal(alternateExportSheet.getCell("AK119").value, finalStudents.filter(({ gender }) => gender === "Male").length);
  assert.equal(alternateExportSheet.getCell("AL119").value, finalStudents.filter(({ gender }) => gender === "Female").length);
  assert.equal(alternateExportSheet.getCell("AM119").value.result, finalStudents.length);
  assert.equal(
    alternateExportSheet.getCell("AK123").value.formula,
    'COUNTA($B$13:$B$62)-COUNTIF($AH$13:$AH$62,"NLS*")-COUNTIF($AH$13:$AH$62,"Transferred Out*")+COUNTIF($AH$13:$AH$62,"Transferred In*")',
  );
  assert.equal(alternateExportSheet.getCell("AK123").value.result, finalStudents.filter(({ gender }) => gender === "Male").length);
  assert.equal(alternateExportSheet.getCell("AM123").value.result, finalStudents.length);
  assert.deepEqual(
    importedSchoolDays.map((_, index) => exportedSF2Sheet.getCell(10, 7 + index).value),
    importedSchoolDays,
    "exported date headings follow the selected month's school days in chronological order",
  );
  assert.deepEqual(
    importedSchoolDays.map((day, index) => exportedSF2Sheet.getCell(11, 7 + index).value),
    importedSchoolDays.map((day) => ["M", "T", "W", "TH", "F"][
      (new Date(`${attendanceMonth}-${String(day).padStart(2, "0")}T00:00:00.000Z`).getUTCDay() + 6) % 7
    ]),
    "each exported weekday heading matches its selected-month date",
  );
  assert.equal(typeof exportedSF2Sheet.getCell(63, 2).value, "string", "the male section total row is preserved");
  const exportedRowsByName = new Map(sf2StudentRows
    .map((rowNumber) => [exportedSF2Sheet.getCell(rowNumber, 2).value, rowNumber])
    .filter(([name]) => name));
  assert.deepEqual([...exportedRowsByName.keys()].sort(), finalStudents.map(({ name }) => name).sort());
  const maleRosterCount = finalStudents.filter(({ gender }) => gender === "Male").length;
  const femaleRosterCount = finalStudents.filter(({ gender }) => gender === "Female").length;
  assert.equal(exportedSF2Sheet.getCell("AK119").value, maleRosterCount);
  assert.equal(exportedSF2Sheet.getCell("AL119").value, femaleRosterCount);
  assert.equal(exportedSF2Sheet.getCell("AM119").value.formula, "AK119+AL119");
  assert.equal(exportedSF2Sheet.getCell("AM119").value.result, finalStudents.length);
  assert.equal(
    exportedSF2Sheet.getCell("AK123").value.formula,
    'COUNTA($B$13:$B$62)-COUNTIF($AH$13:$AH$62,"NLS*")-COUNTIF($AH$13:$AH$62,"Transferred Out*")+COUNTIF($AH$13:$AH$62,"Transferred In*")',
  );
  assert.equal(exportedSF2Sheet.getCell("AK123").value.result, maleRosterCount);
  assert.equal(exportedSF2Sheet.getCell("AL123").value.result, femaleRosterCount);
  assert.equal(exportedSF2Sheet.getCell("AM123").value.formula, "AK123+AL123");
  assert.equal(exportedSF2Sheet.getCell("AM123").value.result, finalStudents.length);
  assert.equal(exportedSF2Sheet.getCell("AK127").value.formula, "IFERROR(AG63/AJ117,0)");
  assert.equal(exportedSF2Sheet.getCell("AK127").value.result, 1);
  assert.equal(
    exportedSF2Sheet.getCell("AL127").value.result,
    (femaleRosterCount * importedSchoolDays.length - 1.5) / importedSchoolDays.length,
  );
  assert.equal(
    exportedSF2Sheet.getCell("AM127").value.result,
    (finalStudents.length * importedSchoolDays.length - 1.5) / importedSchoolDays.length,
  );
  assert.equal(exportedSF2Sheet.getCell("AG63").value.result, maleRosterCount * importedSchoolDays.length);
  assert.equal(exportedSF2Sheet.getCell("AG114").value.result,
    femaleRosterCount * importedSchoolDays.length - 1.5);
  assert.equal(exportedSF2Sheet.getCell("AG115").value.result,
    finalStudents.length * importedSchoolDays.length - 1.5);
  for (const rowNumber of sf2StudentRows) {
    assert.equal(
      exportedSF2Sheet.getRow(rowNumber).hidden,
      !exportedRowsByName.has(exportedSF2Sheet.getCell(rowNumber, 2).value),
      `only populated roster rows are visible (row ${rowNumber})`,
    );
  }
  const addedStudentExportRow = exportedRowsByName.get(addedStudent.student.name);
  const importedStudentExportRow = exportedRowsByName.get(importedStudents[1].name);
  assert.ok(addedStudentExportRow >= 64, "female learners are placed in the female section");
  assert.ok(importedStudentExportRow >= 13 && importedStudentExportRow <= 62, "male learners are placed in the male section");
  const exportedFirstDayColumn = Array.from({ length: 25 }, (_, index) => index + 7)
    .find((column) => exportedSF2Sheet.getCell(10, column).value === importedSchoolDays[0]);
  const exportedSecondDayColumn = Array.from({ length: 25 }, (_, index) => index + 7)
    .find((column) => exportedSF2Sheet.getCell(10, column).value === importedSchoolDays[1]);
  assert.equal(exportedSF2Sheet.getCell(addedStudentExportRow, exportedFirstDayColumn).value, "x");
  assert.equal(exportedSF2Sheet.getCell(addedStudentExportRow, exportedSecondDayColumn).value, "h");
  assert.equal(exportedSF2Sheet.getCell(importedStudentExportRow, exportedFirstDayColumn).value, "t");
  assert.equal(exportedSF2Sheet.getCell(importedStudentExportRow, exportedSecondDayColumn).value, "e");
  const dateSpecificRemarks = exportedSF2Sheet.getCell(addedStudentExportRow, 34).value;
  assert.match(dateSpecificRemarks, /\[Grade Portal absence reasons\]/);
  assert.match(dateSpecificRemarks, new RegExp(`${attendanceMonth}-01: Medical appointment`));
  const roundTripSF2Import = await fetch(
    `${baseUrl}/api/attendance/month/import?month=${attendanceMonth}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream", Cookie: teacherCookie },
      body: exportedSF2Buffer,
    },
  );
  assert.equal(roundTripSF2Import.status, 200);
  const roundTripAttendance = await fetch(
    `${baseUrl}/api/attendance/month?month=${attendanceMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  const roundTripAttendanceResult = await roundTripAttendance.json();
  assert.equal(roundTripAttendanceResult.students[0].reasons[0], "Medical appointment");
  assert.equal(roundTripAttendanceResult.students[0].remarks, "Monthly follow-up");
  const unknownLearnerSF2 = await createSF2Upload(attendanceMonth, [
    { name: addedStudent.student.name, gender: addedStudent.student.gender, marks: { 1: "x" }, remarks: "Medical appointment" },
    { name: importedStudents[1].name, gender: importedStudents[1].gender, marks: { 1: "t" }, remarks: "Early dismissal" },
    { name: "Unknown, Learner", marks: {}, remarks: "" },
  ]);
  const unknownLearnerImport = await fetch(
    `${baseUrl}/api/attendance/month/import?month=${attendanceMonth}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream", Cookie: teacherCookie },
      body: unknownLearnerSF2,
    },
  );
  assert.equal(unknownLearnerImport.status, 400);
  assert.match((await unknownLearnerImport.json()).error, /No unique learner.*No attendance was changed/);
  const absentWithoutSF2Remark = await createSF2Upload(attendanceMonth, [
    { name: addedStudent.student.name, gender: addedStudent.student.gender, marks: { 1: "x" }, remarks: "" },
  ]);
  const absentWithoutSF2RemarkImport = await fetch(
    `${baseUrl}/api/attendance/month/import?month=${attendanceMonth}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream", Cookie: teacherCookie },
      body: absentWithoutSF2Remark,
    },
  );
  assert.equal(absentWithoutSF2RemarkImport.status, 400);
  assert.match((await absentWithoutSF2RemarkImport.json()).error, /no SF2 remarks.*absence reason/);
  const unchangedAfterRejectedImport = await fetch(
    `${baseUrl}/api/attendance/month?month=${attendanceMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  assert.equal((await unchangedAfterRejectedImport.json()).students[0].attendance[0], "Absent");
  const unsupportedSF2Mark = await createSF2Upload(attendanceMonth, [
    { name: addedStudent.student.name, gender: addedStudent.student.gender, marks: { 1: "?" }, remarks: "" },
  ]);
  const unsupportedSF2Import = await fetch(
    `${baseUrl}/api/attendance/month/import?month=${attendanceMonth}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream", Cookie: teacherCookie },
      body: unsupportedSF2Mark,
    },
  );
  assert.equal(unsupportedSF2Import.status, 400);
  assert.match((await unsupportedSF2Import.json()).error, /Unsupported SF2 mark/);
  const attendanceBeforeNoClass = await fetch(
    `${baseUrl}/api/attendance/month?month=${attendanceMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  const attendanceBeforeNoClassData = await attendanceBeforeNoClass.json();
  const noClassDateIndex = attendanceBeforeNoClassData.dates.findIndex((date) => {
    const weekday = new Date(`${date}T00:00:00.000Z`).getUTCDay();
    return weekday !== 0 && weekday !== 6;
  });
  const noClassDate = attendanceBeforeNoClassData.dates[noClassDateIndex];
  const noClassRecords = attendanceBeforeNoClassData.students.map((student) => ({
    studentId: student.id,
    attendance: student.attendance.map((status, index) => index === noClassDateIndex ? "" : status),
    reasons: student.reasons.map((reason, index) => index === noClassDateIndex ? "" : reason),
    remarks: student.remarks,
  }));
  const savedNoClassDate = await fetch(`${baseUrl}/api/attendance/month`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      month: attendanceMonth,
      records: noClassRecords,
      noClassDates: [noClassDate],
    }),
  });
  assert.equal(savedNoClassDate.status, 200, await savedNoClassDate.clone().text());
  const savedNoClassMonth = await fetch(
    `${baseUrl}/api/attendance/month?month=${attendanceMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  const savedNoClassData = await savedNoClassMonth.json();
  assert.deepEqual(savedNoClassData.noClassDates, [noClassDate]);
  assert.equal(savedNoClassData.students.every((student) =>
    student.attendance[noClassDateIndex] === "" && student.reasons[noClassDateIndex] === ""), true);
  const summaryAfterNoClass = await fetch(`${baseUrl}/api/attendance/summary`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(summaryAfterNoClass.status, 200);
  assert.deepEqual((await summaryAfterNoClass.json()).students.map(({ id, present, absent }) =>
    ({ id, present, absent })), savedNoClassData.students.map((student) => {
    const present = student.attendance.reduce((total, status) =>
      total + (status === "Present" || status === "Tardy" ? 1 : status === "Half-day" ? 0.5 : 0), 0);
    const absent = student.attendance.filter((status) => status === "Absent").length;
    const dailyStatus = summaryDailyStatuses.get(student.id);
    return {
      id: student.id,
      present: present + (dailyStatus === "Present" || dailyStatus === "Tardy" ? 1 : 0),
      absent: absent + (dailyStatus === "Absent" ? 1 : 0),
    };
  }));
  const noClassExport = await fetch(
    `${baseUrl}/api/attendance/month/export?month=${attendanceMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  assert.equal(noClassExport.status, 200, await noClassExport.clone().text());
  const noClassExportBuffer = Buffer.from(await noClassExport.arrayBuffer());
  const noClassExportWorkbook = new ExcelJS.Workbook();
  await noClassExportWorkbook.xlsx.load(noClassExportBuffer);
  const noClassExportSheet = noClassExportWorkbook.getWorksheet(exportedMonthName);
  assert.equal(noClassExportSheet.getCell("AA6").value, "JANUARY");
  assert.equal(noClassExportSheet.getCell("AG117").value, "JANUARY");
  assert.equal(noClassExportSheet.getCell("AG118").value, "JANUARY");
  const expectedClassDayCount = importedSchoolDays.length - 1;
  assert.equal(noClassExportSheet.getCell("AQ9").value, expectedClassDayCount);
  assert.equal(noClassExportSheet.getCell("AJ117").value.formula, "AQ9");
  assert.equal(noClassExportSheet.getCell("AJ117").value.result, expectedClassDayCount);
  assert.equal(noClassExportSheet.getCell("AJ118").value.result, expectedClassDayCount);
  assert.equal(noClassExportSheet.getCell("AK119").value, maleRosterCount);
  assert.equal(noClassExportSheet.getCell("AL119").value, femaleRosterCount);
  assert.equal(noClassExportSheet.getCell("AM119").value.result, finalStudents.length);
  assert.equal(noClassExportSheet.getCell("AK123").value.result, maleRosterCount);
  assert.equal(noClassExportSheet.getCell("AL123").value.result, femaleRosterCount);
  assert.equal(noClassExportSheet.getCell("AM123").value.result, finalStudents.length);
  const noClassDateColumn = Array.from({ length: 25 }, (_, index) => index + 7)
    .find((column) => noClassExportSheet.getCell(10, column).value === Number(noClassDate.slice(-2)));
  assert.notEqual(noClassDateColumn, undefined, "the no-class date appears in the exported date headings");
  assert.equal(
    noClassExportSheet.getCell(115, noClassDateColumn).value.formula,
    `${noClassExportSheet.getColumn(noClassDateColumn).letter}63+${noClassExportSheet.getColumn(noClassDateColumn).letter}114`,
  );
  for (const rowNumber of exportedRowsByName.values()) {
    const noClassCell = noClassExportSheet.getCell(rowNumber, noClassDateColumn);
    assert.equal(noClassCell.value, "NC", `learner row ${rowNumber} displays NC`);
    assert.equal(noClassCell.fill.fgColor.argb, "FFE6E6E6", `learner row ${rowNumber} has a light-gray fill`);
  }
  assert.notEqual(noClassExportSheet.getCell(10, noClassDateColumn).fill.fgColor?.argb, "FFE6E6E6");
  assert.notEqual(noClassExportSheet.getCell(11, noClassDateColumn).fill.fgColor?.argb, "FFE6E6E6");
  const classDayColumn = noClassDateColumn === 7 ? 8 : 7;
  for (const rowNumber of exportedRowsByName.values()) {
    assert.notEqual(
      noClassExportSheet.getCell(rowNumber, classDayColumn).fill.fgColor?.argb,
      "FFE6E6E6",
      `class-day cell in learner row ${rowNumber} remains unshaded`,
    );
  }
  assert.equal(noClassExportSheet.getCell("AK127").value.result, maleRosterCount);
  assert.equal(noClassExportSheet.getCell("AL127").value.result,
    (femaleRosterCount * expectedClassDayCount - 0.5) / expectedClassDayCount);
  assert.equal(noClassExportSheet.getCell("AM127").value.result,
    (finalStudents.length * expectedClassDayCount - 0.5) / expectedClassDayCount);
  const importedNoClassExport = await fetch(
    `${baseUrl}/api/attendance/month/import?month=${attendanceMonth}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream", Cookie: teacherCookie },
      body: noClassExportBuffer,
    },
  );
  assert.equal(importedNoClassExport.status, 200, await importedNoClassExport.clone().text());
  assert.deepEqual(await importedNoClassExport.json(), {
    imported: finalStudents.length,
    importedDays: expectedClassDayCount * finalStudents.length,
    month: attendanceMonth,
  });
  const noClassAfterExportImport = await fetch(
    `${baseUrl}/api/attendance/month?month=${attendanceMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  const noClassAfterExportImportData = await noClassAfterExportImport.json();
  assert.deepEqual(noClassAfterExportImportData.noClassDates, [noClassDate]);
  assert.equal(noClassAfterExportImportData.students.every((student) =>
    student.attendance[noClassDateIndex] === ""), true);
  assert.ok(noClassAfterExportImportData.students.length >= 3, "status summary test has three learners");
  const monthlyStatusByStudent = new Map(noClassAfterExportImportData.students.slice(0, 3).map((student, index) => [
    student.id,
    ["NLS", "Transferred Out", "Transferred In"][index],
  ]));
  const monthlyStatusRecords = noClassAfterExportImportData.students.map((student) => {
    const monthlyStatus = monthlyStatusByStudent.get(student.id) || student.remarks;
    const statusDate = monthlyStatusByStudent.has(student.id) ? `${attendanceMonth}-15` : student.statusDate || "";
    const beforeTransferIn = (date) => monthlyStatus === "Transferred In" && date < statusDate;
    return {
      studentId: student.id,
      attendance: student.attendance.map((status, index) => beforeTransferIn(noClassAfterExportImportData.dates[index]) ? "" : status),
      reasons: student.reasons.map((reason, index) => beforeTransferIn(noClassAfterExportImportData.dates[index]) ? "" : reason),
      remarks: monthlyStatus,
      statusDate,
    };
  });
  const saveMonthlyStatuses = await fetch(`${baseUrl}/api/attendance/month`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      month: attendanceMonth,
      records: monthlyStatusRecords,
      noClassDates: noClassAfterExportImportData.noClassDates,
    }),
  });
  assert.equal(saveMonthlyStatuses.status, 200, await saveMonthlyStatuses.clone().text());
  const monthWithStatuses = await fetch(
    `${baseUrl}/api/attendance/month?month=${attendanceMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  const monthWithStatusesData = await monthWithStatuses.json();
  const transferInStudentId = [...monthlyStatusByStudent].find(([, status]) => status === "Transferred In")[0];
  const transferInStudent = monthWithStatusesData.students.find((student) => student.id === transferInStudentId);
  assert.deepEqual(
    transferInStudent.lockedDates,
    monthWithStatusesData.dates.filter((date) => date < `${attendanceMonth}-15`),
    "Transferred In's effective date excludes earlier attendance days",
  );
  const monthlyStatusExport = await fetch(
    `${baseUrl}/api/attendance/month/export?month=${attendanceMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  assert.equal(monthlyStatusExport.status, 200, await monthlyStatusExport.clone().text());
  const monthlyStatusExportBuffer = Buffer.from(await monthlyStatusExport.arrayBuffer());
  const monthlyStatusWorkbook = new ExcelJS.Workbook();
  await monthlyStatusWorkbook.xlsx.load(monthlyStatusExportBuffer);
  const monthlyStatusSheet = monthlyStatusWorkbook.getWorksheet(exportedMonthName);
  const expectedRegisteredByGender = { Male: 0, Female: 0 };
  for (const student of noClassAfterExportImportData.students) {
    expectedRegisteredByGender[student.gender] += 1;
    const monthlyStatus = monthlyStatusByStudent.get(student.id);
    if (monthlyStatus === "NLS" || monthlyStatus === "Transferred Out") {
      expectedRegisteredByGender[student.gender] -= 1;
    } else if (monthlyStatus === "Transferred In") {
      expectedRegisteredByGender[student.gender] += 1;
    }
  }
  const expectedRegisteredTotal = expectedRegisteredByGender.Male + expectedRegisteredByGender.Female;
  assert.match(monthlyStatusSheet.getCell("AK123").value.formula, /COUNTIF\(\$AH\$13:\$AH\$62,"NLS\*"\)/);
  assert.match(monthlyStatusSheet.getCell("AK123").value.formula, /COUNTIF\(\$AH\$13:\$AH\$62,"Transferred Out\*"\)/);
  assert.match(monthlyStatusSheet.getCell("AK123").value.formula, /COUNTIF\(\$AH\$13:\$AH\$62,"Transferred In\*"\)/);
  for (const [studentId, status] of monthlyStatusByStudent) {
    const student = noClassAfterExportImportData.students.find((entry) => entry.id === studentId);
    const studentRow = sf2StudentRows.find((rowNumber) => monthlyStatusSheet.getCell(rowNumber, 2).value === student.name);
    assert.match(
      String(monthlyStatusSheet.getCell(studentRow, 34).value),
      new RegExp(`\\[Grade Portal status date: ${attendanceMonth}-15\\]`),
    );
    assert.equal(status, monthlyStatusSheet.getCell(studentRow, 34).value.split("\n")[0]);
  }
  for (const [cellAddress, expected] of [
    ["AK123", expectedRegisteredByGender.Male],
    ["AL123", expectedRegisteredByGender.Female],
    ["AM123", expectedRegisteredTotal],
    ["AM124", expectedRegisteredTotal],
  ]) {
    const result = monthlyStatusSheet.getCell(cellAddress).value.result;
    assert.ok(result === expected || (expected === 0 && result === undefined),
      `${cellAddress} should have cached enrollment count ${expected}`);
  }
  assert.equal(monthlyStatusSheet.getCell("AK132").value.result || 0,
    noClassAfterExportImportData.students.slice(0, 3).filter((student) =>
      monthlyStatusByStudent.get(student.id) === "NLS" && student.gender === "Male").length);
  assert.equal(monthlyStatusSheet.getCell("AM134").value.result, 1);
  assert.equal(monthlyStatusSheet.getCell("AM136").value.result, 1);
  const monthlyStatusClassDays = monthlyStatusSheet.getCell("AJ117").value.result;
  const monthlyStatusMalePresent = monthlyStatusSheet.getCell("AG63").value.result || 0;
  const monthlyStatusFemalePresent = monthlyStatusSheet.getCell("AG114").value.result || 0;
  const monthlyStatusAverage = monthlyStatusSheet.getCell("AM127").value.result || 0;
  assert.equal(
    monthlyStatusAverage,
    (monthlyStatusMalePresent + monthlyStatusFemalePresent) / monthlyStatusClassDays,
    "average daily attendance is calculated from end-of-month registered learner attendance",
  );
  assert.ok((monthlyStatusSheet.getCell("AK127").value.result || 0)
    <= (monthlyStatusSheet.getCell("AK123").value.result || 0));
  assert.ok((monthlyStatusSheet.getCell("AL127").value.result || 0)
    <= (monthlyStatusSheet.getCell("AL123").value.result || 0));
  assert.ok(monthlyStatusAverage <= expectedRegisteredTotal);
  const firstClassDateColumn = noClassAfterExportImportData.dates.findIndex((date) =>
    !noClassAfterExportImportData.noClassDates.includes(date)) + 7;
  const firstClassDateLetter = monthlyStatusSheet.getColumn(firstClassDateColumn).letter;
  const dailySummaryFormula = monthlyStatusSheet.getCell(`${firstClassDateLetter}63`).value.formula;
  assert.ok(dailySummaryFormula.includes(`COUNTIF(${firstClassDateLetter}13:${firstClassDateLetter}62,"N/A")`));
  assert.match(dailySummaryFormula, /COUNTIF\(AH13:AH62,"Transferred Out\*"\)/);
  const followingMonth = new Date(Date.UTC(
    Number(attendanceMonth.slice(0, 4)),
    Number(attendanceMonth.slice(5, 7)),
    1,
  )).toISOString().slice(0, 7);
  const followingMonthResponse = await fetch(
    `${baseUrl}/api/attendance/month?month=${followingMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  assert.equal(followingMonthResponse.status, 200);
  const followingMonthData = await followingMonthResponse.json();
  const inactiveStudents = noClassAfterExportImportData.students.filter((student) =>
    ["NLS", "Transferred Out"].includes(monthlyStatusByStudent.get(student.id)));
  for (const student of inactiveStudents) {
    const followingMonthStudent = followingMonthData.students.find((entry) => entry.id === student.id);
    assert.ok(followingMonthStudent, "inactive learner remains listed in following month");
    assert.equal(followingMonthStudent.remarks, monthlyStatusByStudent.get(student.id));
    assert.equal(followingMonthStudent.attendanceLocked, true);
    assert.equal(followingMonthStudent.attendance.every((status) => status === ""), true);
    assert.equal(followingMonthStudent.lockedDates.length, followingMonthData.dates.length);
  }
  const lockedStudent = followingMonthData.students.find((student) => student.attendanceLocked);
  const editLockedAttendance = await fetch(`${baseUrl}/api/attendance/month`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      month: followingMonth,
      records: followingMonthData.students.map((student) => ({
        studentId: student.id,
        attendance: student.id === lockedStudent.id
          ? student.attendance.map((_, index) => index === 0 ? "Present" : "")
          : student.attendance,
        reasons: student.reasons,
        remarks: student.remarks,
        statusDate: student.statusDate,
      })),
      noClassDates: followingMonthData.noClassDates,
    }),
  });
  assert.equal(editLockedAttendance.status, 400);
  assert.match((await editLockedAttendance.json()).error, /attendance is locked/);
  const editLockedDailyAttendance = await fetch(`${baseUrl}/api/attendance`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      date: followingMonthData.dates[0],
      records: followingMonthData.students.map((student) => ({ studentId: student.id, status: "Present" })),
    }),
  });
  assert.equal(editLockedDailyAttendance.status, 400);
  assert.match((await editLockedDailyAttendance.json()).error, /attendance is locked/);
  const followingMonthExport = await fetch(
    `${baseUrl}/api/attendance/month/export?month=${followingMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  assert.equal(followingMonthExport.status, 200, await followingMonthExport.clone().text());
  const followingMonthWorkbook = new ExcelJS.Workbook();
  await followingMonthWorkbook.xlsx.load(Buffer.from(await followingMonthExport.arrayBuffer()));
  const followingMonthName = `${monthNames[Number(followingMonth.slice(5, 7)) - 1]} ${followingMonth.slice(0, 4)}`;
  const followingMonthSheet = followingMonthWorkbook.getWorksheet(followingMonthName);
  for (const student of inactiveStudents) {
    const rowNumber = sf2StudentRows.find((row) =>
      followingMonthSheet.getCell(row, 2).value === student.name);
    assert.notEqual(rowNumber, undefined, "locked learner remains in the SF2 roster");
    assert.equal(followingMonthSheet.getCell(rowNumber, 7).value, "N/A",
      "locked learner dates are not counted as attendance in SF2");
  }
  const importedStatusMonth = await fetch(
    `${baseUrl}/api/attendance/month/import?month=${attendanceMonth}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream", Cookie: teacherCookie },
      body: monthlyStatusExportBuffer,
    },
  );
  assert.equal(importedStatusMonth.status, 200, await importedStatusMonth.clone().text());
  const statusDateRoundTrip = await fetch(
    `${baseUrl}/api/attendance/month?month=${attendanceMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  const statusDateRoundTripData = await statusDateRoundTrip.json();
  for (const student of statusDateRoundTripData.students.filter((entry) => monthlyStatusByStudent.has(entry.id))) {
    assert.equal(student.statusDate, `${attendanceMonth}-15`, "SF2 import preserves the exact status date");
  }
  const restoredMonthlyRemarks = await fetch(`${baseUrl}/api/attendance/month`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      month: attendanceMonth,
      records: noClassAfterExportImportData.students.map((student) => ({
        studentId: student.id,
        attendance: student.attendance,
        reasons: student.reasons,
        remarks: student.remarks,
      })),
      noClassDates: noClassAfterExportImportData.noClassDates,
    }),
  });
  assert.equal(restoredMonthlyRemarks.status, 200, await restoredMonthlyRemarks.clone().text());
  const saveOtherMonth = await fetch(`${baseUrl}/api/attendance/month`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      month: alternateMonth,
      records: (await (await fetch(
        `${baseUrl}/api/attendance/month?month=${alternateMonth}`,
        { headers: { Cookie: teacherCookie } },
      )).json()).students.map((student) => ({
        studentId: student.id,
        attendance: student.attendance,
        reasons: student.reasons,
        remarks: student.remarks,
      })),
      noClassDates: [],
    }),
  });
  assert.equal(saveOtherMonth.status, 200, await saveOtherMonth.clone().text());
  const noClassAfterOtherMonthSave = await fetch(
    `${baseUrl}/api/attendance/month?month=${attendanceMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  assert.deepEqual((await noClassAfterOtherMonthSave.json()).noClassDates, [noClassDate]);
  const sf2WithHolidayAttendance = await createSF2Upload(attendanceMonth, [
    {
      name: addedStudent.student.name,
      gender: addedStudent.student.gender,
      marks: { [Number(noClassDate.slice(-2))]: "x" },
      remarks: "SF2 test absence",
    },
  ]);
  const importOnNoClassDate = await fetch(
    `${baseUrl}/api/attendance/month/import?month=${attendanceMonth}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/octet-stream", Cookie: teacherCookie },
      body: sf2WithHolidayAttendance,
    },
  );
  assert.equal(importOnNoClassDate.status, 200, await importOnNoClassDate.clone().text());
  const afterImportOnNoClassDate = await fetch(
    `${baseUrl}/api/attendance/month?month=${attendanceMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  const afterImportOnNoClassData = await afterImportOnNoClassDate.json();
  assert.deepEqual(afterImportOnNoClassData.noClassDates, [noClassDate]);
  assert.equal(afterImportOnNoClassData.students[0].attendance[noClassDateIndex], "");
  const attendanceOnNoClassDate = await fetch(`${baseUrl}/api/attendance`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      date: noClassDate,
      records: savedNoClassData.students.map((student) => ({ studentId: student.id, status: "Present" })),
    }),
  });
  assert.equal(attendanceOnNoClassDate.status, 400);
  assert.match((await attendanceOnNoClassDate.json()).error, /marked Holiday\/No Class.*locked/);
  const restoredAttendance = await fetch(`${baseUrl}/api/attendance/month`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      month: attendanceMonth,
      records: attendanceBeforeNoClassData.students.map((student) => ({
        studentId: student.id,
        attendance: student.attendance,
        reasons: student.reasons,
        remarks: student.remarks,
      })),
      noClassDates: [],
    }),
  });
  assert.equal(restoredAttendance.status, 200, await restoredAttendance.clone().text());
  const restoredMonth = await fetch(
    `${baseUrl}/api/attendance/month?month=${attendanceMonth}`,
    { headers: { Cookie: teacherCookie } },
  );
  const restoredMonthData = await restoredMonth.json();
  assert.deepEqual(restoredMonthData.noClassDates, []);
  assert.deepEqual(
    restoredMonthData.students.map((student) => student.attendance[noClassDateIndex]),
    attendanceBeforeNoClassData.students.map((student) => student.attendance[noClassDateIndex]),
  );
  const invalidAttendanceRemark = await fetch(`${baseUrl}/api/attendance/month`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      month: attendanceMonth,
      records: monthlyRecords.map((record, index) => ({
        ...record,
        remarks: index === 0 ? "x".repeat(501) : record.remarks,
      })),
    }),
  });
  assert.equal(invalidAttendanceRemark.status, 400);
  const absentWithoutReason = await fetch(`${baseUrl}/api/attendance/month`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      month: attendanceMonth,
      records: monthlyRecords.map((record, index) => ({
        ...record,
        reasons: index === 0 ? record.reasons.map(() => "") : record.reasons,
      })),
    }),
  });
  assert.equal(absentWithoutReason.status, 400);
  const invalidAttendanceMonth = await fetch(`${baseUrl}/api/attendance/month?month=2026-13`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(invalidAttendanceMonth.status, 400);

  const notExcel = await fetch(`${baseUrl}/api/students/import`, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream", Cookie: teacherCookie },
    body: Buffer.from("This is not an Excel workbook."),
  });
  assert.equal(notExcel.status, 400);
  assert.match((await notExcel.json()).error, /not a supported Excel workbook/i);

  const unauthenticatedImport = await fetch(`${baseUrl}/api/students/import`, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream" },
    body: sf1File,
  });
  assert.equal(unauthenticatedImport.status, 401);

  const duplicateLrn = await fetch(`${baseUrl}/api/students`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify(studentDetails),
  });
  assert.equal(duplicateLrn.status, 409);

  const credentials = {
    name: "Casey Example",
    email: "casey.example@mail.com",
    password: "secure-pass-123",
    role: "parent",
    extra: teacherResult.user.classCode,
  };

  const registration = await request("/api/auth/register", credentials);
  assert.equal(registration.status, 201);
  const parentCookie = registration.headers.get("set-cookie").split(";")[0];
  assert.deepEqual(await registration.json(), {
    user: { name: "Casey Example", role: "parent" },
  });

  const rosterAfterParentRegistration = await fetch(`${baseUrl}/api/students`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(rosterAfterParentRegistration.status, 200);
  assert.equal(
    (await rosterAfterParentRegistration.json()).parents.some(({ email }) => email === credentials.email),
    true,
    "refreshing the roster makes a newly registered parent available for linking",
  );

  const secondParentRegistration = await request("/api/auth/register", {
    name: "Jordan Example",
    email: "jordan.example@mail.com",
    password: "secure-pass-123",
    role: "parent",
    extra: teacherResult.user.classCode,
  });
  assert.equal(secondParentRegistration.status, 201);

  const parentPortalBeforeLink = await fetch(`${baseUrl}/api/parent/portal`, {
    headers: { Cookie: parentCookie },
  });
  assert.equal(parentPortalBeforeLink.status, 200);
  assert.deepEqual((await parentPortalBeforeLink.json()).learners, []);

  const rosterWithParentsResponse = await fetch(`${baseUrl}/api/students`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(rosterWithParentsResponse.status, 200);
  const rosterWithParents = await rosterWithParentsResponse.json();
  assert.ok(rosterWithParents.students.some(({ id }) => id === addedStudent.student.id));
  const registeredParents = rosterWithParents.parents;
  assert.deepEqual(registeredParents.map(({ name, email }) => ({ name, email })), [
    { name: credentials.name, email: credentials.email },
    { name: "Jordan Example", email: "jordan.example@mail.com" },
  ]);
  const assignParent = (parentUserId) => fetch(
    `${baseUrl}/api/students/${encodeURIComponent(addedStudent.student.id)}/parent`,
    {
      method: "PUT",
      headers: { "Content-Type": "application/json", Cookie: teacherCookie },
      body: JSON.stringify({ parentUserId }),
    },
  );
  const invalidParentAssignment = await assignParent("P999999");
  assert.equal(invalidParentAssignment.status, 400);
  const parentAssignment = await assignParent(registeredParents[0].id);
  assert.equal(parentAssignment.status, 200);
  assert.deepEqual(await parentAssignment.json(), {
    saved: true,
    parentUserId: registeredParents[0].id,
  });
  const unlinkedStudent = await assignParent("");
  assert.equal(unlinkedStudent.status, 200);
  assert.deepEqual(await unlinkedStudent.json(), { saved: true, parentUserId: "" });
  const reassignedStudent = await assignParent(registeredParents[0].id);
  assert.equal(reassignedStudent.status, 200);

  const parentPortal = await fetch(`${baseUrl}/api/parent/portal`, {
    headers: { Cookie: parentCookie },
  });
  assert.equal(parentPortal.status, 200);
  const parentPortalResult = await parentPortal.json();
  assert.equal(parentPortalResult.parentName, credentials.name);
  assert.equal(parentPortalResult.email, credentials.email);
  assert.equal(parentPortalResult.classroomNumber, "12");
  assert.equal(parentPortalResult.teacherName, teacherCredentials.name);
  assert.deepEqual(parentPortalResult.learners.map(({ id }) => id), [addedStudent.student.id]);
  assert.equal(parentPortalResult.learners[0].profile.find(({ label }) => label === "LRN").value, studentDetails.lrn);
  assert.equal(parentPortalResult.learners[0].grades["Term 1"].english, 91);
  assert.equal(parentPortalResult.learners[0].grades["Term 1"].mapeh, 88);
  assert.equal(parentPortalResult.learners[0].reading.BOSY.English.level, "Instructional");
  assert.equal(parentPortalResult.learners[0].reading.MOSY.Filipino.level, "Frustration");
  assert.equal(parentPortalResult.learners[0].reading.EOSY.Filipino.level, "Independent");
  const parentAttendanceHistory = parentPortalResult.learners[0].attendance;
  assert.deepEqual(parentAttendanceHistory.filter(({ date }) => !date.startsWith(`${attendanceMonth}-`)), [
    { date: attendanceDate, status: "Present" },
  ]);
  assert.deepEqual(
    parentAttendanceHistory.filter(({ date }) => date.startsWith(`${attendanceMonth}-`))
      .sort((left, right) => left.date.localeCompare(right.date)),
    importedSchoolDays.map((day) => {
      const date = `${attendanceMonth}-${String(day).padStart(2, "0")}`;
      if (day === 1) return { date, status: "Absent", reason: "Medical appointment" };
      return { date, status: day === 2 ? "Half-day" : "Present" };
    }),
  );
  assert.deepEqual(parentPortalResult.gradeTerms, ["Term 1", "Term 2", "Term 3"]);
  const parentLearnerGrades = parentPortalResult.learners[0].grades;
  const coreGradeSubjects = [
    "filipino", "english", "mathematics", "science",
    "aralingPanlipunan", "gmrc", "eppTle", "mapeh",
  ];
  for (const term of parentPortalResult.gradeTerms) {
    const grades = coreGradeSubjects.map((subject) => parentLearnerGrades[term][subject]);
    assert.ok(grades.every((grade) => typeof grade === "number"), `${term} core grades are available for the parent average`);
    const generalAverage = grades.reduce((total, grade) => total + grade, 0) / grades.length;
    assert.equal(Number.isFinite(generalAverage), true, `${term} General Average can be calculated`);
  }
  for (const subject of parentPortalResult.gradeSubjects) {
    const termGrades = parentPortalResult.gradeTerms.map((term) => parentLearnerGrades[term][subject.key]);
    assert.ok(termGrades.every((grade) => typeof grade === "number"), `${subject.label} has all three term grades`);
    const generalGrade = termGrades.reduce((total, grade) => total + grade, 0) / termGrades.length;
    assert.equal(Number.isFinite(generalGrade), true, `${subject.label} General grade can be calculated`);
  }
  assert.match(portalPage, /headerRow\.append\(generalHeader\)/);
  assert.match(portalPage, /generalHeader\.textContent="Average"/);
  assert.match(portalPage, /generalAverageRow\.append\(overallAverageCell\)/);
  assert.match(portalPage, /allCoreGrades\.push\(learner\.grades\[term\]\[subject\]\)/);
  assert.match(portalPage, /overallAverageCell\.className="parent-grade-overall"/);
  assert.match(portalPage, /allCoreGrades\.reduce\(function\(total,value\)\{return total\+value\},0\)\/allCoreGrades\.length/);
  assert.doesNotMatch(portalPage, /termAverages/);
  assert.doesNotMatch(portalPage, /parent-grade-tabs?(?:["'.\s{]|$)/);
  assert.deepEqual(parentPortalResult.readingPeriods, ["BOSY", "MOSY", "EOSY"]);
  assert.deepEqual(parentPortalResult.learners[0].reading.BOSY.English.attempts.map((attempt) => attempt.attemptNo), [1, 2]);
  assert.match(parentPortalResult.learners[0].reading.BOSY.English.attempts[1].assessedDate, /^\d{4}-\d{2}-\d{2}$/);
  assert.equal(parentPortalResult.learners[0].reading.BOSY.English.level, "Instructional");
  assert.equal(parentPortalResult.learners[0].reading.BOSY.English.attemptNo, 2);
  assert.match(portalPage, /id="readingAttemptNumber" type="number" min="1" step="1" value="1" inputmode="numeric" required/);
  assert.match(portalPage, /Date set automatically when saved/);
  assert.match(portalPage, /saved\.assessedDates\[language\]/);
  assert.match(portalPage, /Assessment history/);
  assert.match(portalPage, /button\.dataset\.parentTarget==="parentReadingSection"[\s\S]*?refreshParentPortalData\(\$\("refreshParentReading"\)/);
  assert.match(portalPage, /id="refreshParentReading" type="button">Refresh reading progress/);
  assert.match(portalPage, /async function refreshParentPortalData\(button,successMessage\)/);
  assert.match(portalPage, /var updatedPortal=await api\("\/api\/parent\/portal"\)/);
  assert.match(portalPage, /renderParentLearner\(nextIndex<0\?0:nextIndex\)/);
  const teacherParentPortal = await fetch(`${baseUrl}/api/parent/portal`, {
    headers: { Cookie: teacherCookie },
  });
  assert.equal(teacherParentPortal.status, 403);
  const unauthenticatedParentPortal = await fetch(`${baseUrl}/api/parent/portal`);
  assert.equal(unauthenticatedParentPortal.status, 401);
  const thirdReadingAttempt = await fetch(`${baseUrl}/api/reading-assessments`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({
      studentId: addedStudent.student.id,
      period: "BOSY",
      attemptNo: 3,
      levels: { English: "Independent", Filipino: "Instructional" },
    }),
  });
  assert.equal(thirdReadingAttempt.status, 200);
  const refreshedParentPortal = await fetch(`${baseUrl}/api/parent/portal`, {
    headers: { Cookie: parentCookie },
  });
  assert.equal(refreshedParentPortal.status, 200);
  const refreshedParentPortalResult = await refreshedParentPortal.json();
  assert.equal(refreshedParentPortalResult.learners[0].reading.BOSY.English.level, "Independent");
  assert.equal(refreshedParentPortalResult.learners[0].reading.BOSY.English.attemptNo, 3);
  assert.deepEqual(refreshedParentPortalResult.learners[0].reading.BOSY.English.attempts.map((attempt) => attempt.attemptNo), [1, 2, 3]);

  const parentStudentList = await fetch(`${baseUrl}/api/students`, {
    headers: { Cookie: parentCookie },
  });
  assert.equal(parentStudentList.status, 403);
  const parentCannotAssignParent = await fetch(
    `${baseUrl}/api/students/${encodeURIComponent(addedStudent.student.id)}/parent`,
    {
      method: "PUT",
      headers: { "Content-Type": "application/json", Cookie: parentCookie },
      body: JSON.stringify({ parentUserId: "" }),
    },
  );
  assert.equal(parentCannotAssignParent.status, 403);
  const parentCannotUnlink = await fetch(
    `${baseUrl}/api/students/${encodeURIComponent(addedStudent.student.id)}/parent`,
    {
      method: "DELETE",
      headers: { Cookie: parentCookie },
    },
  );
  assert.equal(parentCannotUnlink.status, 403);
  const parentGrades = await fetch(gradesUrl, { headers: { Cookie: parentCookie } });
  assert.equal(parentGrades.status, 403);
  const parentGradeSummary = await fetch(`${baseUrl}/api/grades/summary`, {
    headers: { Cookie: parentCookie },
  });
  assert.equal(parentGradeSummary.status, 403);
  const parentReading = await fetch(`${baseUrl}/api/reading-assessments`, {
    headers: { Cookie: parentCookie },
  });
  assert.equal(parentReading.status, 403);
  const parentAttendance = await fetch(`${baseUrl}/api/attendance?date=${attendanceDate}`, {
    headers: { Cookie: parentCookie },
  });
  assert.equal(parentAttendance.status, 403);
  const parentMonthlyAttendance = await fetch(
    `${baseUrl}/api/attendance/month?month=${attendanceMonth}`,
    { headers: { Cookie: parentCookie } },
  );
  assert.equal(parentMonthlyAttendance.status, 403);
  const parentProfile = await fetch(`${baseUrl}/api/teacher/profile`, {
    headers: { Cookie: parentCookie },
  });
  assert.equal(parentProfile.status, 403);
  const parentProfilePhoto = await fetch(`${baseUrl}/api/teacher/profile/photo`, {
    headers: { Cookie: parentCookie },
  });
  assert.equal(parentProfilePhoto.status, 403);
  const parentPassword = await fetch(`${baseUrl}/api/teacher/profile/password`, {
    method: "PUT",
    headers: { "Content-Type": "application/json", Cookie: parentCookie },
    body: JSON.stringify({ currentPassword: "secure-pass-123", newPassword: "other-secure-pass-123" }),
  });
  assert.equal(parentPassword.status, 403);
  const duplicateEmailProfile = await updateProfile(teacherCookie, {
    name: teacherCredentials.name,
    email: credentials.email,
    classroomNumber: "12",
  });
  assert.equal(duplicateEmailProfile.status, 409);

  const duplicate = await request("/api/auth/register", credentials);
  assert.equal(duplicate.status, 409);

  const failedLogin = await request("/api/auth/login", {
    email: credentials.email,
    password: "incorrect-password",
    role: credentials.role,
  });
  assert.equal(failedLogin.status, 401);

  const persistentGradeLock = await fetch(`${baseUrl}/api/grades/lock`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: teacherCookie },
    body: JSON.stringify({ studentId: addedStudent.student.id, term: "Term 1", locked: true }),
  });
  assert.equal(persistentGradeLock.status, 200);

  await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
  server = createServer({ databasePath, migrationWorkbookPath });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));

  const restartedUrl = `http://127.0.0.1:${server.address().port}`;
  const login = await fetch(`${restartedUrl}/api/auth/login`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      email: credentials.email,
      password: credentials.password,
      role: credentials.role,
    }),
  });
  assert.equal(login.status, 200);
  assert.deepEqual(await login.json(), {
    user: { name: "Casey Example", role: "parent" },
  });
  const restartedParentCookie = login.headers.get("set-cookie").split(";")[0];
  const persistedParentPortal = await fetch(`${restartedUrl}/api/parent/portal`, {
    headers: { Cookie: restartedParentCookie },
  });
  assert.equal(persistedParentPortal.status, 200);
  const persistedParentPortalResult = await persistedParentPortal.json();
  assert.deepEqual(persistedParentPortalResult, {
    ...refreshedParentPortalResult,
  });
  assert.deepEqual(persistedParentPortalResult.learners[0].attendance, parentAttendanceHistory);

  const teacherLogin = await fetch(`${restartedUrl}/api/auth/login`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      email: teacherCredentials.email,
      password: teacherCredentials.password,
      role: teacherCredentials.role,
    }),
  });
  assert.equal(teacherLogin.status, 200);
  const restartedTeacherCookie = teacherLogin.headers.get("set-cookie").split(";")[0];
  const persistedProfileResponse = await fetch(`${restartedUrl}/api/teacher/profile`, {
    headers: { Cookie: restartedTeacherCookie },
  });
  assert.equal(persistedProfileResponse.status, 200);
  assert.deepEqual(await persistedProfileResponse.json(), {
    name: teacherCredentials.name,
    email: teacherCredentials.email,
    classroomNumber: "12",
    grade: "Grade 5",
    section: "Mabini",
    schoolYear: "2026-2027",
    schoolName: "Sample Elementary School",
    schoolId: "123456",
    schoolHead: "Alex Principal",
    hasPhoto: false,
  });
  const persistedGradeLockResponse = await fetch(
    `${restartedUrl}/api/grades/lock?studentId=${encodeURIComponent(addedStudent.student.id)}&term=Term%201`,
    {
      headers: { Cookie: restartedTeacherCookie },
    },
  );
  assert.equal(persistedGradeLockResponse.status, 200);
  assert.deepEqual(await persistedGradeLockResponse.json(), {
    studentId: addedStudent.student.id,
    term: "Term 1",
    locked: true,
  });
  const persistedLockedGradeSave = await fetch(`${restartedUrl}/api/grades`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: restartedTeacherCookie },
    body: JSON.stringify({
      studentId: addedStudent.student.id,
      term: "Term 1",
      grades: termGrades,
    }),
  });
  assert.equal(persistedLockedGradeSave.status, 409);
  const persistedGradeUnlock = await fetch(`${restartedUrl}/api/grades/lock`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: restartedTeacherCookie },
    body: JSON.stringify({ studentId: addedStudent.student.id, term: "Term 1", locked: false }),
  });
  assert.equal(persistedGradeUnlock.status, 200);
  const persistedGradesUrl = `${restartedUrl}/api/grades?studentId=${encodeURIComponent(addedStudent.student.id)}&term=Term%201`;
  const persistedGradesResponse = await fetch(persistedGradesUrl, {
    headers: { Cookie: restartedTeacherCookie },
  });
  assert.equal(persistedGradesResponse.status, 200);
  assert.deepEqual((await persistedGradesResponse.json()).grades, {
    filipino: "88",
    english: "91",
    mathematics: "86.5",
    science: "90",
    aralingPanlipunan: "89",
    gmrc: "95",
    eppTle: "87",
    musicAndArts: "92",
    physicalEducationAndHealth: "84",
  });
  const persistedSummaryResponse = await fetch(`${restartedUrl}/api/grades/summary`, {
    headers: { Cookie: restartedTeacherCookie },
  });
  assert.equal(persistedSummaryResponse.status, 200);
  const persistedSummary = await persistedSummaryResponse.json();
  assert.deepEqual(persistedSummary.terms, ["Term 1", "Term 2", "Term 3"]);
  assert.equal(persistedSummary.students[0].grades["Term 1"].filipino, 88);
  assert.equal(persistedSummary.students[0].grades["Term 2"].filipino, 87);
  assert.equal(persistedSummary.students[0].grades["Term 3"].filipino, 86);
  const persistedReadingResponse = await fetch(`${restartedUrl}/api/reading-assessments?period=BOSY`, {
    headers: { Cookie: restartedTeacherCookie },
  });
  assert.equal(persistedReadingResponse.status, 200);
  const persistedReading = await persistedReadingResponse.json();
  assert.equal(persistedReading.students[0].levels.English, "Independent");
  assert.equal(persistedReading.students[0].levels.Filipino, "Instructional");
  const persistedSecondAttemptResponse = await fetch(`${restartedUrl}/api/reading-assessments?period=BOSY&attemptNo=2`, {
    headers: { Cookie: restartedTeacherCookie },
  });
  assert.equal(persistedSecondAttemptResponse.status, 200);
  assert.equal((await persistedSecondAttemptResponse.json()).students[0].levels.English, "Instructional");
  const persistedAttendance = await fetch(`${restartedUrl}/api/attendance?date=${attendanceDate}`, {
    headers: { Cookie: restartedTeacherCookie },
  });
  assert.equal(persistedAttendance.status, 200);
  assert.equal((await persistedAttendance.json()).students[0].status, "Present");
  const persistedAttendanceMonth = await fetch(
    `${restartedUrl}/api/attendance/month?month=${attendanceMonth}`,
    { headers: { Cookie: restartedTeacherCookie } },
  );
  assert.equal(persistedAttendanceMonth.status, 200);
  const persistedAttendanceMonthResult = await persistedAttendanceMonth.json();
  assert.equal(persistedAttendanceMonthResult.students[0].attendance[0], "Absent");
  assert.equal(persistedAttendanceMonthResult.students[0].reasons[0], "Medical appointment");
  assert.equal(
    persistedAttendanceMonthResult.students[0].attendance[attendanceDates.length - 1],
    "Present",
  );
  assert.equal(persistedAttendanceMonthResult.students[0].remarks, "Monthly follow-up");
  for (const assessment of [
    { period: "MOSY", english: "Instructional", filipino: "Frustration" },
    { period: "EOSY", english: "Independent", filipino: "Independent" },
  ]) {
    const response = await fetch(`${restartedUrl}/api/reading-assessments?period=${assessment.period}`, {
      headers: { Cookie: restartedTeacherCookie },
    });
    assert.equal(response.status, 200);
    const report = await response.json();
    assert.equal(report.students[0].levels.English, assessment.english);
    assert.equal(report.students[0].levels.Filipino, assessment.filipino);
  }

  const clearReadingResponse = await fetch(`${restartedUrl}/api/reading-assessments`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: restartedTeacherCookie },
    body: JSON.stringify({
      studentId: addedStudent.student.id,
      period: "BOSY",
      levels: { English: "", Filipino: "" },
    }),
  });
  assert.equal(clearReadingResponse.status, 200);
  const clearedReadingResponse = await fetch(`${restartedUrl}/api/reading-assessments`, {
    headers: { Cookie: restartedTeacherCookie },
  });
  const clearedReading = await clearedReadingResponse.json();
  assert.equal(clearedReading.students[0].levels.English, "");
  assert.equal(clearedReading.students[0].levels.Filipino, "");
  const preservedSecondAttemptResponse = await fetch(`${restartedUrl}/api/reading-assessments?period=BOSY&attemptNo=2`, {
    headers: { Cookie: restartedTeacherCookie },
  });
  assert.equal((await preservedSecondAttemptResponse.json()).students[0].levels.English, "Instructional");

  const workbook = readPortalDatabaseWorkbook(databasePath);
  const users = workbook.getWorksheet("Users");
  const headerMap = new Map();
  users.getRow(1).eachCell((cell, column) => headerMap.set(String(cell.value).toLowerCase(), column));
  const emailColumn = headerMap.get("email");
  const row = users.getRows(2, users.rowCount - 1)
    .find((candidate) => String(candidate.getCell(emailColumn).value).toLowerCase() === credentials.email);
  const teacherRow = users.getRows(2, users.rowCount - 1)
    .find((candidate) => String(candidate.getCell(emailColumn).value).toLowerCase() === teacherCredentials.email);

  assert.ok(row, "registered user is present in the Users sheet");
  assert.equal(row.number, 3, "accounts are written into the workbook's existing user rows");
  assert.match(String(row.getCell(headerMap.get("password hash (pbkdf2)")).value), /^pbkdf2\$/);
  assert.equal(row.getCell(headerMap.get("sign-ins")).value, 1);
  assert.ok(row.getCell(headerMap.get("last sign-in")).value instanceof Date);
  assert.ok(teacherRow, "teacher registration is present in the Users sheet");
  assert.equal(teacherRow.number, 2);
  assert.match(String(teacherRow.getCell(headerMap.get("password hash (pbkdf2)")).value), /^pbkdf2\$/);
  const classSheet = workbook.getWorksheet("Classes");
  const classHeaders = new Map();
  classSheet.getRow(1).eachCell((cell, column) => classHeaders.set(String(cell.value).toLowerCase(), column));
  const classRow = classSheet.getRows(2, classSheet.rowCount - 1)
    .find((candidate) => candidate.getCell(classHeaders.get("class code (parents enter this)")).value === teacherResult.user.classCode);
  assert.ok(classRow, "the teacher's class and parent join code are saved");
  assert.equal(classRow.number, 2);
  assert.equal(classRow.getCell(classHeaders.get("students")).value, 3);
  assert.equal(classRow.getCell(classHeaders.get("parents signed up")).value, 2);
  const gradeSheet = workbook.getWorksheet("Term Grades");
  assert.ok(gradeSheet, "term grades are persisted in a dedicated worksheet");
  const gradeHeaders = new Map();
  gradeSheet.getRow(1).eachCell((cell, column) => gradeHeaders.set(String(cell.value).toLowerCase(), column));
  const savedFilipinoGrade = gradeSheet.getRows(2, gradeSheet.rowCount - 1)
    .find((candidate) => candidate.getCell(gradeHeaders.get("subject")).value === "filipino"
      && candidate.getCell(gradeHeaders.get("term")).value === "Term 1");
  assert.ok(savedFilipinoGrade, "the student's Filipino grade is saved");
  assert.equal(savedFilipinoGrade.getCell(gradeHeaders.get("grade")).value, 88);
  const studentSheet = workbook.getWorksheet("Students");
  const studentHeaders = new Map();
  studentSheet.getRow(1).eachCell((cell, column) => {
    studentHeaders.set(String(cell.value).toLowerCase(), column);
  });
  assert.equal(studentSheet.getRow(2).getCell(studentHeaders.get("lrn")).value, studentDetails.lrn);
  assert.equal(studentSheet.getRow(2).getCell(studentHeaders.get("teacher user id")).value, "T001");
  assert.equal(studentSheet.getRow(2).getCell(studentHeaders.get("birthdate")).value.toISOString().slice(0, 10), studentDetails.birthdate);
  const signInLog = workbook.getWorksheet("SignInLog");
  const logHeaders = new Map();
  signInLog.getRow(1).eachCell((cell, column) => logHeaders.set(String(cell.value).toLowerCase(), column));
  const attempts = [];
  for (let rowNumber = 2; rowNumber <= signInLog.rowCount; rowNumber += 1) {
    const logRow = signInLog.getRow(rowNumber);
    if (String(logRow.getCell(logHeaders.get("email")).value).toLowerCase() === credentials.email) {
      attempts.push(logRow.getCell(logHeaders.get("result")).value);
    }
  }
  assert.deepEqual(attempts, ["Failed", "Success"]);

  const rejectedRosterPassword = await fetch(`${restartedUrl}/api/students/clear/verify`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: restartedTeacherCookie },
    body: JSON.stringify({ password: "incorrect-password" }),
  });
  assert.equal(rejectedRosterPassword.status, 400);
  assert.match((await rejectedRosterPassword.json()).error, /password doesn't match/i);
  const verifyRosterClear = await fetch(`${restartedUrl}/api/students/clear/verify`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: restartedTeacherCookie },
    body: JSON.stringify({ password: teacherCredentials.password }),
  });
  assert.equal(verifyRosterClear.status, 200);
  assert.deepEqual(await verifyRosterClear.json(), { verified: true });
  const rosterBeforeClear = await fetch(`${restartedUrl}/api/students`, {
    headers: { Cookie: restartedTeacherCookie },
  });
  assert.equal((await rosterBeforeClear.json()).students.length, 3);
  const unconfirmedRosterClear = await fetch(`${restartedUrl}/api/students/clear`, {
    method: "DELETE",
    headers: { "Content-Type": "application/json", Cookie: restartedTeacherCookie },
    body: JSON.stringify({ password: teacherCredentials.password, confirmed: false }),
  });
  assert.equal(unconfirmedRosterClear.status, 400);
  const rejectedRosterClear = await fetch(`${restartedUrl}/api/students/clear`, {
    method: "DELETE",
    headers: { "Content-Type": "application/json", Cookie: restartedTeacherCookie },
    body: JSON.stringify({ password: "incorrect-password", confirmed: true }),
  });
  assert.equal(rejectedRosterClear.status, 400);
  const clearRosterResponse = await fetch(`${restartedUrl}/api/students/clear`, {
    method: "DELETE",
    headers: { "Content-Type": "application/json", Cookie: restartedTeacherCookie },
    body: JSON.stringify({ password: teacherCredentials.password, confirmed: true }),
  });
  assert.equal(clearRosterResponse.status, 200);
  assert.deepEqual(await clearRosterResponse.json(), { deleted: 3 });
  const clearedRosterResponse = await fetch(`${restartedUrl}/api/students`, {
    headers: { Cookie: restartedTeacherCookie },
  });
  const clearedRoster = await clearedRosterResponse.json();
  assert.deepEqual(clearedRoster.students, []);
  assert.equal(clearedRoster.parents.length, 2, "clearing the roster preserves parent accounts");
  const clearedDatabase = readPortalDatabaseWorkbook(databasePath);
  for (const sheetName of ["Students", "Term Grades", "Term Grade Locks", "Reading Assessments", "Attendance Records", "Attendance Monthly Notes"]) {
    const sheet = clearedDatabase.getWorksheet(sheetName);
    if (sheet) {
      const studentIdColumn = new Map(
        sheet.getRow(1).values.slice(1).map((value, index) => [String(value).toLowerCase(), index + 1]),
      ).get("student id");
      if (studentIdColumn) {
        for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
          assert.equal(sheet.getRow(rowNumber).getCell(studentIdColumn).value || "", "", `${sheetName} has no remaining learner rows`);
        }
      }
    }
  }
  const clearedClassSheet = clearedDatabase.getWorksheet("Classes");
  const clearedClassHeaders = new Map();
  clearedClassSheet.getRow(1).eachCell((cell, column) => {
    clearedClassHeaders.set(String(cell.value).toLowerCase(), column);
  });
  const clearedClassRow = clearedClassSheet.getRows(2, clearedClassSheet.rowCount - 1)
    .find((candidate) => candidate.getCell(clearedClassHeaders.get("class code (parents enter this)")).value === teacherResult.user.classCode);
  assert.equal(clearedClassRow.getCell(clearedClassHeaders.get("students")).value, 0);
});

test("registration rejects a class code that is not in the workbook", async (context) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "grade5-portal-"));
  const databasePath = path.join(directory, "database.sqlite");
  const migrationWorkbookPath = path.join(directory, "migration.xlsx");
  await createEmptyTestWorkbook(migrationWorkbookPath);
  const server = createServer({ databasePath, migrationWorkbookPath });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  context.after(async () => {
    await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
    await fs.rm(directory, { recursive: true, force: true });
  });

  const response = await fetch(`http://127.0.0.1:${server.address().port}/api/auth/register`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: "Casey Example",
      email: "casey.example@mail.com",
      password: "secure-pass-123",
      role: "parent",
      extra: "not-a-real-code",
    }),
  });

  assert.equal(response.status, 400);
  assert.match((await response.json()).error, /teacher to sign up first/i);
});

test("teacher registration rejects a duplicate classroom number", async (context) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "grade5-portal-"));
  const databasePath = path.join(directory, "database.sqlite");
  const migrationWorkbookPath = path.join(directory, "migration.xlsx");
  await createEmptyTestWorkbook(migrationWorkbookPath);
  const server = createServer({ databasePath, migrationWorkbookPath });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  context.after(async () => {
    await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
    await fs.rm(directory, { recursive: true, force: true });
  });

  const register = (name, email) => fetch(`http://127.0.0.1:${server.address().port}/api/auth/register`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name,
      email,
      password: "secure-pass-123",
      role: "teacher",
      extra: "12",
    }),
  });

  assert.equal((await register("Taylor Teacher", "taylor.teacher@example.com")).status, 201);
  const duplicate = await register("Another Teacher", "another.teacher@example.com");
  assert.equal(duplicate.status, 409);
  assert.match((await duplicate.json()).error, /classroom number is already registered/i);
});

test("SQLite migration preserves existing portal accounts", async (context) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), "grade5-portal-migration-"));
  const migrationWorkbookPath = path.join(directory, "existing.xlsx");
  const databasePath = path.join(directory, "portal.sqlite");
  await fs.copyFile(sourceWorkbook, migrationWorkbookPath);
  const server = createServer({ databasePath, migrationWorkbookPath });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  context.after(async () => {
    if (server.listening) {
      await new Promise((resolve, reject) =>
        server.close((error) => error ? reject(error) : resolve()));
    }
    await fs.rm(directory, { recursive: true, force: true });
  });

  const response = await fetch(`http://127.0.0.1:${server.address().port}/api/health`);
  assert.equal(response.status, 200);
  assert.deepEqual(await response.json(), { status: "ok", storage: "sqlite" });
  const migratedWorkbook = readPortalDatabaseWorkbook(databasePath);
  const users = migratedWorkbook.getWorksheet("Users");
  const headers = new Map();
  users.getRow(1).eachCell((cell, column) => headers.set(String(cell.value).toLowerCase(), column));
  const emailColumn = headers.get("email");
  assert.ok(emailColumn);
  const migratedAccounts = [];
  for (let rowNumber = 2; rowNumber <= users.rowCount; rowNumber += 1) {
    const email = users.getRow(rowNumber).getCell(emailColumn).value;
    if (email) migratedAccounts.push(email);
  }
  assert.ok(migratedAccounts.length > 0, "existing accounts are present after migration");
});
