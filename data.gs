// data.gs

// 🔁 Parse the COVER sheet into the EMPMASTER object
function UpdateEmpMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("COVER");
  const data = sheet.getDataRange().getValues();

  let currentRole = "MANAGER";
  const parsed = { MANAGER: {}, INSIDER: {}, DRIVER: {} };
  let employeeCount = 0;

  for (let i = 9; i < data.length; i++) {
    const nameCell = (data[i][1] || "").toString().trim();
    const upperName = nameCell.toUpperCase();

    if (upperName === "INSIDER" || upperName === "DRIVER") {
      currentRole = upperName;
      continue;
    }

    if (!nameCell) continue;

    parsed[currentRole][nameCell] = [
      data[i][2], data[i][3], data[i][4],
      data[i][5], data[i][6], data[i][7],
      data[i][8], data[i][9], data[i][10]
    ];

    employeeCount++;
  }

  if (employeeCount > 0) {
    EMPMASTER = parsed;
  }
}

// 💾 Save EMPMASTER into document properties (for recovery)
function SaveEmpMasterBackup() {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty("EMPMASTER_BACKUP", JSON.stringify(EMPMASTER));
}

function NormalizePosition(value) {
  return value.toString().trim().toUpperCase();
}

function IsValidPosition(position) {
  return ["MANAGER", "INSIDER", "DRIVER"].includes(position);
}

function ResetInputForm() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("INPUT");

  sheet.getRange("C3").clearContent();
  sheet.getRange("D3").clearContent();
  sheet.getRange("F3:L3").clearContent();
  sheet.getRange("F7").clearContent();
}

// ➕ Add a new employee to EMPMASTER from INPUT sheet
function AddEmployee() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("INPUT");

  if (!ValidateEntry("add")) return;

  UpdateEmpMaster();

  const rawName = sheet.getRange("C3").getValue().toString().trim();
  const name = rawName.replace(/\s+/g, " ");
  const position = NormalizePosition(sheet.getRange("D3").getValue());
  const availability = sheet.getRange("F3:L3").getValues()[0];
  const notes = sheet.getRange("F7").getValue().toString().trim();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM.dd.yy");

  if (!IsValidPosition(position)) {
    SpreadsheetApp.getUi().alert("⚠️ Position must be Manager, Insider, or Driver.");
    return;
  }

  if (!EMPMASTER[position]) EMPMASTER[position] = {};

  if (EMPMASTER[position][name]) {
    SpreadsheetApp.getUi().alert("⚠️ Cannot add. Employee already exists in this position.");
    return;
  }

  EMPMASTER[position][name] = [...availability, timestamp, notes];

  SaveEmpMasterBackup();
  UpdateCover();
  ResetInputForm();
}

// 📝 Update an existing employee’s availability
function UpdateEmployee() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("INPUT");

  if (!ValidateEntry("update")) return;

  UpdateEmpMaster();

  const rawName = sheet.getRange("C3").getValue().toString().trim();
  const name = rawName.replace(/\s+/g, " ");
  const position = NormalizePosition(sheet.getRange("D3").getValue());
  const availability = sheet.getRange("F3:L3").getValues()[0];
  const notes = sheet.getRange("F7").getValue().toString().trim();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM.dd.yy");

  if (!IsValidPosition(position)) {
    SpreadsheetApp.getUi().alert("⚠️ Position must be Manager, Insider, or Driver.");
    return;
  }

  if (EMPMASTER[position] && EMPMASTER[position][name]) {
    EMPMASTER[position][name] = [...availability, timestamp, notes];

    SaveEmpMasterBackup();
    UpdateCover();
    ResetInputForm();
  } else {
    SpreadsheetApp.getUi().alert("⚠️ Cannot update. Employee not found.");
  }
}

// ❌ Remove an employee from EMPMASTER
function RemoveEmployee() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("INPUT");

  if (!ValidateEntry("remove")) return;

  UpdateEmpMaster();

  const rawName = sheet.getRange("C3").getValue().toString().trim();
  const name = rawName.replace(/\s+/g, " ");
  const position = NormalizePosition(sheet.getRange("D3").getValue());

  if (!IsValidPosition(position)) {
    SpreadsheetApp.getUi().alert("⚠️ Position must be Manager, Insider, or Driver.");
    return;
  }

  if (EMPMASTER[position] && EMPMASTER[position][name]) {
    delete EMPMASTER[position][name];

    SaveEmpMasterBackup();
    UpdateCover();
    ResetInputForm();
  } else {
    SpreadsheetApp.getUi().alert("⚠️ Cannot remove. Employee not found.");
  }
}

function ClearEmployeeDatabase() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    "⚠️ Clear Employee Database",
    "This will permanently remove all employees from the database and clear them from the sheet. Continue?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  EMPMASTER = { MANAGER: {}, INSIDER: {}, DRIVER: {} };

  const props = PropertiesService.getDocumentProperties();
  props.setProperty("EMPMASTER_BACKUP", JSON.stringify(EMPMASTER));

  ResetInputForm();
  UpdateCover();
  Restore();

  ui.alert("✅ Employee database cleared.");
}
