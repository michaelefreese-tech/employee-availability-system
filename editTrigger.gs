// editTrigger.gs

// 🎯 Main onEdit trigger — intercepts user changes
function onEdit(e) {
  const ss = e.source;
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  const range = e.range;
  const editedCell = range.getA1Notation();
  const row = range.getRow();

  const props = PropertiesService.getDocumentProperties();
  const isLocked = props.getProperty("docLocked") === "true";

  const cellKey = sheetName + "!" + editedCell;
  const backup = TempCellBackup?.[cellKey] || { value: "", formula: "" };

  // If document is not locked, allow edit
  if (!isLocked) return;

  // ================================
  // SHEET: COVER
  // ================================
  if (sheetName === "COVER") {
    Logger.log(`Edited COVER cell: ${editedCell}`);

    const ranges = LogEmployeeRanges();
    let foundGroup = null;

    for (const group in ranges) {
      if (ranges[group].includes(row)) {
        foundGroup = group;
        break;
      }
    }

    if (foundGroup) {
      const groupStart = ranges[foundGroup][0];
      const rowIndex = row - groupStart;
      RestoreAvail(foundGroup, rowIndex, row);
    } else {
      SpreadsheetApp.getUi().alert("⛔ This area is locked.");

      Restore();

      const currentWeekIndex = parseInt(props.getProperty("currentWeekIndex"), 10) || 0;
      UpdateWeekView(currentWeekIndex);

      LoadEmpMaster();
      UpdateCover();
    }

    return;
  }

  // ================================
  // SHEET: INPUT
  // ================================
  if (sheetName === "INPUT") {
    if (Info["INPUT"].ALLOWEDCELLS.includes(editedCell)) {
      return;
    }

    SpreadsheetApp.getUi().alert("⛔ You don't have permission to edit this cell: " + editedCell);
    range.setValue(backup.value);
    return;
  }

  // ================================
  // SHEET: INFO
  // ================================
  if (sheetName === "INFO") {
    if (!Info["INFO"].ALLOWEDCELLS.includes(editedCell)) {
      SpreadsheetApp.getUi().alert("⛔ You don't have permission to edit this INFO cell: " + editedCell);
      range.setValue(backup.value);

      Restore();

      const currentWeekIndex = parseInt(props.getProperty("currentWeekIndex"), 10) || 0;
      UpdateWeekView(currentWeekIndex);

      LoadEmpMaster();
      UpdateCover();
    }

    return;
  }
}

function onChange(e) {
  const props = PropertiesService.getDocumentProperties();
  const isLocked = props.getProperty("docLocked") === "true";
  if (!isLocked) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingNames = ss.getSheets().map(s => s.getName());

  REQUIRED_SHEETS.forEach(name => {
    if (!existingNames.includes(name)) {
      ss.insertSheet(name);
    }
  });

  // If anything structural changed OR sheets were missing → full rebuild
  LoadEmpMaster();
  Restore();

  const currentWeekIndex = parseInt(props.getProperty("currentWeekIndex"), 10) || 0;
  UpdateWeekView(currentWeekIndex);

  UpdateCover();
  ApplySheetProtections();
}

function InstallChangeTrigger() {
  const ss = SpreadsheetApp.getActive();

  const existing = ScriptApp.getProjectTriggers();
  existing.forEach(t => {
    if (t.getHandlerFunction() === "onChange") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("onChange")
    .forSpreadsheet(ss)
    .onChange()
    .create();
}

// 📊 Log row ranges for each role in COVER
function LogEmployeeRanges() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COVER");
  const lastRow = sheet.getLastRow();
  const ranges = { MANAGER: [], INSIDER: [], DRIVER: [] };

  let currentRole = "MANAGER";

  for (let r = 10; r <= lastRow; r++) {
    const cellValue = sheet.getRange(r, 2).getValue().toString().trim().toUpperCase();

    if (cellValue === "INSIDER") {
      currentRole = "INSIDER";
      continue;
    }

    if (cellValue === "DRIVER") {
      currentRole = "DRIVER";
      continue;
    }

    if (cellValue) {
      ranges[currentRole].push(r);
    }
  }

  return ranges;
}

// 🧯 Restore one employee's row from backup
function RestoreAvail(group, index, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const coverSheet = ss.getSheetByName("COVER");

  const rawBackup = PropertiesService.getDocumentProperties().getProperty("EMPMASTER_BACKUP");
  const backup = rawBackup ? JSON.parse(rawBackup) : {};

  const names = Object.keys(backup[group] || {});
  const name = names[index];
  const data = backup[group]?.[name];

  if (!data) {
    SpreadsheetApp.getUi().alert(`⚠️ No data found at index ${index} in group '${group}'.`);
    return;
  }

  // Restore full row
  coverSheet.getRange(row, 2).setValue(name);

  // C:I = availability
  for (let i = 0; i < 7; i++) {
    coverSheet.getRange(row, 3 + i).setValue(data[i]);
  }

  // J = updated on, K = notes
  coverSheet.getRange(row, 10).setValue(data[7] || "");
  coverSheet.getRange(row, 11).setValue(data[8] || "");

  SpreadsheetApp.getUi().alert(
    `✅ Restored ${name}'s row:\n` +
    `Availability: ${data.slice(0, 7).join(", ")}\n` +
    `Last Updated: ${data[7] || ""}\n` +
    `Notes: ${data[8] || ""}`
  );
}

function EmergencyUnlock() {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty("docLocked", "false");
}

function EmergencyRestoreEmployees() {
  LoadEmpMaster();
  Restore();
  UpdateCover();
}

function LogEmployeeBackup() {
  const backup = PropertiesService.getDocumentProperties().getProperty("EMPMASTER_BACKUP");
  Logger.log(backup || "NO BACKUP FOUND");
}
