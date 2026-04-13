// validation.gs

// ✅ Main validation controller
function ValidateEntry(type) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("INPUT");

  // Reset colors
  sheet.getRange("C3:D3").setBackground(ColorDict["white"]);
  sheet.getRange("F3:L3").setBackground("white");

  // Add input cell borders
  for (let row = 2; row <= 3; row++) {
    for (let col = 3; col <= 4; col++) {
      sheet.getRange(row, col).setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }

  for (let row = 2; row <= 3; row++) {
    for (let col = 6; col <= 12; col++) {
      sheet.getRange(row, col).setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }

  const employeeOK = ValidateCellRange(sheet, 3, 3, 'employee info');
  const availabilityOK = type !== "remove" ? ValidateCellRange(sheet, 6, 3, 'availability') : true;

  return employeeOK && availabilityOK;
}

// ✅ Validates a block of cells
function ValidateCellRange(sheet, startColumn, startRow, rangeType) {
  let labels = [];

  if (rangeType === "employee info") {
    labels = ["Employee Name", "Employee Position"];
  } else if (rangeType === "availability") {
    labels = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"];
  }

  for (let i = 0; i < labels.length; i++) {
    const cell = sheet.getRange(startRow, startColumn + i);
    if (cell.isBlank()) {
      SpreadsheetApp.getUi().alert("⛔ Missing value: " + labels[i]);
      cell.setBackground("red");
      return false;
    }
    cell.setBackground(ColorDict["white"]);
  }

  return true;
}
