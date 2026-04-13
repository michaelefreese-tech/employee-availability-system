// cover.gs

// 🧼 Clear COVER sheet before redrawing
function ClearCover() {
  const cover = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("COVER");
  const lastRow = Math.max(cover.getLastRow(), 10);
  const range = cover.getRange("B10:K" + lastRow);

  range.clearContent();
  range.clearFormat();
  range.setBorder(false, false, false, false, false, false);
}

// 🖼 Render EMPMASTER into COVER sheet
function UpdateCover() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cover = ss.getSheetByName("COVER");

  ClearCover();

  const nameFontSize = 10;
  const headerFontSize = 18;
  let row = 10;

  const positions = ["MANAGER", "INSIDER", "DRIVER"];
  positions.forEach(position => {
    if (position !== "MANAGER") {
      const header = cover.getRange(row, 2, 1, 8);
      header.merge();
      header
        .setValue(position)
        .setFontSize(headerFontSize)
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setBackground(ColorDict["header"]);
      header.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      row++;
    }

    const employees = Object.keys(EMPMASTER[position]).sort((a, b) =>
      a.split(" ").pop().localeCompare(b.split(" ").pop())
    );
    const startRow = row;

    employees.forEach(name => {
      for (let col = 2; col <= 9; col++) {
        cover.getRange(row, col)
          .setFontSize(nameFontSize)
          .setFontWeight("normal")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
      }

      cover.getRange(row, 2).setValue(name);

      const availability = EMPMASTER[position][name];
      if (Array.isArray(availability)) {
        for (let i = 0; i < 7; i++) {
          cover.getRange(row, 3 + i).setValue(availability[i]);
        }

        cover.getRange(row, 10).setValue(availability[7] || "");
        cover.getRange(row, 11).setValue(availability[8] || "");
      }

      row++;
    });

    if (employees.length > 0) {
      for (let r = startRow; r < row; r++) {
        for (let c = 2; c <= 9; c++) {
          cover.getRange(r, c)
            .setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
        }
      }
    }

    if (employees.length > 0 && position !== "DRIVER") {
      row++;
    }
  });

  // 🗓 Weekday headers
  const days = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"];
  days.forEach((day, i) => cover.getRange(9, 3 + i).setValue(day));

  // 🗓 Dates above days (row 8)
  const rawStartDate = cover.getRange("E4").getValue();
  if (rawStartDate instanceof Date && !isNaN(rawStartDate)) {
    const start = new Date(rawStartDate);
    for (let i = 0; i < 7; i++) {
      const d = new Date(start);
      d.setDate(start.getDate() + i);
      cover.getRange(8, 3 + i).setValue(Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd/yyyy"));
    }
  }
}
