// restore.gs

// 🧼 Full sheet structure restoration using Info dictionary
function Restore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const SheetInfo = Info[sheetName];
    if (!SheetInfo) return;

    // ↔ Merging
    if (SheetInfo.MERGE) {
      SheetInfo.MERGE.forEach(rangeStr => {
        sheet.getRange(rangeStr).merge();
      });
    }

    if (SheetInfo.MERGERIGHT) {
      SheetInfo.MERGERIGHT.forEach(rangeStr => {
        sheet.getRange(rangeStr).merge();
      });
    }

    if (SheetInfo.MERGELEFT) {
      SheetInfo.MERGELEFT.forEach(rangeStr => {
        sheet.getRange(rangeStr).merge();
      });
    }

    if (SheetInfo.MERGECENTER) {
      SheetInfo.MERGECENTER.forEach(rangeStr => {
        sheet.getRange(rangeStr).merge();
      });
    }

    // ↔ Alignment
    if (SheetInfo.MERGE) {
      SheetInfo.MERGE.forEach(rangeStr => {
        sheet.getRange(rangeStr).setHorizontalAlignment("left").setVerticalAlignment("middle");
      });
    }

    if (SheetInfo.RIGHT) {
      SheetInfo.RIGHT.forEach(rangeStr => {
        sheet.getRange(rangeStr).setHorizontalAlignment("right").setVerticalAlignment("middle");
      });
    }

    if (SheetInfo.MERGERIGHT) {
      SheetInfo.MERGERIGHT.forEach(rangeStr => {
        sheet.getRange(rangeStr).setHorizontalAlignment("right").setVerticalAlignment("middle");
      });
    }

    if (SheetInfo.MERGELEFT) {
      SheetInfo.MERGELEFT.forEach(rangeStr => {
        sheet.getRange(rangeStr).setHorizontalAlignment("left").setVerticalAlignment("middle");
      });
    }

    if (SheetInfo.MERGECENTER) {
      SheetInfo.MERGECENTER.forEach(rangeStr => {
        sheet.getRange(rangeStr).setHorizontalAlignment("center").setVerticalAlignment("middle");
      });
    }

    // 🎨 Background Colors
    if (SheetInfo.COLORS) {
      for (let colorKey in SheetInfo.COLORS) {
        const hex = ColorDict[colorKey.toLowerCase()];
        SheetInfo.COLORS[colorKey].forEach(rangeStr => {
          sheet.getRange(rangeStr).setBackground(hex);
        });
      }
    }

    // 🔲 Borders
    if (SheetInfo.OUTLINETHIN) {
      SheetInfo.OUTLINETHIN.forEach(rangeStr => {
        sheet.getRange(rangeStr).setBorder(
          true, true, true, true, false, false,
          "black",
          SpreadsheetApp.BorderStyle.SOLID
        );
      });
    }

    if (SheetInfo.OUTLINEMED) {
      SheetInfo.OUTLINEMED.forEach(rangeStr => {
        sheet.getRange(rangeStr).setBorder(
          true, true, true, true, false, false,
          "black",
          SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        );
      });
    }

    // 📝 Strings
    if (SheetInfo.STRING) {
      for (let cell in SheetInfo.STRING) {
        const range = sheet.getRange(cell);
        range
          .setValue(SheetInfo.STRING[cell])
          .setFontFamily("Arial")
          .setFontSize(10);
      }
    }

    // 🧮 Formulas
    if (SheetInfo.FORMULAS) {
      for (let cell in SheetInfo.FORMULAS) {
        sheet.getRange(cell).setFormula(SheetInfo.FORMULAS[cell]);
      }
    }



    if (sheetName === "COVER") {
      sheet.getRange("B2").setFontSize(31);
      sheet.getRange("B5").setFontSize(22);
      sheet.getRange("B6").setFontSize(15);
      sheet.getRange("B7").setFontSize(18);
      sheet.getRange("H4").setFontSize(31);
      sheet.getRange("D4").setFontSize(12);
      sheet.getRange("B7").setFontWeight("bold");
    }
  });
}
