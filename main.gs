// main.gs

//----------------------------------------------------------------------------------------------------
// DICTIONARIES AND VARIABLES
//----------------------------------------------------------------------------------------------------

// // Main configuration dictionary
const Info = {
  "INFO": {
    "STRING": {
      "B2": "Period number:",
      "B3": "First Day of Period:",
      "B5": "Week range on cover:"
    },
    "COLORS": {
      "splash": [],
      "white": []
    },
    "RIGHT": ["B2", "B3", "B5"],
    "MERGECENTER": [],
    "ALLOWEDCELLS": []
  },

  "COVER": {
    "STRING": {
      "B2": "REQUEST OFF SHEET  ",
      "D4": "This request sheet is for the week of the :    ",
      "F4": "To the :",
      "B5": " ALL REQUEST NEED TO BE IN BY SUNDAY :",
      "B6": "  SCHEDULE WILL BE POSTED ON TUES PENDING CHANGES",
      "B7": "MANAGER",
      "B8": "DATE",
      "B9": "NAME",
      "J7": "UPDATED ON:",
      "K7": "NOTES"
    },
    "FORMULAS": {
    },
    "COLORS": {
      "header": ["B2:I3", "B7:I7"]
    },
    "VARIABLES": {
    },
    "OUTLINEMED": ["B2:I3", "B7:I7", "H4:I6", "B4:G6", "B4:B6"],
    "OUTLINETHIN": [],
    "MERGE": [],
    "RIGHT": [],
    "MERGERIGHT": ["B4:D4", "B5:G5"],
    "MERGELEFT": ["B6:G6"],
    "MERGECENTER": ["B8", "B7:I7", "B9", "B2:I3", "B7:I7", "C8", "D8", "E8", "F8", "G8", "H8", "I8", "E4", "F4", "G4", "H4:I6", "J7", "K7"],
    "ALLOWEDCELLS": []
  },

  "INPUT": {
    "STRING": {
      "C2": "Name:",
      "D2": "Position:",
      "F6": "Notes",
      "F2": "MONDAY",
      "G2": "TUESDAY",
      "H2": "WEDNESDAY",
      "I2": "THURSDAY",
      "J2": "FRIDAY",
      "K2": "SATURDAY",
      "L2": "SUNDAY"
    },
    "FORMULAS": {
    },
    "COLORS": {
      "orange": ["C2:D2", "F2:L2", "F6:L6"],
    },
    "VARIABLES": {
    },
    "OUTLINEMED": ["C2", "D2", "C3", "D3", "F6:L6", "F7:L8"],
    "OUTLINETHIN": [],
    "MERGE": [],
    "RIGHT": [],
    "MERGERIGHT": [],
    "MERGELEFT": [],
    "MERGECENTER": ["F2", "G2", "H2", "I2", "J2", "K2", "L2", "F6:L6", "F7:L8"],
    "ALLOWEDCELLS": ["C3", "D3", "F3", "F7","G3", "H3", "I3", "J3", "K3", "L3"]
  }
};

// Dictionary of color names and corresponding hex codes.
const ColorDict = {
  "header": "#f3f3f3",
  "shift": "#f1cb65",
  "close": "#4285f4",
  "orange": "#FFA500",
  "light blue": "#ADD8E6",
  "red": "#FF0000",
  "yellow": "#FFFF00",
  "purple": "#800080",
  "pink": "#FFC0CB",
  "brown": "#A52A2A",
  "black": "#000000",
  "white": "#ffffff"
};

let EMPMASTER = { MANAGER: {}, INSIDER: {}, DRIVER: {} };
let CURRENT_PERIOD_START = null;
const TempCellBackup = {};

const REQUIRED_SHEETS = ["COVER", "INPUT", "INFO"];

function onOpen() {
  BuildMenus();

  if (!IsFileInitialized()) {
    BuildSetupMenu();
  }

  LoadEmpMaster();
  GetWeekInfo();
  Restore();
  UpdateCover();
}

function BuildMenus() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const isLocked = props.getProperty("docLocked") === "true";

  const security = ui.createMenu("🔒 Security");

  if (isLocked) {
    security.addItem("🔓 Unlock Document", "Unlock");
  } else {
    security.addItem("🔐 Lock Document", "Lock");
  }

  security
    .addSeparator()
    .addItem("🛠 Set/Change Password", "SetupPassword")
    .addItem("❓ Recover Password", "RecoverPassword")
    .addToUi();

  const tools = ui.createMenu("🛠 Script Tools");

  if (!isLocked) {
    tools.addItem("Set New Version", "PromptNewVersion")
         .addSeparator();
  }

  tools
    .addItem("Show Current Version", "ShowVersion")
    .addItem("View Changelog", "ViewChangeLog")
    .addToUi();
}

function LoadEmpMaster() {
  const props = PropertiesService.getDocumentProperties();
  const backup = props.getProperty("EMPMASTER_BACKUP");

  EMPMASTER = backup
    ? JSON.parse(backup)
    : { MANAGER: {}, INSIDER: {}, DRIVER: {} };
}

function IsFileInitialized() {
  const props = PropertiesService.getDocumentProperties();
  return props.getProperty("FILE_INITIALIZED") === "true";
}

function BuildSetupMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⚙️ Setup")
    .addItem("Initialize This File", "InitializeThisFile")
    .addToUi();
}

function InitializeThisFile() {
  SetupPassword();
  ApplySheetProtections();
  InstallChangeTrigger();

  LoadEmpMaster();
  GetWeekInfo();
  Restore();
  UpdateCover();

  const props = PropertiesService.getDocumentProperties();
  props.setProperty("FILE_INITIALIZED", "true");

  SpreadsheetApp.getUi().alert("Setup complete.");
  onOpen();
}

function EnsureRequiredSheetsExist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingNames = ss.getSheets().map(s => s.getName());

  REQUIRED_SHEETS.forEach(name => {
    if (!existingNames.includes(name)) {
      ss.insertSheet(name);
    }
  });
}
