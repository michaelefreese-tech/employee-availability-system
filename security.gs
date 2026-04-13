// security.gs


function ApplySheetProtections() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();

  const input = ss.getSheetByName("INPUT");
  const cover = ss.getSheetByName("COVER");
  const info = ss.getSheetByName("INFO");

  const me = Session.getEffectiveUser().getEmail();

  // Remove old sheet protections we created before reapplying
  ss.getSheets().forEach(sheet => {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(protection => {
      try {
        protection.remove();
      } catch (err) {
        // ignore protections we cannot remove
      }
    });
  });

  // COVER = fully locked
  const coverProtection = cover.protect().setDescription("LOCKED_COVER");
  coverProtection.setWarningOnly(false);
  try {
    coverProtection.addEditor(me);
    coverProtection.removeEditors(
      coverProtection.getEditors().filter(user => user.getEmail() !== me)
    );
    if (coverProtection.canDomainEdit()) coverProtection.setDomainEdit(false);
  } catch (err) {}

  // INFO = fully locked
  const infoProtection = info.protect().setDescription("LOCKED_INFO");
  infoProtection.setWarningOnly(false);
  try {
    infoProtection.addEditor(me);
    infoProtection.removeEditors(
      infoProtection.getEditors().filter(user => user.getEmail() !== me)
    );
    if (infoProtection.canDomainEdit()) infoProtection.setDomainEdit(false);
  } catch (err) {}

  // INPUT = locked except allowed entry cells
  const inputProtection = input.protect().setDescription("LOCKED_INPUT");
  inputProtection.setWarningOnly(false);

  const unprotected = [
    input.getRange("C3:D3"),
    input.getRange("F3:L3")
  ];

  inputProtection.setUnprotectedRanges(unprotected);

  try {
    inputProtection.addEditor(me);
    inputProtection.removeEditors(
      inputProtection.getEditors().filter(user => user.getEmail() !== me)
    );
    if (inputProtection.canDomainEdit()) inputProtection.setDomainEdit(false);
  } catch (err) {}

  props.setProperty("docLocked", "true");
}

function RemoveSheetProtections() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();

  ss.getSheets().forEach(sheet => {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(protection => {
      try {
        protection.remove();
      } catch (err) {
        // ignore protections we cannot remove
      }
    });
  });

  props.setProperty("docLocked", "false");
}

// 🔒 Lock the document
function Lock() {
  const props = PropertiesService.getDocumentProperties();
  const status = props.getProperty("docLocked");

  if (status === "true") {
    SpreadsheetApp.getUi().alert("Document is already locked.");
    return;
  }

  ApplySheetProtections();
  SpreadsheetApp.getUi().alert("Document is now locked.");
  onOpen();
}

// 🔓 Unlock with password prompt
function Unlock() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const status = props.getProperty("docLocked");

  if (status === "false") {
    ui.alert("Document is already unlocked.");
    return;
  }

  const response = ui.prompt("Enter password to unlock:");
  const password = response.getResponseText().trim();
  const stored = JSON.parse(props.getProperty("PASSWORD_DATA") || "null");

  if (!stored) {
    ui.alert("No password set. Please run SetupPassword() first.");
    return;
  }

  if (password === stored.password) {
    RemoveSheetProtections();
    props.setProperty("justUnlocked", "true");
    ui.alert("Document is now unlocked.");
    onOpen();
  } else {
    const recovery = ui.prompt("Incorrect password.\nType 'recover' to reset using security questions, or Cancel.");
    if (recovery.getResponseText().toLowerCase() === "recover") {
      RecoverPassword();
    } else {
      ui.alert("Access denied.");
    }
  }
}

// 🛠 Setup password and security questions
function SetupPassword() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const stored = JSON.parse(props.getProperty("PASSWORD_DATA") || "null");

  if (stored) {
    const current = ui.prompt("Enter current password to change:");
    if (current.getSelectedButton() !== ui.Button.OK || current.getResponseText().trim() !== stored.password) {
      ui.alert("Incorrect password. Cannot change settings.");
      return;
    }
  }

  const password = ui.prompt("Set a new password:").getResponseText().trim();
  const q1 = ui.prompt("Security Question 1 (e.g., Favorite color?)").getResponseText().trim();
  const a1 = ui.prompt("Answer to Question 1:").getResponseText().trim().toLowerCase();
  const q2 = ui.prompt("Security Question 2 (e.g., First pet's name?)").getResponseText().trim();
  const a2 = ui.prompt("Answer to Question 2:").getResponseText().trim().toLowerCase();

  const data = { password, q1, a1, q2, a2 };
  props.setProperty("PASSWORD_DATA", JSON.stringify(data));
  ui.alert("Password and security questions set.");
}

// 🛠 Recover password via Q&A
function RecoverPassword() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const stored = JSON.parse(props.getProperty("PASSWORD_DATA") || "null");

  if (!stored) {
    ui.alert("No recovery data found.");
    return;
  }

  const a1 = ui.prompt(stored.q1).getResponseText().trim().toLowerCase();
  const a2 = ui.prompt(stored.q2).getResponseText().trim().toLowerCase();

  if (a1 === stored.a1 && a2 === stored.a2) {
    const newPass = ui.prompt("Enter your new password:").getResponseText().trim();
    stored.password = newPass;
    props.setProperty("PASSWORD_DATA", JSON.stringify(stored));
    ui.alert("Password successfully reset.");
  } else {
    ui.alert("Security answers did not match. Cannot reset.");
  }
}

// 🚫 Admin tool: Clear all password and recovery data
function ClearPassword() {
  PropertiesService.getDocumentProperties().deleteProperty("PASSWORD_DATA");
  SpreadsheetApp.getUi().alert("All password data cleared.");
}

// 📌 Show current script version
function ShowVersion() {
  const version = PropertiesService.getDocumentProperties().getProperty("SCRIPT_VERSION") || "Not set";
  SpreadsheetApp.getUi().alert("Current Version: " + version);
}

// 🛠 Manually add version + changelog
function PromptNewVersion() {
  const ui = SpreadsheetApp.getUi();
  const versionPrompt = ui.prompt("Enter new version:");
  if (versionPrompt.getSelectedButton() !== ui.Button.OK) return;

  const newVersion = versionPrompt.getResponseText().trim();
  const logPrompt = ui.prompt("Enter changelog for version " + newVersion + ":");
  if (logPrompt.getSelectedButton() !== ui.Button.OK) return;

  const changelog = logPrompt.getResponseText();
  UpdateVersionWithLog(newVersion, changelog);
}

// 📚 Store changelog data
function UpdateVersionWithLog(version, changelog) {
  const props = PropertiesService.getDocumentProperties();
  const rawLog = props.getProperty("SCRIPT_CHANGELOG_DICT") || "[]";
  const log = JSON.parse(rawLog);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM.dd.yyyy");

  log.push({ date: timestamp, version, log: changelog });

  props.setProperty("SCRIPT_CHANGELOG_DICT", JSON.stringify(log));
  props.setProperty("SCRIPT_VERSION", version);
}

// 📜 Display changelog history
function ViewChangeLog() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const rawLog = props.getProperty("SCRIPT_CHANGELOG_DICT") || "[]";
  const logList = JSON.parse(rawLog);

  if (logList.length === 0) {
    ui.alert("No changelog available.");
    return;
  }

  let message = "Version History:\n\n";
  logList.forEach(entry => {
    message += `${entry.date} : ${entry.version} : ${entry.log}\n`;
  });

  ui.alert(message);
}

// 🧹 Admin: Clear version history
function ClearVersionHistory() {
  const ui = SpreadsheetApp.getUi();
  const confirmation = ui.alert("⚠️ WARNING", "This will permanently delete version history. Proceed?", ui.ButtonSet.YES_NO);
  if (confirmation !== ui.Button.YES) return;

  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty("SCRIPT_VERSION");
  props.deleteProperty("SCRIPT_CHANGELOG_DICT");

  ui.alert("✅ All version history cleared.");
}
