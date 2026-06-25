// Use this script only for System Website User permissions.
// It reads the "Users" sheet tab and returns one active user's role by email.
const USERS_SHEET_NAME = "Users";
const USERS_HEADERS = [
  "Email",
  "Name",
  "Role",
  "Active"
];

function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback ? String(e.parameter.callback) : "";
  const email = e && e.parameter && e.parameter.email ? String(e.parameter.email).trim().toLowerCase() : "";

  try {
    if (!email) {
      return callbackResponse(callback, { ok: false, error: "Email is required" });
    }

    const user = findUserByEmail(email);
    if (!user) {
      return callbackResponse(callback, { ok: false, error: "Email is not approved" });
    }

    if (!user.active) {
      return callbackResponse(callback, { ok: false, error: "User is inactive" });
    }

    return callbackResponse(callback, { ok: true, user });
  } catch (error) {
    return callbackResponse(callback, {
      ok: false,
      error: String(error && error.message ? error.message : error)
    });
  }
}

function findUserByEmail(email) {
  const sheet = getUsersSheet();
  const values = sheet.getDataRange().getValues();

  for (let index = 1; index < values.length; index += 1) {
    const row = userValuesToRow(values[index]);
    if (row.email === email) return row;
  }

  return null;
}

function userValuesToRow(values) {
  const role = String(values[2] || "Viewer").trim();

  return {
    email: String(values[0] || "").trim().toLowerCase(),
    name: String(values[1] || "").trim(),
    role: ["Admin", "Editor", "Viewer"].includes(role) ? role : "Viewer",
    active: values[3] === true || String(values[3]).toUpperCase() === "TRUE"
  };
}

function callbackResponse(callback, payload) {
  const safeCallback = String(callback || "");

  if (safeCallback) {
    if (!/^[A-Za-z_$][\w$]*(\.[A-Za-z_$][\w$]*)*$/.test(safeCallback)) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: "Invalid callback" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(`${safeCallback}(${JSON.stringify(payload)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getUsersSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(USERS_SHEET_NAME) || spreadsheet.insertSheet(USERS_SHEET_NAME);

  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === "") {
    sheet.getRange(1, 1, 1, USERS_HEADERS.length).setValues([USERS_HEADERS]);
  }

  return sheet;
}
