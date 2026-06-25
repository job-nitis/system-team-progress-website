// Use this script only for License Time.
// It reads and writes the "License detail" sheet tab in the License Time spreadsheet.
// If this Apps Script is not created from Extensions > Apps Script inside the spreadsheet,
// paste the spreadsheet ID between the quotes below.
const LICENSE_SPREADSHEET_ID = "";
const LICENSE_SHEET_NAME = "License detail";
const USER_PERMISSION_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzPdDus7JcJsjBNzYpFt9L5fFh1T3GJhOaOH1st-yvXBFXpzJWceRJtqd1X_ePJ7frR7g/exec";
const LICENSE_HEADERS = [
  "ID",
  "License Name",
  "Expired Date",
  "Remaining Time",
  "Updated At"
];
const LICENSE_TIME_ZONE = "Asia/Bangkok";

function doGet(e) {
  try {
    const callback = e && e.parameter && e.parameter.callback ? String(e.parameter.callback) : "";
    const rows = readLicenseRows();
    return callbackResponse(callback, {
      ok: true,
      message: "License Time API is running",
      rows
    });
  } catch (error) {
    const callback = e && e.parameter && e.parameter.callback ? String(e.parameter.callback) : "";
    return callbackResponse(callback, {
      ok: false,
      error: String(error && error.message ? error.message : error)
    });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok: false, error: "No POST data received" });
    }

    const payload = JSON.parse(e.postData.contents);
    assertPermission(payload, "edit");
    const rows = Array.isArray(payload.rows) ? payload.rows.map(normalizeLicenseRow) : [];
    const sheet = getLicenseSheet();
    const existingRows = readLicenseRows();
    if (hasDeletedRows(existingRows, rows)) assertPermission(payload, "delete");

    sheet.clearContents();
    sheet.appendRow(LICENSE_HEADERS);
    rows.forEach((row) => sheet.appendRow(licenseToSheetValues(row)));
    formatLicenseSheet(sheet);

    return jsonResponse({ ok: true, rows: rows.length });
  } catch (error) {
    return jsonResponse({
      ok: false,
      error: String(error && error.message ? error.message : error)
    });
  }
}

function readLicenseRows() {
  const sheet = getLicenseSheet();
  formatLicenseSheet(sheet);
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) return [];

  return values.slice(1)
    .filter((row) => row.some((cell) => cell !== ""))
    .map(licenseValuesToRow);
}

function licenseValuesToRow(values) {
  return normalizeLicenseRow({
    id: values[0],
    name: values[1],
    expiredDate: values[2],
    updatedAt: values.length >= 5 ? values[4] : values[3]
  });
}

function licenseToSheetValues(row) {
  const days = daysUntil(row.expiredDate);

  return [
    safeSheetValue(row.id),
    safeSheetValue(row.name),
    safeSheetValue(row.expiredDate),
    remainingTimeLabel(days),
    row.updatedAt || new Date()
  ];
}

function hasDeletedRows(existingRows, incomingRows) {
  const incomingIds = new Set(incomingRows.map((row) => String(row && row.id ? row.id : "")));
  return existingRows.some((row) => row.id && !incomingIds.has(String(row.id)));
}

function normalizeLicenseRow(row) {
  return {
    id: String(row && row.id ? row.id : `license-${Date.now()}`),
    name: String((row && (row.name || row.licenseName)) || "").trim(),
    expiredDate: formatDate(row && (row.expiredDate || row.expiryDate)),
    updatedAt: row && row.updatedAt ? row.updatedAt : new Date()
  };
}

function formatDate(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;

  const date = new Date(text);
  if (Number.isNaN(date.getTime())) return text;

  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function daysUntil(dateValue) {
  const inputDate = formatDate(dateValue);
  if (!inputDate) return null;

  const expired = dateOnly(inputDate);
  const today = dateOnly(Utilities.formatDate(new Date(), LICENSE_TIME_ZONE, "yyyy-MM-dd"));

  return Math.ceil((expired.getTime() - today.getTime()) / 86400000);
}

function dateOnly(dateText) {
  const parts = String(dateText || "").split("-").map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function remainingTimeLabel(days) {
  if (days === null) return "";
  if (days < 0) return `Expired ${Math.abs(days)} days`;
  if (days === 0) return "Expires today";
  return `${days} days`;
}

function remainingColor(days) {
  if (days === null || days < 30) return { background: "#dc2626", font: "#ffffff" };
  if (days < 60) return { background: "#facc15", font: "#1f2937" };
  return { background: "#16a34a", font: "#ffffff" };
}

function formatLicenseSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = LICENSE_HEADERS.length;

  sheet.getRange(1, 1, 1, lastColumn)
    .setValues([LICENSE_HEADERS])
    .setFontWeight("bold")
    .setBackground("#e8f0fe");
  sheet.setFrozenRows(1);

  if (lastRow <= 1) return;

  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const remainingValues = [];
  const backgrounds = [];
  const fonts = [];

  values.forEach((row) => {
    const days = daysUntil(row[2]);
    const color = remainingColor(days);
    remainingValues.push([remainingTimeLabel(days)]);
    backgrounds.push([color.background]);
    fonts.push([color.font]);
  });

  sheet.getRange(2, 4, remainingValues.length, 1)
    .setValues(remainingValues)
    .setBackgrounds(backgrounds)
    .setFontColors(fonts)
    .setFontWeight("bold");
  sheet.autoResizeColumns(1, lastColumn);
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function assertPermission(payload, action) {
  const user = lookupUser(payload && payload.actorEmail);
  const role = user && user.role;

  if (action === "delete" && role !== "Admin") {
    throw new Error("Admin permission required");
  }

  if (action === "edit" && !["Admin", "Editor"].includes(role)) {
    throw new Error("Editor permission required");
  }
}

function lookupUser(email) {
  const cleanEmail = String(email || "").trim().toLowerCase();
  if (!cleanEmail) throw new Error("Login email is required");
  if (!USER_PERMISSION_WEB_APP_URL) throw new Error("User permission URL is not set");

  const response = UrlFetchApp.fetch(`${USER_PERMISSION_WEB_APP_URL}?email=${encodeURIComponent(cleanEmail)}`, {
    muteHttpExceptions: true
  });
  const payload = JSON.parse(response.getContentText() || "{}");
  if (!payload.ok || !payload.user || !payload.user.active) {
    throw new Error("User is not approved or inactive");
  }

  return payload.user;
}

function safeSheetValue(value) {
  const text = String(value || "").slice(0, 5000);
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}

function callbackResponse(callback, payload) {
  const safeCallback = String(callback || "");
  if (!safeCallback) return jsonResponse(payload);

  if (!/^[A-Za-z_$][\w$]*(\.[A-Za-z_$][\w$]*)*$/.test(safeCallback)) {
    return jsonResponse({ ok: false, error: "Invalid callback" });
  }

  return ContentService
    .createTextOutput(`${safeCallback}(${JSON.stringify(payload)});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getLicenseSheet() {
  const spreadsheet = LICENSE_SPREADSHEET_ID
    ? SpreadsheetApp.openById(LICENSE_SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error("No active spreadsheet. Create this script from Extensions > Apps Script in the License Time spreadsheet, or set LICENSE_SPREADSHEET_ID.");
  }

  const sheet = spreadsheet.getSheetByName(LICENSE_SHEET_NAME) || spreadsheet.insertSheet(LICENSE_SHEET_NAME);

  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === "") {
    sheet.getRange(1, 1, 1, LICENSE_HEADERS.length).setValues([LICENSE_HEADERS]);
  }

  return sheet;
}
