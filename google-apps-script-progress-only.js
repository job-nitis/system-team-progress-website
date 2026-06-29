// Use this script only for the System Team Progress table.
// It reads and writes the ProgressRows sheet tab and uploads project pictures to the SYSTEM WEBSITE Drive folder.
const SHEET_NAME = "ProgressRows";
const PROJECT_IMAGE_FOLDER_ID = "1J4M5xA3rZoZc-ZDg8MSKpy8zlD-QDdXU";
const USER_PERMISSION_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzPdDus7JcJsjBNzYpFt9L5fFh1T3GJhOaOH1st-yvXBFXpzJWceRJtqd1X_ePJ7frR7g/exec";
const MAX_IMAGE_BYTES = 5 * 1024 * 1024;
const ALLOWED_IMAGE_TYPES = ["image/jpeg", "image/png", "image/webp"];
const HEADERS = [
  "Year",
  "ID",
  "Initiative Project",
  "Owner",
  "Tech Preparation",
  "TOR",
  "PR",
  "SAP PR",
  "PO",
  "Delivery",
  "FAT/SAT",
  "Close MOC",
  "Overall Plan",
  "Project Image URLs",
  "Plan",
  "Actual",
  "Updated At"
];

function authorizeProgressPermissions() {
  const sheet = getSheet();
  sheet.getName();
  const folder = DriveApp.getFolderById(PROJECT_IMAGE_FOLDER_ID);
  folder.getName();
  UrlFetchApp.fetch(`${USER_PERMISSION_WEB_APP_URL}?email=authorization-check@example.com`, {
    muteHttpExceptions: true
  });

  return "Progress permissions are authorized";
}

function doGet(e) {
  const year = e && e.parameter && e.parameter.year ? String(e.parameter.year) : "";
  const callback = e && e.parameter && e.parameter.callback ? String(e.parameter.callback) : "";
  const rows = readRows(year);
  const payload = JSON.stringify({
    ok: true,
    message: "System Team Progress API is running",
    year,
    rows
  });

  if (callback) {
    return callbackResponse(callback, JSON.parse(payload));
  }

  return jsonResponse(JSON.parse(payload));
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok: false, error: "No POST data received" });
    }

    const payload = JSON.parse(e.postData.contents);
    if (payload.action === "deleteImage") {
      assertPermission(payload, "delete");
      return deleteProjectImage(payload.id, payload.imageUrl);
    }

    assertPermission(payload, "edit");
    const year = String(payload.year || "");
    const existingRows = readRows("");
    const incomingRawRows = Array.isArray(payload.rows) ? payload.rows : [];
    if (hasDeletedRows(existingRows, incomingRawRows, year)) assertPermission(payload, "delete");

    const incomingRows = incomingRawRows.map((row) => prepareIncomingRow(row, year, existingRows));
    const otherRows = existingRows.filter((row) => String(row.year) !== year);
    const allRows = otherRows.concat(incomingRows);
    const sheet = getSheet();

    sheet.clearContents();
    sheet.appendRow(HEADERS);
    allRows.forEach((row) => sheet.appendRow(rowToSheetValues(row)));

    return jsonResponse({ ok: true, rows: incomingRows.length });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function hasDeletedRows(existingRows, incomingRows, year) {
  const incomingIds = new Set(incomingRows.map((row) => String(row && row.id ? row.id : "")));
  return existingRows
    .filter((row) => String(row.year) === String(year))
    .some((row) => row.id && !incomingIds.has(String(row.id)));
}

function deleteProjectImage(id, imageUrl) {
  const sheet = getSheet();
  const rowNumber = findProjectRowNumber(sheet, id);

  if (!rowNumber) {
    return jsonResponse({ ok: false, error: "Project not found", id });
  }

  const row = sheetValuesToRow(sheet.getRange(rowNumber, 1, 1, HEADERS.length).getValues()[0]);
  const targetUrl = normalizeDriveImageUrl(imageUrl);
  row.projectImageUrls = normalizeImageUrls(row.projectImageUrls || row.projectImageUrl)
    .filter((url) => normalizeDriveImageUrl(url) !== targetUrl);
  row.projectImageUrl = row.projectImageUrls[0] || "";
  sheet.getRange(rowNumber, 1, 1, HEADERS.length).setValues([rowToSheetValues(row)]);
  trashDriveFile(imageUrl);

  return jsonResponse({ ok: true, id, imageUrls: row.projectImageUrls });
}

function prepareIncomingRow(row, year, existingRows) {
  const newImageUrls = saveProjectImages(row.id || `project-${Date.now()}`, row.projectImages || row.projectImage || []);
  const existingRow = (existingRows || []).find((item) => String(item.id) === String(row.id));
  const incomingImageUrls = normalizeImageUrls(row.projectImageUrls || row.projectImageUrl);
  const existingImageUrls = normalizeImageUrls((existingRow && existingRow.projectImageUrls) || (existingRow && existingRow.projectImageUrl));
  const hasIncomingImageUrls = row.projectImageUrls !== undefined || row.projectImageUrl !== undefined;
  const imageUrls = (hasIncomingImageUrls ? incomingImageUrls : existingImageUrls).concat(newImageUrls);

  return {
    ...row,
    year,
    projectImageUrl: imageUrls[0] || "",
    projectImageUrls: imageUrls,
    updatePlan: normalizeMilestones(row.updatePlan),
    updateActual: normalizeMilestones(row.updateActual)
  };
}

function saveProjectImages(id, images) {
  const list = Array.isArray(images) ? images : [images];
  return list
    .map((image, index) => saveProjectImage(`${id}-${index + 1}`, image))
    .filter(Boolean);
}

function saveProjectImage(id, dataUrl) {
  const matches = String(dataUrl).match(/^data:(.+);base64,(.+)$/);
  if (!matches) return "";

  const contentType = matches[1];
  if (!ALLOWED_IMAGE_TYPES.includes(contentType)) return "";

  const extension = contentType.split("/")[1] || "png";
  const bytes = Utilities.base64Decode(matches[2]);
  if (bytes.length > MAX_IMAGE_BYTES) return "";

  const blob = Utilities.newBlob(bytes, contentType, `${id}.${extension}`);
  const file = DriveApp.getFolderById(PROJECT_IMAGE_FOLDER_ID).createFile(blob);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w2400`;
}

function authorizeDriveAccess() {
  const folder = DriveApp.getFolderById(PROJECT_IMAGE_FOLDER_ID);
  const file = folder.createFile("drive-authorization-test.txt", "Drive authorization test");

  file.setTrashed(true);
  return folder.getName();
}

function readRows(year) {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) return [];

  return values.slice(1)
    .filter((row) => row.some((cell) => cell !== ""))
    .map(sheetValuesToRow)
    .filter((row) => !year || String(row.year) === String(year));
}

function sheetValuesToRow(values) {
  return {
    year: values[0],
    id: values[1],
    project: values[2],
    owner: values[3],
    techPrep: values[4],
    tor: values[5],
    pr: values[6],
    sapPr: values[7],
    po: values[8],
    delivery: values[9],
    fatSat: values[10],
    closeMoc: values[11],
    overallPlan: values[12],
    projectImageUrl: values[13],
    projectImageUrls: normalizeImageUrls(values[13]),
    updatePlan: parseMilestoneText(values[14]),
    updateActual: parseMilestoneText(values[15])
  };
}

function rowToSheetValues(row) {
  return [
    safeSheetValue(row.year),
    safeSheetValue(row.id),
    safeSheetValue(row.project),
    safeSheetValue(row.owner),
    safeSheetValue(row.techPrep),
    safeSheetValue(row.tor),
    safeSheetValue(row.pr),
    safeSheetValue(row.sapPr),
    safeSheetValue(row.po),
    safeSheetValue(row.delivery),
    safeSheetValue(row.fatSat),
    safeSheetValue(row.closeMoc),
    safeSheetValue(row.overallPlan),
    imageUrlsText(row.projectImageUrls || row.projectImageUrl),
    milestoneText(row.updatePlan),
    milestoneText(row.updateActual),
    new Date()
  ];
}

function findProjectRowNumber(sheet, id) {
  const values = sheet.getDataRange().getValues();

  for (let index = 1; index < values.length; index += 1) {
    if (String(values[index][1]) === String(id)) return index + 1;
  }

  return 0;
}

function normalizeDriveImageUrl(url) {
  const id = driveFileId(url);
  return id ? `https://drive.google.com/thumbnail?id=${id}&sz=w2400` : String(url || "").trim();
}

function driveFileId(url) {
  const text = String(url || "");
  const match = text.match(/[?&]id=([^&]+)/) || text.match(/\/d\/([^/]+)/);
  return match ? decodeURIComponent(match[1]) : "";
}

function trashDriveFile(url) {
  const id = driveFileId(url);
  if (!id) return;

  try {
    DriveApp.getFileById(id).setTrashed(true);
  } catch (error) {
    // The sheet row is still updated when the file is already gone or inaccessible.
  }
}

function normalizeImageUrls(value) {
  if (Array.isArray(value)) {
    return value.map((item) => String(item || "").trim()).filter(Boolean);
  }

  return String(value || "")
    .split(/\n|,/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function imageUrlsText(value) {
  return normalizeImageUrls(value).join("\n");
}

function normalizeMilestones(value) {
  if (typeof value === "string") return parseMilestoneText(value);
  if (!Array.isArray(value)) return [];

  return value
    .map((item) => ({
      task: String(item && item.task ? item.task : "").trim(),
      months: normalizeMonths(item && item.months ? item.months : [])
    }))
    .filter((item) => item.task);
}

function normalizeMonths(months) {
  const source = Array.isArray(months) ? months : String(months || "").split(",");
  return Array.from(new Set(
    source
      .map(parseMonthValue)
      .filter((month) => month >= 1 && month <= 12)
  )).sort((a, b) => a - b);
}

function parseMonthValue(value) {
  const text = String(value || "").trim();
  if (!text || text === "-") return 0;

  const number = Number(text);
  if (!Number.isNaN(number)) return number;

  const isoMatch = text.match(/\b\d{4}[-/](\d{1,2})\b/);
  if (isoMatch) return Number(isoMatch[1]);

  const monthNames = {
    jan: 1,
    january: 1,
    feb: 2,
    february: 2,
    mar: 3,
    march: 3,
    apr: 4,
    april: 4,
    may: 5,
    jun: 6,
    june: 6,
    jul: 7,
    july: 7,
    aug: 8,
    august: 8,
    sep: 9,
    sept: 9,
    september: 9,
    oct: 10,
    october: 10,
    nov: 11,
    november: 11,
    dec: 12,
    december: 12
  };
  const nameMatch = text.toLowerCase().match(/[a-z]+/);
  return nameMatch ? monthNames[nameMatch[0]] || 0 : 0;
}

function parseMilestoneText(text) {
  return String(text || "")
    .split("\n")
    .map((line) => {
      const parts = line.split("|");
      return {
        task: String(parts[0] || "").trim(),
        months: normalizeMonths(parts[1] || "")
      };
    })
    .filter((item) => item.task);
}

function milestoneText(milestones) {
  return normalizeMilestones(milestones)
    .map((item) => `${item.task} | ${item.months.join(",")}`)
    .join("\n");
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function callbackResponse(callback, payload) {
  const safeCallback = String(callback || "");
  if (!/^[A-Za-z_$][\w$]*(\.[A-Za-z_$][\w$]*)*$/.test(safeCallback)) {
    return jsonResponse({ ok: false, error: "Invalid callback" });
  }

  return ContentService
    .createTextOutput(`${safeCallback}(${JSON.stringify(payload)});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
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

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

  return sheet;
}
