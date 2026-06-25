// Use this script only for Highlight data.
// It reads and writes the "Highlight information" sheet tab and uploads pictures to the SYSTEM WEBSITE Drive folder.
const HIGHLIGHT_SHEET_NAME = "Highlight information";
const HIGHLIGHT_FOLDER_ID = "1J4M5xA3rZoZc-ZDg8MSKpy8zlD-QDdXU";
const USER_PERMISSION_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzPdDus7JcJsjBNzYpFt9L5fFh1T3GJhOaOH1st-yvXBFXpzJWceRJtqd1X_ePJ7frR7g/exec";
const MAX_IMAGE_BYTES = 5 * 1024 * 1024;
const ALLOWED_IMAGE_TYPES = ["image/jpeg", "image/png", "image/webp"];
const HIGHLIGHT_HEADERS = [
  "ID",
  "Year",
  "Month",
  "Topic",
  "Detail",
  "Image URLs",
  "Created At"
];

function doGet(e) {
  const year = e && e.parameter && e.parameter.year ? String(e.parameter.year) : "";
  const callback = e && e.parameter && e.parameter.callback ? String(e.parameter.callback) : "";
  const rows = readHighlights(year);
  const payload = JSON.stringify({
    ok: true,
    message: "Highlight data API is running",
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
    if (payload.action === "delete") {
      assertPermission(payload, "delete");
      return deleteHighlight(payload.id);
    }
    if (payload.action === "deleteImage") {
      assertPermission(payload, "delete");
      return deleteHighlightImage(payload.id, payload.imageUrl);
    }

    assertPermission(payload, "edit");
    return saveHighlight(payload);
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function deleteHighlightImage(id, imageUrl) {
  const sheet = getHighlightSheet();
  const rowNumber = findHighlightRowNumber(sheet, id);

  if (!rowNumber) {
    return jsonResponse({ ok: false, error: "Highlight not found", id });
  }

  const row = highlightValuesToRow(sheet.getRange(rowNumber, 1, 1, HIGHLIGHT_HEADERS.length).getValues()[0]);
  const targetUrl = normalizeDriveImageUrl(imageUrl);
  row.imageUrls = normalizeImageUrls(row.imageUrls || row.imageUrl)
    .filter((url) => normalizeDriveImageUrl(url) !== targetUrl);
  row.imageUrl = row.imageUrls[0] || "";
  sheet.getRange(rowNumber, 1, 1, HIGHLIGHT_HEADERS.length).setValues([highlightToSheetValues(row)]);
  trashDriveFile(imageUrl);

  return jsonResponse({ ok: true, id, imageUrls: row.imageUrls });
}

function deleteHighlight(id) {
  const sheet = getHighlightSheet();
  const rowNumber = findHighlightRowNumber(sheet, id);

  if (!rowNumber) {
    return jsonResponse({ ok: false, error: "Highlight not found", id });
  }

  sheet.deleteRow(rowNumber);
  return jsonResponse({ ok: true, deleted: id });
}

function saveHighlight(payload) {
  const sheet = getHighlightSheet();
  const id = payload.id || `highlight-${Date.now()}`;
  const existingRowNumber = findHighlightRowNumber(sheet, id);
  const existingRow = existingRowNumber ? highlightValuesToRow(sheet.getRange(existingRowNumber, 1, 1, HIGHLIGHT_HEADERS.length).getValues()[0]) : null;
  const payloadImageUrls = normalizeImageUrls(payload.imageUrls);
  const existingImageUrls = normalizeImageUrls((existingRow && existingRow.imageUrls) || (existingRow && existingRow.imageUrl));
  const hasPayloadImageUrls = payload.imageUrls !== undefined || payload.imageUrl !== undefined;
  const imageUrls = (hasPayloadImageUrls ? payloadImageUrls : existingImageUrls)
    .concat(saveHighlightImages(id, payload.images || payload.image || []));
  const row = {
    id,
    year: String(payload.year || ""),
    month: String(payload.month || ""),
    topic: payload.topic || "",
    detail: payload.detail || "",
    imageUrl: imageUrls[0] || "",
    imageUrls,
    createdAt: existingRow && existingRow.createdAt ? existingRow.createdAt : new Date()
  };

  if (existingRowNumber) {
    sheet.getRange(existingRowNumber, 1, 1, HIGHLIGHT_HEADERS.length).setValues([highlightToSheetValues(row)]);
  } else {
    sheet.appendRow(highlightToSheetValues(row));
  }

  return jsonResponse({ ok: true, highlight: row });
}

function saveHighlightImages(id, images) {
  const list = Array.isArray(images) ? images : [images];
  return list
    .map((image, index) => saveHighlightImage(`${id}-${index + 1}`, image))
    .filter(Boolean);
}

function saveHighlightImage(id, dataUrl) {
  const matches = String(dataUrl).match(/^data:(.+);base64,(.+)$/);
  if (!matches) return "";

  const contentType = matches[1];
  if (!ALLOWED_IMAGE_TYPES.includes(contentType)) return "";

  const extension = contentType.split("/")[1] || "png";
  const bytes = Utilities.base64Decode(matches[2]);
  if (bytes.length > MAX_IMAGE_BYTES) return "";

  const blob = Utilities.newBlob(bytes, contentType, `${id}.${extension}`);
  const file = DriveApp.getFolderById(HIGHLIGHT_FOLDER_ID).createFile(blob);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w1600`;
}

function authorizeDriveAccess() {
  const folder = DriveApp.getFolderById(HIGHLIGHT_FOLDER_ID);
  const file = folder.createFile("drive-authorization-test.txt", "Drive authorization test");

  file.setTrashed(true);
  return folder.getName();
}

function readHighlights(year) {
  const sheet = getHighlightSheet();
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) return [];

  return values.slice(1)
    .filter((row) => row.some((cell) => cell !== ""))
    .map(highlightValuesToRow)
    .filter((row) => !year || String(row.year) === String(year));
}

function highlightValuesToRow(values) {
  const imageUrls = normalizeImageUrls(values[5]);

  return {
    id: values[0],
    year: values[1],
    month: values[2],
    topic: values[3],
    detail: values[4],
    imageUrl: imageUrls[0] || "",
    imageUrls,
    createdAt: values[6]
  };
}

function highlightToSheetValues(row) {
  return [
    safeSheetValue(row.id),
    safeSheetValue(row.year),
    safeSheetValue(row.month),
    safeSheetValue(row.topic),
    safeSheetValue(row.detail),
    imageUrlsText(row.imageUrls || row.imageUrl),
    row.createdAt || new Date()
  ];
}

function findHighlightRowNumber(sheet, id) {
  const values = sheet.getDataRange().getValues();

  for (let index = 1; index < values.length; index += 1) {
    if (String(values[index][0]) === String(id)) return index + 1;
  }

  return 0;
}

function normalizeDriveImageUrl(url) {
  const id = driveFileId(url);
  return id ? `https://drive.google.com/thumbnail?id=${id}&sz=w1600` : String(url || "").trim();
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

function getHighlightSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(HIGHLIGHT_SHEET_NAME) || spreadsheet.insertSheet(HIGHLIGHT_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HIGHLIGHT_HEADERS);
  }

  return sheet;
}
