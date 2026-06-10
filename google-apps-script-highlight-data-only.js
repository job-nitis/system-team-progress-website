// Use this script only for Highlight data.
// It reads and writes the "Highlight information" sheet tab and uploads pictures to the SYSTEM WEBSITE Drive folder.
const HIGHLIGHT_SHEET_NAME = "Highlight information";
const HIGHLIGHT_FOLDER_ID = "1J4M5xA3rZoZc-ZDg8MSKpy8zlD-QDdXU";
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
    return ContentService
      .createTextOutput(`${callback}(${payload});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return jsonResponse(JSON.parse(payload));
}

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return jsonResponse({ ok: false, error: "No POST data received" });
  }

  const payload = JSON.parse(e.postData.contents);
  if (payload.action === "delete") return deleteHighlight(payload.id);
  if (payload.action === "deleteImage") return deleteHighlightImage(payload.id, payload.imageUrl);

  return saveHighlight(payload);
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
  const extension = contentType.split("/")[1] || "png";
  const bytes = Utilities.base64Decode(matches[2]);
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
    row.id || "",
    row.year || "",
    row.month || "",
    row.topic || "",
    row.detail || "",
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

function getHighlightSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(HIGHLIGHT_SHEET_NAME) || spreadsheet.insertSheet(HIGHLIGHT_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HIGHLIGHT_HEADERS);
  }

  return sheet;
}
