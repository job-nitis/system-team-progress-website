// Use this script only for the System Team Progress table.
// It reads and writes the ProgressRows sheet tab and uploads project pictures to the SYSTEM WEBSITE Drive folder.
const SHEET_NAME = "ProgressRows";
const PROJECT_IMAGE_FOLDER_ID = "1J4M5xA3rZoZc-ZDg8MSKpy8zlD-QDdXU";
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
  const year = String(payload.year || "");
  const existingRows = readRows("");
  const incomingRows = (payload.rows || []).map((row) => prepareIncomingRow(row, year, existingRows));
  const otherRows = existingRows.filter((row) => String(row.year) !== year);
  const allRows = otherRows.concat(incomingRows);
  const sheet = getSheet();

  sheet.clearContents();
  sheet.appendRow(HEADERS);
  allRows.forEach((row) => sheet.appendRow(rowToSheetValues(row)));

  return jsonResponse({ ok: true, rows: incomingRows.length });
}

function prepareIncomingRow(row, year, existingRows) {
  const newImageUrls = saveProjectImages(row.id || `project-${Date.now()}`, row.projectImages || row.projectImage || []);
  const existingRow = (existingRows || []).find((item) => String(item.id) === String(row.id));
  const incomingImageUrls = normalizeImageUrls(row.projectImageUrls || row.projectImageUrl);
  const existingImageUrls = normalizeImageUrls((existingRow && existingRow.projectImageUrls) || (existingRow && existingRow.projectImageUrl));
  const imageUrls = (incomingImageUrls.length ? incomingImageUrls : existingImageUrls).concat(newImageUrls);

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
  const extension = contentType.split("/")[1] || "png";
  const bytes = Utilities.base64Decode(matches[2]);
  const blob = Utilities.newBlob(bytes, contentType, `${id}.${extension}`);
  const file = DriveApp.getFolderById(PROJECT_IMAGE_FOLDER_ID).createFile(blob);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/uc?export=view&id=${file.getId()}`;
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
    row.year || "",
    row.id || "",
    row.project || "",
    row.owner || "",
    row.techPrep || "",
    row.tor || "",
    row.pr || "",
    row.sapPr || "",
    row.po || "",
    row.delivery || "",
    row.fatSat || "",
    row.closeMoc || "",
    row.overallPlan || "",
    imageUrlsText(row.projectImageUrls || row.projectImageUrl),
    milestoneText(row.updatePlan),
    milestoneText(row.updateActual),
    new Date()
  ];
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
      .map((month) => Number(String(month).trim()))
      .filter((month) => month >= 1 && month <= 12)
  )).sort((a, b) => a - b);
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

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

  return sheet;
}
