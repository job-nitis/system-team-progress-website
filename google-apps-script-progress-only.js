// Use this older script only for the System Team Progress table.
// It reads and writes the ProgressRows sheet tab, but does not handle Highlight data or image uploads.
const SHEET_NAME = "ProgressRows";
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

  return ContentService
    .createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return jsonResponse({ ok: false, error: "No POST data received" });
  }

  const payload = JSON.parse(e.postData.contents);
  const year = String(payload.year || "");
  const incomingRows = (payload.rows || []).map((row) => ({ ...row, year }));
  const otherRows = readRows("").filter((row) => String(row.year) !== year);
  const allRows = otherRows.concat(incomingRows);
  const sheet = getSheet();

  sheet.clearContents();
  sheet.appendRow(HEADERS);
  allRows.forEach((row) => sheet.appendRow(rowToSheetValues(row)));

  return jsonResponse({ ok: true, rows: incomingRows.length });
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
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
    overallPlan: values[12]
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
    new Date()
  ];
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

  return sheet;
}
