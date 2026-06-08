const SHEET_NAME = "ProgressRows";

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: "System Team Progress API is running" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: "No POST data received" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const payload = JSON.parse(e.postData.contents);
  const sheet = getSheet();
  const rows = payload.rows || [];
  const year = payload.year || "";

  sheet.clearContents();
  sheet.appendRow([
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
  ]);

  rows.forEach((row) => {
    sheet.appendRow([
      year,
      row.id,
      row.project,
      row.owner,
      row.techPrep,
      row.tor,
      row.pr,
      row.sapPr,
      row.po,
      row.delivery,
      row.fatSat,
      row.closeMoc,
      row.overallPlan,
      new Date()
    ]);
  });

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}
