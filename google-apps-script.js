const SHEET_NAME = "ProgressRows";

function doPost(e) {
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
    "PR/PO",
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
      row.prpo,
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
