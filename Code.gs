// ══════════════════════════════════════════════════════════════════════
//  CONTACT LENS ORDER — Google Apps Script Web App
//  File: Code.gs
//
//  DEPLOYMENT STEPS:
//  1. Open https://script.google.com → New Project → paste this code
//  2. Click "Deploy" → "New deployment"
//  3. Type: Web App
//     Execute as: Me
//     Who has access: Anyone
//  4. Click Deploy → copy the Web App URL
//  5. Paste that URL into the HTML file's CONFIG.GAS_URL field
//
//  SPREADSHEET TARGET:
//  File: 系統資料
//  Sheet: 工作表1
//  Spreadsheet ID: 10MAsKjnh2oW1gR-M-yFsmVeZQ14GbtshU82O931VPYI
// ══════════════════════════════════════════════════════════════════════

const SPREADSHEET_ID = "10MAsKjnh2oW1gR-M-yFsmVeZQ14GbtshU82O931VPYI";
const SHEET_NAME     = "工作表1";

// Column headers written to row 1 on first use
const HEADERS = [
  "Timestamp",
  "Submitter",
  "Item #",
  "Product Code",
  "Power (Degree)",
  "Color",
  "Cyl",
  "Axis",
  "Quantity",
  "Notes"
];


// ─────────────────────────────────────────────
//  doPost — receives JSON payload from the app
// ─────────────────────────────────────────────
function doPost(e) {
  try {
    // Parse incoming JSON body
    const payload = JSON.parse(e.postData.contents);

    const submitter  = payload.submitter  || "Unknown";
    const timestamp  = payload.timestamp  || new Date().toISOString();
    const items      = payload.items      || [];

    if (items.length === 0) {
      return buildResponse({ status: "error", message: "No items in payload." });
    }

    const sheet = getOrCreateSheet();

    // Write each item as its own row
    const rows = items.map(item => [
      timestamp,
      submitter,
      item.index  || "",
      item.code   || "",
      item.power  || "",
      item.color  || "",
      item.cyl    || "",
      item.axis   || "",
      item.qty    || "",
      item.notes  || ""
    ]);

    const lastRow = Math.max(sheet.getLastRow(), 1);
    sheet.getRange(lastRow + 1, 1, rows.length, HEADERS.length).setValues(rows);

    return buildResponse({ status: "ok", rowsWritten: rows.length });

  } catch(err) {
    return buildResponse({ status: "error", message: err.message });
  }
}


// ─────────────────────────────────────────────
//  doGet — health-check / CORS preflight helper
// ─────────────────────────────────────────────
function doGet(e) {
  return buildResponse({ status: "ok", message: "Contact Lens Order API is running." });
}


// ─────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────

/**
 * Returns the target sheet, creating headers if the sheet is brand-new.
 */
function getOrCreateSheet() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    // Sheet doesn't exist yet — create it
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Write headers if row 1 is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);

    // Style the header row
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight("bold")
               .setBackground("#c8603a")
               .setFontColor("#ffffff")
               .setHorizontalAlignment("center");

    // Freeze header row
    sheet.setFrozenRows(1);

    // Auto-resize columns
    for (let i = 1; i <= HEADERS.length; i++) {
      sheet.setColumnWidth(i, 130);
    }
  }

  return sheet;
}

/**
 * Builds a CORS-friendly JSON response.
 */
function buildResponse(obj) {
  const output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);

  // Note: Apps Script automatically handles CORS for Web Apps
  // deployed as "Anyone can access". No manual headers needed.
  return output;
}


// ──────────────────────────────────────────────────────────────────────
//  TEST FUNCTION (run from the script editor to verify everything works)
// ──────────────────────────────────────────────────────────────────────
function testWrite() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        submitter:  "Test User",
        timestamp:  new Date().toISOString(),
        items: [
          {
            index: 1,
            code:  "NX-01A",
            power: "-3.00",
            color: "Brown",
            cyl:   "-0.75",
            axis:  "180",
            qty:   "2",
            notes: "Test order from script editor"
          }
        ]
      })
    }
  };

  const result = doPost(fakeEvent);
  Logger.log(result.getContent());
}
