/**
 * Data Fetcher → Recap Sheet Synchronization
 *
 * This script synchronizes a dynamic dataset (retrieved from an API,
 * BI platform, or other automated data source) into a structured
 * recap/reporting sheet inside Google Sheets.
 *
 * The dataset size may change every time the source query runs.
 * This script ensures the recap sheet remains consistent by:
 *
 * 1. Detecting the latest available dataset
 * 2. Adjusting row counts automatically
 * 3. Copying values into the recap sheet
 * 4. Preserving formulas used for reporting
 * 5. Optionally sending a webhook notification
 *
 * Typical pipeline:
 *
 * BI Platform / API
 *        ↓
 * Raw Data Sheet
 *        ↓
 * Recap Sheet Sync (this script)
 *        ↓
 * Reporting / Dashboard Sheets
 *
 * NOTE:
 * Replace placeholder configuration values before deployment.
 */

const CONFIG = {

  /** SOURCE DATA SETTINGS */
  SOURCE_SPREADSHEET_ID: "SOURCE_SPREADSHEET_ID_PLACEHOLDER",
  SOURCE_SHEET_NAME: "RAW_DATA_SHEET_NAME",

  /** TARGET RECAP SHEET */
  TARGET_SHEET_NAME: "RECAP_SHEET_NAME",

  /** NUMBER OF VALUE COLUMNS TO COPY */
  VALUE_COLUMNS: 5,

  /** OPTIONAL NOTIFICATION WEBHOOK */
  WEBHOOK_URL: "WEBHOOK_URL_PLACEHOLDER"

};


/**
 * Main entry point
 */
function syncRecapSheet() {

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = getSheet(ss, CONFIG.TARGET_SHEET_NAME);

    const latestDate = findLatestDate(targetSheet);
    const sourceData = fetchSourceData();

    updateRecapRows(targetSheet, sourceData, latestDate);

    const nextDate = appendNextDateRows(
      targetSheet,
      sourceData,
      latestDate
    );

    sendNotification(ss, latestDate, nextDate);

    Logger.log("Recap sheet synchronization completed.");

  } catch (error) {

    Logger.log(`Sync error: ${error.message}`);
    throw error;

  }
}


/**
 * Safely get sheet
 */
function getSheet(spreadsheet, sheetName) {

  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found.`);
  }

  return sheet;
}


/**
 * Find latest date in column A
 */
function findLatestDate(sheet) {

  const values = sheet.getRange("A:A").getValues();

  for (let i = values.length - 1; i >= 0; i--) {

    if (values[i][0]) {
      return values[i][0];
    }

  }

  throw new Error("No date found in column A.");

}


/**
 * Retrieve source dataset
 */
function fetchSourceData() {

  const sourceSS = SpreadsheetApp.openById(
    CONFIG.SOURCE_SPREADSHEET_ID
  );

  const sourceSheet = getSheet(
    sourceSS,
    CONFIG.SOURCE_SHEET_NAME
  );

  return sourceSheet.getDataRange().getValues();

}


/**
 * Synchronize rows for the latest date
 */
function updateRecapRows(sheet, sourceData, latestDate) {

  const columnValues = sheet.getRange("A:A").getValues();

  const targetRows = columnValues
    .map((row, i) =>
      isSameDate(row[0], latestDate) ? i + 1 : null
    )
    .filter(Boolean);

  const rowsForLatest = sourceData
    .filter(r => isSameDate(r[0], latestDate))
    .map(r => r.slice(0, CONFIG.VALUE_COLUMNS));

  if (rowsForLatest.length > targetRows.length) {

    const extra = rowsForLatest.length - targetRows.length;
    const insertAfter = targetRows[targetRows.length - 1];

    sheet.insertRowsAfter(insertAfter, extra);

    for (let i = 1; i <= extra; i++) {
      targetRows.push(insertAfter + i);
    }

  }

  writeRows(sheet, targetRows, rowsForLatest);

}


/**
 * Append rows for the next dataset date
 */
function appendNextDateRows(sheet, sourceData, latestDate) {

  const nextDate = new Date(latestDate);
  nextDate.setDate(nextDate.getDate() + 1);

  const rowsForNext = sourceData
    .filter(r => isSameDate(r[0], nextDate))
    .map(r => r.slice(0, CONFIG.VALUE_COLUMNS));

  if (!rowsForNext.length) {
    return nextDate;
  }

  const lastRow = sheet.getLastRow();

  sheet.insertRowsAfter(lastRow, rowsForNext.length);

  rowsForNext.forEach((values, index) => {

    const rowNumber = lastRow + 1 + index;

    sheet
      .getRange(rowNumber, 1, 1, CONFIG.VALUE_COLUMNS)
      .setValues([values]);

    applyFormulas(sheet, rowNumber);

  });

  return nextDate;

}


/**
 * Write dataset rows
 */
function writeRows(sheet, rows, values) {

  rows.forEach((rowNumber, index) => {

    sheet.getRange(rowNumber, 1, 1, 8).clearContent();

    if (index < values.length) {

      sheet
        .getRange(rowNumber, 1, 1, CONFIG.VALUE_COLUMNS)
        .setValues([values[index]]);

    }

    applyFormulas(sheet, rowNumber);

  });

}


/**
 * Apply recap formulas
 */
function applyFormulas(sheet, rowNumber) {

  sheet
    .getRange(rowNumber, 6)
    .setFormula(`=ISOWEEKNUM(A${rowNumber})`);

  sheet
    .getRange(rowNumber, 7)
    .setFormula(`=MONTH(A${rowNumber})`);

  sheet
    .getRange(rowNumber, 8)
    .setFormula(`=YEAR(A${rowNumber})`);

}


/**
 * Compare two dates
 */
function isSameDate(a, b) {

  if (a instanceof Date && b instanceof Date) {

    return (
      a.getFullYear() === b.getFullYear() &&
      a.getMonth() === b.getMonth() &&
      a.getDate() === b.getDate()
    );

  }

  return a === b;

}


/**
 * Send optional webhook notification
 */
function sendNotification(ss, latestDate, nextDate) {

  if (!CONFIG.WEBHOOK_URL ||
      CONFIG.WEBHOOK_URL === "WEBHOOK_URL_PLACEHOLDER") {
    return;
  }

  const message = {

    text:
      `Recap sheet synchronized for ` +
      `${formatDate(ss, latestDate)} → ${formatDate(ss, nextDate)}`

  };

  UrlFetchApp.fetch(CONFIG.WEBHOOK_URL, {

    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(message)

  });

}


/**
 * Format date for logs
 */
function formatDate(ss, date) {

  return Utilities.formatDate(
    date,
    ss.getSpreadsheetTimeZone(),
    "yyyy-MM-dd"
  );

}
