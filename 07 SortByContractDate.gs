/**
 * Sort TRACKER data rows by Contract Start Date (oldest to newest)
 * Both header rows (row 1 = year labels, row 2 = month labels) are preserved.
 * Column A (NO) is skipped — it is driven by the SEQUENCE formula in A3.
 * Rows without a contract start date are pushed to the bottom.
 */
function sortByContractStartDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);

  if (!trackerSheet) {
    Logger.log('ERROR: TRACKER sheet not found');
    SpreadsheetApp.getUi().alert('❌ TRACKER sheet not found.');
    return;
  }

  const lastRow = trackerSheet.getLastRow();
  const lastCol = trackerSheet.getLastColumn();

  // Data rows start at row 3 (rows 1 & 2 are the two header rows)
  if (lastRow < 3) {
    SpreadsheetApp.getUi().alert('ℹ️ No data rows to sort.');
    return;
  }

  const numDataRows = lastRow - 2; // rows 3 → lastRow
  const CONTRACT_START_IDX = CONFIG.TRACKER_COLS.CONTRACT_START - 1; // 0-based index

  // Read all data rows (row 3 onwards), all columns
  const data = trackerSheet.getRange(3, 1, numDataRows, lastCol).getValues();

  // Separate rows that have a contract start date from those that don't
  const rowsWithDate    = [];
  const rowsWithoutDate = [];

  for (const row of data) {
    const companyName   = row[CONFIG.TRACKER_COLS.COMPANY_NAME - 1];
    const contractStart = row[CONTRACT_START_IDX];

    // Treat fully empty rows (no company name) as undated — push to bottom
    if (!companyName || companyName.toString().trim() === '') {
      rowsWithoutDate.push(row);
      continue;
    }

    if (contractStart && contractStart !== '') {
      rowsWithDate.push(row);
    } else {
      rowsWithoutDate.push(row);
    }
  }

  // Sort rows with dates from oldest to newest
  rowsWithDate.sort((a, b) => {
    const dateA = new Date(a[CONTRACT_START_IDX]);
    const dateB = new Date(b[CONTRACT_START_IDX]);
    return dateA - dateB;
  });

  // Dated rows first (sorted), then undated rows at the bottom
  const sortedData = [...rowsWithDate, ...rowsWithoutDate];

  // Write back columns B onwards only — column A is the SEQUENCE formula, leave it untouched
  // slice(1) drops the column A value from each row
  const sortedWithoutColA = sortedData.map(row => row.slice(1));

  if (lastCol > 1) {
    trackerSheet.getRange(3, 2, numDataRows, lastCol - 1).setValues(sortedWithoutColA);
  }

  Logger.log(`Sort complete: ${rowsWithDate.length} rows sorted, ${rowsWithoutDate.length} without date pushed to bottom.`);

  const undatedNote = rowsWithoutDate.length > 0
    ? `\n${rowsWithoutDate.length} company(s) without a contract date moved to the bottom.`
    : '';

  SpreadsheetApp.getUi().alert(
    `✅ Sort complete\n\n` +
    `${rowsWithDate.length} company(s) sorted by Contract Start Date (oldest → newest).` +
    undatedNote
  );
}