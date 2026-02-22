/**
 * Color-code TRACKER data rows (cols B–H) by renewal urgency:
 *   Light green  — Renewal Status = "Renewed"
 *   Light red    — Contract End is this month or already past
 *   Light orange — 1 month away
 *   Light yellow — 2–3 months away
 *   Clear        — > 3 months away, no date, Terminated, or Not Renewing
 */
function highlightRenewalUrgency() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);

  if (!trackerSheet) {
    SpreadsheetApp.getUi().alert('❌ TRACKER sheet not found.');
    return;
  }

  const lastRow = trackerSheet.getLastRow();
  if (lastRow < 3) {
    SpreadsheetApp.getUi().alert('ℹ️ No data rows to highlight.');
    return;
  }

  const numDataRows   = lastRow - 2;
  const HIGHLIGHT_END = 7; // cols B–H (7 columns starting at col 2)
  const today         = new Date();
  today.setHours(0, 0, 0, 0);

  // Read cols B–H for all data rows (0-based within this 7-col slice)
  const RENEWAL_IDX  = CONFIG.TRACKER_COLS.RENEWAL_STATUS - 2; // F(6) - B(2) = 4
  const END_DATE_IDX = CONFIG.TRACKER_COLS.CONTRACT_END   - 2; // H(8) - B(2) = 6

  const data        = trackerSheet.getRange(3, 2, numDataRows, HIGHLIGHT_END).getValues();
  const backgrounds = [];

  const counts = { red: 0, orange: 0, yellow: 0, green: 0, clear: 0 };

  for (const row of data) {
    const companyName   = row[0];
    const renewalStatus = row[RENEWAL_IDX];
    const contractEnd   = row[END_DATE_IDX];

    if (!companyName || companyName.toString().trim() === '') {
      backgrounds.push(Array(HIGHLIGHT_END).fill(null));
      counts.clear++;
      continue;
    }

    let color = null;

    if (renewalStatus === 'Renewed') {
      color = '#D9EAD3'; // light green
    } else if (renewalStatus === 'Terminated' || renewalStatus === 'Not Renewing') {
      color = null; // about to be archived or decision already made
    } else if (contractEnd && contractEnd !== '') {
      const endDate = new Date(contractEnd);
      endDate.setHours(0, 0, 0, 0);
      const months = urgencyMonthsDiff_(today, endDate);

      if (months <= 0)      color = '#F4CCCC'; // light red   — expired / expiring this month
      else if (months === 1) color = '#FCE5CD'; // light orange — 1 month away
      else if (months <= 3)  color = '#FFF2CC'; // light yellow — 2–3 months away
    }

    backgrounds.push(Array(HIGHLIGHT_END).fill(color));

    if      (color === '#F4CCCC') counts.red++;
    else if (color === '#FCE5CD') counts.orange++;
    else if (color === '#FFF2CC') counts.yellow++;
    else if (color === '#D9EAD3') counts.green++;
    else                          counts.clear++;
  }

  trackerSheet.getRange(3, 2, numDataRows, HIGHLIGHT_END).setBackgrounds(backgrounds);

  SpreadsheetApp.getUi().alert(
    `✅ Renewal urgency highlighted\n\n` +
    `🔴 Expired / expiring this month : ${counts.red}\n` +
    `🟠 1 month away                  : ${counts.orange}\n` +
    `🟡 2–3 months away               : ${counts.yellow}\n` +
    `🟢 Renewed                       : ${counts.green}\n` +
    `⬜ No urgency / no date          : ${counts.clear}`
  );
}

/**
 * Clear urgency highlight colors from TRACKER data rows (cols B–H).
 */
function clearRenewalHighlights() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);

  if (!trackerSheet) {
    SpreadsheetApp.getUi().alert('❌ TRACKER sheet not found.');
    return;
  }

  const lastRow = trackerSheet.getLastRow();
  if (lastRow < 3) return;

  trackerSheet.getRange(3, 2, lastRow - 2, 7).setBackground(null);
  SpreadsheetApp.getUi().alert('✅ Urgency highlights cleared.');
}

/** Whole-month difference: positive if endDate is in the future. */
function urgencyMonthsDiff_(today, endDate) {
  return (endDate.getFullYear() - today.getFullYear()) * 12 +
         (endDate.getMonth()    - today.getMonth());
}

// ─────────────────────────────────────────────────────────────────────────────

/**
 * Highlight the current month column in TRACKER
 */
function highlightCurrentMonth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  
  if (!sheet) {
    Logger.log('ERROR: TRACKER sheet not found');
    return;
  }
  
  const now = new Date();
  const currentMonthYear = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM-yyyy');
  
  // Get month headers from row 2 (row 1 = year labels, row 2 = month names like Feb-2024)
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  // Clear previous highlighting from FIRST_MONTH column onwards (both header rows + data)
  const monthRange = sheet.getRange(1, CONFIG.TRACKER_COLS.FIRST_MONTH, sheet.getLastRow(), lastCol - CONFIG.TRACKER_COLS.FIRST_MONTH + 1);
  monthRange.setBackground(null);

  // Find and highlight current month column (handles both string and Date-formatted headers)
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    let headerStr = '';
    if (typeof header === 'string') {
      headerStr = header.trim();
    } else if (header instanceof Date) {
      headerStr = Utilities.formatDate(header, Session.getScriptTimeZone(), 'MMM-yyyy');
    }

    if (headerStr === currentMonthYear) {
      const colNumber = i + 1;
      const highlightRange = sheet.getRange(1, colNumber, sheet.getLastRow(), 1);
      highlightRange.setBackground('#00FFFF');
      Logger.log(`Highlighted column ${colNumber}: ${currentMonthYear}`);
      break;
    }
  }
}