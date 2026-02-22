/**
 * Sync renewals from TRACKER col F (Renewal Status)
 * When Renewal Status = "Renew", extend contract by 12 months
 */
function syncRenewals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);

  if (!trackerSheet) {
    Logger.log('ERROR: TRACKER sheet not found');
    return;
  }

  const trackerData = trackerSheet.getDataRange().getValues();
  let renewalsProcessed = 0;
  const renewedCompanies = [];
  const terminatedCompanies = [];
  let emailsSent = 0;

  // Skip header rows (data starts row 3, i=2)
  for (let i = 2; i < trackerData.length; i++) {
    const companyName   = trackerData[i][CONFIG.TRACKER_COLS.COMPANY_NAME - 1];
    const companyEmail  = trackerData[i][CONFIG.TRACKER_COLS.COMPANY_EMAIL - 1];
    const pilotNumber   = trackerData[i][CONFIG.TRACKER_COLS.PILOT_NUMBER - 1];
    const renewalStatus = trackerData[i][CONFIG.TRACKER_COLS.RENEWAL_STATUS - 1];
    const contractEnd   = trackerData[i][CONFIG.TRACKER_COLS.CONTRACT_END - 1];

    if (!contractEnd) continue;

    const currentEndDate = new Date(contractEnd);

    // --- Handle Not Renewing ---
    if (renewalStatus === 'Not Renewing') {
      markMonthCell(trackerSheet, i + 1, currentEndDate, 'terminate');
      trackerSheet.getRange(i + 1, CONFIG.TRACKER_COLS.RENEWAL_STATUS).setValue('Terminated');
      if (companyEmail) {
        sendTerminationEmail(companyName, companyEmail, pilotNumber, currentEndDate);
        emailsSent++;
      }
      terminatedCompanies.push(companyName);
      Logger.log(`Marked ${companyName} as Terminated. Termination email sent. Run Archive Terminated Customers when ready.`);
      continue;
    }

    if (renewalStatus !== 'Renew') continue;

    // New tenure: starts 1st of the month after current end, ends 12 months later
    const newStartDate = new Date(currentEndDate.getFullYear(), currentEndDate.getMonth() + 1, 1);
    const newEndDate   = new Date(currentEndDate);
    newEndDate.setFullYear(newEndDate.getFullYear() + 1);

    // Update contract end date in TRACKER
    trackerSheet.getRange(i + 1, CONFIG.TRACKER_COLS.CONTRACT_END).setValue(newEndDate);

    // Extend month header columns if the new contract end goes beyond existing headers
    extendMonthHeaders(trackerSheet, newEndDate);

    // Populate another 12 months — use day 1 to avoid month-end overflow (e.g. Mar 31 + 1 month = May 1, not Apr 1)
    populate12MonthsFromDate(trackerSheet, i + 1, newStartDate);

    // Send thank you email with new tenure
    if (companyEmail) {
      sendRenewalConfirmationEmail(companyName, companyEmail, pilotNumber, newStartDate, newEndDate);
      emailsSent++;
    }

    // Mark as Renewed in col F
    trackerSheet.getRange(i + 1, CONFIG.TRACKER_COLS.RENEWAL_STATUS).setValue('Renewed');

    const newStartStr = Utilities.formatDate(newStartDate, Session.getScriptTimeZone(), 'MMM yyyy');
    const newEndStr   = Utilities.formatDate(newEndDate,   Session.getScriptTimeZone(), 'MMM yyyy');
    renewedCompanies.push(`${companyName} (${newStartStr} – ${newEndStr})`);
    Logger.log(`Extended contract for ${companyName}: ${newStartDate.toDateString()} – ${newEndDate.toDateString()}`);
    renewalsProcessed++;
  }

  Logger.log(`Renewals processed: ${renewalsProcessed}`);

  const parts = [];

  if (renewedCompanies.length > 0) {
    parts.push(`Renewals extended: ${renewedCompanies.length}\n${renewedCompanies.map(n => `• ${n}`).join('\n')}`);
  }
  if (terminatedCompanies.length > 0) {
    parts.push(`Terminated: ${terminatedCompanies.length}\n${terminatedCompanies.map(n => `• ${n}`).join('\n')}`);
  }

  let message;
  if (parts.length > 0) {
    const emailLine = emailsSent > 0 ? `\n\nEmails sent: ${emailsSent}` : '';
    message = `✅ Sync complete\n\n${parts.join('\n\n')}${emailLine}`;
  } else {
    message = `ℹ️ Nothing to sync.\n\nNo companies are marked "Renew" or "Not Renewing".`;
  }
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Populate 12 months from a specific start date
 */
function populate12MonthsFromDate(sheet, rowNumber, startDate) {
  // Month headers are in row 2 (row 1 = year labels)
  const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];

  for (let monthOffset = 0; monthOffset < 12; monthOffset++) {
    const targetDate = new Date(startDate);
    targetDate.setMonth(startDate.getMonth() + monthOffset);

    const targetMonthYear = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'MMM-yyyy');

    const colIndex = headers.findIndex(header => {
      if (typeof header === 'string') {
        return header.trim() === targetMonthYear;
      }
      if (header instanceof Date) {
        return Utilities.formatDate(header, Session.getScriptTimeZone(), 'MMM-yyyy') === targetMonthYear;
      }
      return false;
    });

    if (colIndex !== -1) {
      const cell = sheet.getRange(rowNumber, colIndex + 1);
      // First month of renewal cycle = 'renew', remaining months = 'paid'
      cell.setValue(monthOffset === 0 ? 'renew' : 'paid');

      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['paid', 'renew', 'terminate', 'not proceed'], true)
        .build();
      cell.setDataValidation(rule);
    } else {
      Logger.log(`Column not found for ${targetMonthYear} — header may be missing or out of range`);
    }
  }
}

/**
 * Extend TRACKER month header columns (rows 1 & 2) up to the given date.
 * Row 2: adds MMM-yyyy labels. Row 1: adds year number at the first month of each new year.
 * No-ops if headers already cover the required range.
 */
function extendMonthHeaders(sheet, upToDate) {
  const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  // Find the last month column by scanning backwards
  let lastMonthDate = null;
  let lastMonthCol = -1; // 1-based column number

  for (let i = headers.length - 1; i >= CONFIG.TRACKER_COLS.FIRST_MONTH - 1; i--) {
    const header = headers[i];
    let parsed = null;

    if (typeof header === 'string' && header.trim() !== '') {
      const parts = header.trim().split('-');
      if (parts.length === 2) {
        const m = monthNames.indexOf(parts[0]);
        const y = parseInt(parts[1]);
        if (m !== -1 && !isNaN(y)) parsed = new Date(y, m, 1);
      }
    } else if (header instanceof Date) {
      parsed = new Date(header.getFullYear(), header.getMonth(), 1);
    }

    if (parsed) {
      lastMonthDate = parsed;
      lastMonthCol = i + 1; // convert to 1-based
      break;
    }
  }

  if (!lastMonthDate) {
    Logger.log('extendMonthHeaders: Could not find last month header');
    return;
  }

  const upToMonth = new Date(upToDate.getFullYear(), upToDate.getMonth(), 1);
  if (upToMonth <= lastMonthDate) {
    Logger.log('extendMonthHeaders: Headers already cover up to ' + Utilities.formatDate(upToMonth, Session.getScriptTimeZone(), 'MMM-yyyy'));
    return;
  }

  // Append new month columns one by one
  let current = new Date(lastMonthDate.getFullYear(), lastMonthDate.getMonth() + 1, 1);
  let colNumber = lastMonthCol + 1;
  let added = 0;

  while (current <= upToMonth) {
    const monthLabel = Utilities.formatDate(current, Session.getScriptTimeZone(), 'MMM-yyyy');

    // Row 2: month label (e.g. Jan-2028)
    sheet.getRange(2, colNumber).setValue(monthLabel);

    // Row 1: year label only on the first month of each new year (January)
    if (current.getMonth() === 0) {
      sheet.getRange(1, colNumber).setValue(current.getFullYear());
    }

    Logger.log('extendMonthHeaders: Added column ' + colNumber + ' → ' + monthLabel);
    current = new Date(current.getFullYear(), current.getMonth() + 1, 1);
    colNumber++;
    added++;
  }

  Logger.log('extendMonthHeaders: Added ' + added + ' new column(s) up to ' + Utilities.formatDate(upToMonth, Session.getScriptTimeZone(), 'MMM-yyyy'));
}

/**
 * Set a specific month column cell to a given value for a row
 */
function markMonthCell(sheet, rowNumber, targetDate, value) {
  const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const targetMonthYear = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'MMM-yyyy');

  const colIndex = headers.findIndex(header => {
    if (typeof header === 'string') return header.trim() === targetMonthYear;
    if (header instanceof Date) return Utilities.formatDate(header, Session.getScriptTimeZone(), 'MMM-yyyy') === targetMonthYear;
    return false;
  });

  if (colIndex !== -1) {
    const cell = sheet.getRange(rowNumber, colIndex + 1);
    cell.setValue(value);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['paid', 'renew', 'terminate', 'not proceed'], true)
      .build();
    cell.setDataValidation(rule);
    Logger.log(`Set ${targetMonthYear} → '${value}' for row ${rowNumber}`);
  } else {
    Logger.log(`Column not found for ${targetMonthYear} — could not mark as '${value}'`);
  }
}