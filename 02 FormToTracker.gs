/**
 * Copy new entries from Form Responses to TRACKER
 * Only copies rows that have a pilot number but aren't yet in TRACKER
 */
function copyNewEntriesToTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(CONFIG.FORM_RESPONSES_SHEET);
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  
  if (!formSheet || !trackerSheet) {
    Logger.log('ERROR: Required sheets not found');
    return;
  }
  
  const formData = formSheet.getDataRange().getValues();
  const trackerData = trackerSheet.getDataRange().getValues();
  
  // Get existing company names in tracker (skip header)
  const existingCompanies = trackerData.slice(1).map(row => row[CONFIG.TRACKER_COLS.COMPANY_NAME - 1]);
  
  let copiedCount = 0;
  
  // Start from row 2 (skip header)
  for (let i = 1; i < formData.length; i++) {
    const companyName = formData[i][CONFIG.FORM_COLS.COMPANY_NAME - 1];
    const email = formData[i][CONFIG.FORM_COLS.EMAIL - 1];
    const location = formData[i][CONFIG.FORM_COLS.WORQ_LOCATION - 1];
    
    // Check if company exists and is not already in tracker
    if (companyName && companyName.toString().trim() !== '' && !existingCompanies.includes(companyName)) {
      
      // Calculate next NO
      const lastNo = trackerData.length > 1 ? trackerData[trackerData.length - 1][CONFIG.TRACKER_COLS.NO - 1] : 0;
      const newNo = parseInt(lastNo) + 1;
      
      // Prepare row data (columns A through J)
      // Note: Pilot number, quotation, PO, deposit are empty initially
      const newRow = [
        newNo,                  // NO
        companyName,            // Company Name
        location,               // WORQ Location
        email,                  // Company Email
        '',                     // Pilot Number (empty - to be filled manually)
        '',                     // Quotation #
        '',                     // PO #
        '',                     // Deposit
        '',                     // Contract Start (empty until pilot number added)
        ''                      // Contract End (empty until pilot number added)
      ];
      
      // Append to tracker
      trackerSheet.appendRow(newRow);
      
      Logger.log(`Copied: ${companyName} - waiting for pilot number`);
      copiedCount++;
    }
  }
  
  Logger.log(`Total new entries copied: ${copiedCount}`);
  
  if (copiedCount > 0) {
    SpreadsheetApp.getUi().alert(`✅ Copied ${copiedCount} new entries to TRACKER\n\nPlease add pilot numbers to activate tracking.`);
  }
}

/**
 * Trigger to run when pilot number is added
 * This function monitors changes and auto-populates dates and payment status
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // Only process edits in TRACKER sheet, column E (Pilot Number)
  if (sheet.getName() !== CONFIG.TRACKER_SHEET || range.getColumn() !== CONFIG.TRACKER_COLS.PILOT_NUMBER) {
    return;
  }
  
  const row = range.getRow();
  if (row === 1) return; // Skip header
  
  const pilotNumber = range.getValue();
  
  // Check if pilot number was just added (not empty)
  if (pilotNumber && pilotNumber.toString().trim() !== '') {
    
    // Check if contract start is empty
    const contractStartCell = sheet.getRange(row, CONFIG.TRACKER_COLS.CONTRACT_START);
    
    if (!contractStartCell.getValue() || contractStartCell.getValue() === '') {
      
      // Set contract start to today
      const startDate = new Date();
      contractStartCell.setValue(startDate);
      
      // Set contract end to +12 months
      const endDate = new Date(startDate);
      endDate.setFullYear(endDate.getFullYear() + 1);
      sheet.getRange(row, CONFIG.TRACKER_COLS.CONTRACT_END).setValue(endDate);
      
      // Populate 12 months of "paid" status
      populate12MonthsPaid(sheet, row);
      
      Logger.log(`Auto-populated contract dates and payment status for row ${row}`);
    }
  }
}

/**
 * Populate 12 months of "paid" status for a given row
 */
function populate12MonthsPaid(sheet, rowNumber) {
  const contractStart = sheet.getRange(rowNumber, CONFIG.TRACKER_COLS.CONTRACT_START).getValue();
  
  if (!contractStart) {
    Logger.log(`No contract start date for row ${rowNumber}`);
    return;
  }
  
  const startDate = new Date(contractStart);
  
  // Get header row to find month columns
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Populate 12 months from contract start
  for (let monthOffset = 0; monthOffset < 12; monthOffset++) {
    const targetDate = new Date(startDate);
    targetDate.setMonth(startDate.getMonth() + monthOffset);
    
    const targetMonthYear = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'MMM-yyyy');
    
    // Find matching column header
    const colIndex = headers.findIndex(header => {
      if (typeof header === 'string') {
        return header.trim() === targetMonthYear;
      }
      return false;
    });
    
    if (colIndex !== -1) {
      // Set dropdown to "paid"
      const cell = sheet.getRange(rowNumber, colIndex + 1);
      cell.setValue('paid');
      
      // Apply dropdown data validation if not already set
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['paid', 'renew', 'terminate', 'not proceed'], true)
        .build();
      cell.setDataValidation(rule);
    }
  }
  
  Logger.log(`Populated 12 months for row ${rowNumber}`);
}