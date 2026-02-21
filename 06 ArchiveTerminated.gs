/**
 * Archive customers with "Terminate" status
 */
function archiveTerminated() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  const archivedSheet = ss.getSheetByName(CONFIG.ARCHIVED_SHEET);
  
  if (!trackerSheet || !archivedSheet) {
    Logger.log('ERROR: Required sheets not found');
    return;
  }
  
  const data = trackerSheet.getDataRange().getValues();
  const headers = data[0];
  let archivedCount = 0;
  
  // Process from bottom to top (to avoid index shifting when deleting rows)
  for (let i = data.length - 1; i >= 1; i--) {
    
    // Check monthly status columns for "terminate"
    let hasTerminate = false;
    
    for (let col = CONFIG.TRACKER_COLS.FIRST_MONTH - 1; col < data[i].length; col++) {
      if (data[i][col] === 'terminate') {
        hasTerminate = true;
        break;
      }
    }
    
    if (hasTerminate) {
      // Copy entire row to ARCHIVED
      archivedSheet.appendRow(data[i]);
      
      // Delete from TRACKER
      trackerSheet.deleteRow(i + 1);
      
      const companyName = data[i][CONFIG.TRACKER_COLS.COMPANY_NAME - 1];
      Logger.log(`Archived: ${companyName}`);
      archivedCount++;
    }
  }
  
  Logger.log(`Total customers archived: ${archivedCount}`);

  if (archivedCount > 0) {
    SpreadsheetApp.getUi().alert(`✅ Archived ${archivedCount} terminated customers`);
  }
}

/**
 * Remove any TRACKER rows whose company name already exists in ARCHIVED.
 * Cleans up companies that were re-added before the duplicate-check fix.
 */
function removeArchivedDuplicatesFromTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  const archivedSheet = ss.getSheetByName(CONFIG.ARCHIVED_SHEET);

  if (!trackerSheet || !archivedSheet) {
    Logger.log('ERROR: Required sheets not found');
    return;
  }

  const archivedData = archivedSheet.getDataRange().getValues();
  // Build a Set of archived company names (skip header row)
  const archivedNames = new Set(
    archivedData.slice(1)
      .map(row => row[CONFIG.TRACKER_COLS.COMPANY_NAME - 1].toString().trim().toLowerCase())
      .filter(name => name !== '')
  );

  const trackerData = trackerSheet.getDataRange().getValues();
  let removedCount = 0;

  // Process bottom-to-top to avoid index shifting on deleteRow
  for (let i = trackerData.length - 1; i >= 2; i--) {
    const companyName = trackerData[i][CONFIG.TRACKER_COLS.COMPANY_NAME - 1].toString().trim();
    if (companyName === '') continue;

    if (archivedNames.has(companyName.toLowerCase())) {
      trackerSheet.deleteRow(i + 1);
      Logger.log(`Removed from TRACKER (already in ARCHIVED): ${companyName}`);
      removedCount++;
    }
  }

  Logger.log(`Cleanup complete: ${removedCount} duplicate(s) removed from TRACKER`);
  SpreadsheetApp.getUi().alert(
    removedCount > 0
      ? `✅ Removed ${removedCount} company(s) from TRACKER that already exist in ARCHIVED.`
      : `✅ No duplicates found. TRACKER is clean.`
  );
}