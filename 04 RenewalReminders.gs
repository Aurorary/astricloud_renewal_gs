/**
 * Check for upcoming contract expirations and send reminder emails
 */
function checkAndSendReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  const renewalSheet = ss.getSheetByName(CONFIG.RENEWAL_SHEET);
  
  if (!trackerSheet) {
    Logger.log('ERROR: TRACKER sheet not found');
    return;
  }
  
  const data = trackerSheet.getDataRange().getValues();
  const now = new Date();
  let remindersSent = 0;
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const contractEnd = data[i][CONFIG.TRACKER_COLS.CONTRACT_END - 1];
    const companyName = data[i][CONFIG.TRACKER_COLS.COMPANY_NAME - 1];
    const companyEmail = data[i][CONFIG.TRACKER_COLS.COMPANY_EMAIL - 1];
    const pilotNumber = data[i][CONFIG.TRACKER_COLS.PILOT_NUMBER - 1];
    
    if (!contractEnd || !companyEmail || !pilotNumber) continue;
    
    const endDate = new Date(contractEnd);
    const monthsUntilExpiry = getMonthsDifference(now, endDate);
    
    // Check if we should send reminder (3, 2, or 1 month before)
    if (CONFIG.REMINDER_MONTHS.includes(monthsUntilExpiry)) {
      
      // Check if already in renewal status with "Renew" status
      const alreadyRenewing = checkIfAlreadyRenewing(renewalSheet, companyName);
      
      if (!alreadyRenewing) {
        sendRenewalReminderEmail(companyName, companyEmail, pilotNumber, endDate, monthsUntilExpiry);
        
        // Add to Renewal Status sheet if not there
        addToRenewalStatus(renewalSheet, data[i], endDate);
        
        remindersSent++;
      }
    }
  }
  
  Logger.log(`Renewal reminders sent: ${remindersSent}`);
}

/**
 * Calculate months between two dates
 */
function getMonthsDifference(date1, date2) {
  const months = (date2.getFullYear() - date1.getFullYear()) * 12 + (date2.getMonth() - date1.getMonth());
  return Math.round(months);
}

/**
 * Check if company is already in renewal status
 */
function checkIfAlreadyRenewing(renewalSheet, companyName) {
  if (!renewalSheet) return false;
  
  const data = renewalSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][CONFIG.RENEWAL_COLS.COMPANIES - 1] === companyName) {
      const finalStatus = data[i][CONFIG.RENEWAL_COLS.FINAL_STATUS - 1];
      if (finalStatus === 'Renew' || finalStatus === 'Pending') {
        return true;
      }
    }
  }
  
  return false;
}

/**
 * Send renewal reminder email
 */
function sendRenewalReminderEmail(companyName, email, pilotNumber, expiryDate, monthsLeft) {
  const expiryStr = Utilities.formatDate(expiryDate, Session.getScriptTimeZone(), 'dd MMM yyyy');
  
  const subject = `⏰ Virtual Landline Renewal Reminder - ${monthsLeft} Month${monthsLeft > 1 ? 's' : ''} Until Expiry`;
  
  const body = `
Dear ${companyName},

This is a friendly reminder that your virtual landline service is expiring soon.

📞 Pilot Number: ${pilotNumber}
📅 Expiry Date: ${expiryStr}
⏳ Time Remaining: ${monthsLeft} month${monthsLeft > 1 ? 's' : ''}

To ensure uninterrupted service, please confirm your renewal at your earliest convenience.

If you have any questions or would like to discuss renewal options, please don't hesitate to reach out.

Best regards,
WORQ IT Operations Team
${CONFIG.EMAIL_FROM}
`;

  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      name: 'WORQ IT Operations'
    });
    
    Logger.log(`Reminder sent to ${companyName} (${email})`);
  } catch (e) {
    Logger.log(`ERROR sending email to ${email}: ${e.message}`);
  }
}

/**
 * Add company to Renewal Status sheet
 */
function addToRenewalStatus(renewalSheet, rowData, contractEnd) {
  if (!renewalSheet) return;
  
  // Check if already exists
  const data = renewalSheet.getDataRange().getValues();
  const companyName = rowData[CONFIG.TRACKER_COLS.COMPANY_NAME - 1];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][CONFIG.RENEWAL_COLS.COMPANIES - 1] === companyName) {
      Logger.log(`${companyName} already in Renewal Status`);
      return;
    }
  }
  
  // Add new row
  const contractStart = rowData[CONFIG.TRACKER_COLS.CONTRACT_START - 1];
  
  const newRow = [
    companyName,                                          // Companies
    rowData[CONFIG.TRACKER_COLS.WORQ_LOCATION - 1],     // Outlet
    contractStart,                                        // Current Start
    contractEnd,                                          // Current End
    'Pending',                                            // Customer Status
    '',                                                   // New Renewal Date (EOM)
    'Pending',                                            // Final Status
    '',                                                   // Remark
    ''                                                    // Email sent to Astricloud
  ];
  
  renewalSheet.appendRow(newRow);
  Logger.log(`Added ${companyName} to Renewal Status`);
}