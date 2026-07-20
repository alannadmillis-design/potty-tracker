// Google Apps Script for Potty Tracker
// Copy this entire code into your Google Apps Script editor

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  
  if (e.parameter.action === 'getData') {
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: getData(sheet)
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (e.parameter.action === 'getTaskConfig') {
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      config: getTaskConfig()
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (e.parameter.action === 'getRedemptions') {
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      redemptions: getRedemptions()
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    success: false,
    error: 'Invalid action'
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
    const redemptionsSheet = getOrCreateRedemptionsSheet();
    const request = JSON.parse(e.postData.contents);
    
    if (request.action === 'saveData') {
      saveData(sheet, request.data);
      return ContentService.createTextOutput(JSON.stringify({
        success: true
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (request.action === 'addEvent') {
      addEvent(sheet, request.event);
      return ContentService.createTextOutput(JSON.stringify({
        success: true
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (request.action === 'updateGameStats') {
      // Update game stats (stars and points) without affecting events
      updateGameStats(sheet, request.totalStars, request.totalGamePoints);
      return ContentService.createTextOutput(JSON.stringify({
        success: true
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (request.action === 'addRedemption') {
      // Record a reward redemption
      addRedemption(redemptionsSheet, request.redemption);
      return ContentService.createTextOutput(JSON.stringify({
        success: true
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (request.action === 'saveTaskConfig') {
      saveTaskConfig(request.config);
      return ContentService.createTextOutput(JSON.stringify({
        success: true
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (request.action === 'setRedemptionFulfilled') {
      setRedemptionFulfilled(redemptionsSheet, request.id, request.fulfilled);
      return ContentService.createTextOutput(JSON.stringify({
        success: true
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: 'Invalid action'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getData(sheet) {
  // Get summary data from row 2
  const totalStars = sheet.getRange('A2').getValue() || 0;
  const totalGamePoints = sheet.getRange('B2').getValue() || 0;
  
  // Get events from rows starting at row 5
  const eventsData = sheet.getRange('A5:F' + sheet.getLastRow()).getValues();
  const events = eventsData
    .filter(row => row[0]) // Filter out empty rows
    .map(row => ({
      label: row[0],
      method: row[1],
      time: row[2],
      stars: row[3],
      timestamp: row[4],
      category: row[5]
    }));
  
  return {
    totalStars: totalStars,
    totalGamePoints: totalGamePoints,
    events: events
  };
}

function saveData(sheet, data) {
  // Save summary to row 2
  sheet.getRange('A2').setValue(data.totalStars || 0);
  sheet.getRange('B2').setValue(data.totalGamePoints || 0);
  
  // Clear old events (starting from row 5)
  if (sheet.getLastRow() >= 5) {
    sheet.getRange('A5:F' + sheet.getLastRow()).clearContent();
  }
  
  // Save events
  if (data.events && data.events.length > 0) {
    const eventsArray = data.events.map(event => [
      event.label || event.type || '',
      event.method,
      event.time,
      event.stars,
      event.timestamp || new Date().toISOString(),
      event.category || ''
    ]);
    
    sheet.getRange(5, 1, eventsArray.length, 6).setValues(eventsArray);
  }
}

function addEvent(sheet, event) {
  // Get current data
  const data = getData(sheet);
  
  // Add new event
  data.events.push(event);
  
  // Update total stars
  data.totalStars += event.stars;
  
  // Save everything back
  saveData(sheet, data);
}

function updateGameStats(sheet, totalStars, totalGamePoints) {
  // Just update the totals, don't touch events
  sheet.getRange('A2').setValue(totalStars || 0);
  sheet.getRange('B2').setValue(totalGamePoints || 0);
}

function getOrCreateTaskConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('TaskConfig');
  
  if (!sheet) {
    sheet = ss.insertSheet('TaskConfig');
    sheet.getRange('A1').setValue('Config JSON (do not edit by hand)');
    sheet.getRange('A1').setFontWeight('bold');
    sheet.getRange('A2').setValue('{}');
    sheet.setColumnWidth(1, 500);
  }
  
  return sheet;
}

function getTaskConfig() {
  const sheet = getOrCreateTaskConfigSheet();
  const raw = sheet.getRange('A2').getValue();
  
  if (!raw) return {};
  
  try {
    return JSON.parse(raw);
  } catch (error) {
    return {};
  }
}

function saveTaskConfig(config) {
  const sheet = getOrCreateTaskConfigSheet();
  sheet.getRange('A2').setValue(JSON.stringify(config || {}));
}

function getOrCreateRedemptionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Redemptions');
  
  if (!sheet) {
    sheet = ss.insertSheet('Redemptions');
    
    // Set up headers
    sheet.getRange('A1').setValue('Date/Time');
    sheet.getRange('B1').setValue('Reward');
    sheet.getRange('C1').setValue('Cost (Points)');
    sheet.getRange('D1').setValue('Emoji');
    sheet.getRange('E1').setValue('ID');
    sheet.getRange('F1').setValue('Fulfilled');
    
    // Format headers
    sheet.getRange('A1:F1').setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    // Set column widths
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(5, 180);
    sheet.setColumnWidth(6, 80);
  } else if (!sheet.getRange('E1').getValue()) {
    // Backfill headers on a sheet created before ID/Fulfilled tracking existed
    sheet.getRange('E1').setValue('ID');
    sheet.getRange('F1').setValue('Fulfilled');
    sheet.getRange('E1:F1').setFontWeight('bold');
  }
  
  return sheet;
}

function getRedemptions() {
  const sheet = getOrCreateRedemptionsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  
  return rows
    .filter(row => row[1]) // must have a reward label
    .map(row => ({
      timestamp: row[0] instanceof Date ? row[0].toISOString() : row[0],
      label: row[1],
      cost: row[2],
      emoji: row[3],
      id: row[4] || '',
      fulfilled: row[5] === true || row[5] === 'TRUE'
    }));
}

function setRedemptionFulfilled(sheet, id, fulfilled) {
  if (!id) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const ids = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      sheet.getRange(i + 2, 6).setValue(!!fulfilled);
      return;
    }
  }
}

function addRedemption(sheet, redemption) {
  // Find the next empty row
  const lastRow = sheet.getLastRow();
  const nextRow = lastRow + 1;
  
  // Add the redemption
  sheet.getRange(nextRow, 1).setValue(new Date());
  sheet.getRange(nextRow, 2).setValue(redemption.label);
  sheet.getRange(nextRow, 3).setValue(redemption.cost);
  sheet.getRange(nextRow, 4).setValue(redemption.emoji);
  sheet.getRange(nextRow, 5).setValue(redemption.id || '');
  sheet.getRange(nextRow, 6).setValue(false);
  
  // Send email notification
  sendRedemptionEmail(redemption);
}

function sendRedemptionEmail(redemption) {
  // Get the email address from script properties (you'll set this)
  const properties = PropertiesService.getScriptProperties();
  let email = properties.getProperty('NOTIFICATION_EMAIL');
  
  // If no email is set, use the owner's email as default
  if (!email) {
    email = Session.getActiveUser().getEmail();
  }
  
  // Create the email
  const subject = redemption.emoji + ' Potty Tracker: Reward Redeemed!';
  const body = `Great news!\n\n` +
               `Your child just redeemed a reward:\n\n` +
               `${redemption.emoji} ${redemption.label}\n` +
               `Cost: ${redemption.cost} points\n` +
               `Time: ${new Date().toLocaleString()}\n\n` +
               `Keep up the great work! 🌟`;
  
  try {
    MailApp.sendEmail(email, subject, body);
  } catch (error) {
    console.error('Failed to send email:', error);
  }
}

// Function to set your notification email
function setNotificationEmail() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Set Notification Email',
    'Enter the email address where you want to receive reward notifications:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const email = response.getResponseText();
    PropertiesService.getScriptProperties().setProperty('NOTIFICATION_EMAIL', email);
    ui.alert('Email set to: ' + email);
  }
}

// Initialize the sheet structure
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Data');
  
  // Create Data sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet('Data');
  }
  
  // Set up headers
  sheet.getRange('A1').setValue('Total Stars');
  sheet.getRange('B1').setValue('Total Game Points');
  
  // Set up event headers
  sheet.getRange('A4').setValue('Description');
  sheet.getRange('B4').setValue('Method');
  sheet.getRange('C4').setValue('Time');
  sheet.getRange('D4').setValue('Stars');
  sheet.getRange('E4').setValue('Timestamp');
  sheet.getRange('F4').setValue('Category');
  
  // Initialize values
  sheet.getRange('A2').setValue(0);
  sheet.getRange('B2').setValue(0);
  
  // Format headers
  sheet.getRange('A1:E1').setFontWeight('bold');
  sheet.getRange('A4:F4').setFontWeight('bold');
  
  // Create Redemptions sheet
  getOrCreateRedemptionsSheet();
  getOrCreateTaskConfigSheet();
  
  // Prompt for notification email
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Email Notifications',
    'Enter your email address to receive notifications when rewards are redeemed (or leave blank to skip):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() == ui.Button.OK && response.getResponseText()) {
    PropertiesService.getScriptProperties().setProperty('NOTIFICATION_EMAIL', response.getResponseText());
  }
  
  SpreadsheetApp.getUi().alert('Sheets setup complete! Data and Redemptions sheets are ready. Now deploy as a web app.');
}
