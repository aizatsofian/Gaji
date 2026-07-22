// ============= ACTIVITY LOGS =============

function getLogs(user, role) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOGS);
    var data = sheet.getDataRange().getValues();
    var logs = [];

    for (var i = 1; i < data.length; i++) {
      logs.push({
        ID: data[i][0],
        Username: data[i][1],
        Action: data[i][2],
        Details: data[i][3],
        Timestamp: new Date(data[i][4]).toISOString()
      });
    }

    return { success: true, data: logs };
  } catch (error) {
    return { success: false, message: 'Error loading logs: ' + error.message };
  }
}

function logActivity(username, action, details) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LOGS);
    var lastRow = sheet.getLastRow();
    var newId = lastRow > 0 ? sheet.getRange(lastRow, 1).getValue() + 1 : 1;

    sheet.appendRow([
      newId,
      username,
      action,
      details,
      new Date().toISOString()
    ]);
  } catch (error) {
    Logger.log('Error logging activity: ' + error.message);
  }
}
