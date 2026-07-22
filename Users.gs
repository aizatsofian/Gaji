// ============= USERS Pengurusan =============

function getUsers(user) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);
    var data = sheet.getDataRange().getValues();
    var users = [];

    for (var i = 1; i < data.length; i++) {
      users.push({
        ID: data[i][0],
        Username: data[i][1],
        Email: data[i][2],
        Password: data[i][3],
        Role: data[i][4],
        Status: data[i][5],
        CreatedAt: new Date(data[i][6]).toISOString()
      });
    }

    return { success: true, data: users };
  } catch (error) {
    return { success: false, message: 'Error loading users: ' + error.message };
  }
}

function addUser(userData) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);
    var lastRow = sheet.getLastRow();
    var newId = lastRow > 0 ? sheet.getRange(lastRow, 1).getValue() + 1 : 1;

    sheet.appendRow([
      newId,
      userData.Username,
      userData.Email,
      userData.Password,
      userData.Role,
      userData.Status || 'Active',
      new Date().toISOString()
    ]);

    logActivity(userData.CreatedBy || 'Sistem', 'Add User', 'Added user: ' + userData.Username);
    return { success: true, message: 'User added successfully' };
  } catch (error) {
    return { success: false, message: 'Error adding user: ' + error.message };
  }
}

function updateUser(id, userData, user, role) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.getRange(i + 1, 2).setValue(userData.Username);
        sheet.getRange(i + 1, 3).setValue(userData.Email);
        sheet.getRange(i + 1, 4).setValue(userData.Password);
        sheet.getRange(i + 1, 5).setValue(userData.Role);
        sheet.getRange(i + 1, 6).setValue(userData.Status);

        logActivity(user, 'Update User', 'Updated user: ' + userData.Username);
        return { success: true, message: 'User updated successfully' };
      }
    }

    return { success: false, message: 'User not found' };
  } catch (error) {
    return { success: false, message: 'Error updating user: ' + error.message };
  }
}

function deleteUser(id, user, role) {
  try {
    if (role !== 'Admin') {
      return { success: false, message: 'Only admins can delete users' };
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        if (data[i][4] === 'Admin') {
          return { success: false, message: 'Cannot delete admin users' };
        }
        sheet.deleteRow(i + 1);
        logActivity(user, 'Delete User', 'Deleted user ID: ' + id);
        return { success: true, message: 'User deleted successfully' };
      }
    }

    return { success: false, message: 'User not found' };
  } catch (error) {
    return { success: false, message: 'Error deleting user: ' + error.message };
  }
}
