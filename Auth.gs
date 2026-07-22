// ============= AUTHENTICATION =============

function authenticateUser(email, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    var userData = userSheet.getDataRange().getValues();

    for (var i = 1; i < userData.length; i++) {
      if (String(userData[i][2]) === String(email) && String(userData[i][3]) === String(password) && userData[i][5] === 'Active') {
        var userRole = userData[i][4];
        var employeeID = userData[i][7] || '';
        var userEmail = userData[i][2];

        // If Employee role, verify and get employee data by matching email
        var employeeData = null;
        if (userRole === 'Employee') {
          var empSheet = ss.getSheetByName(CONFIG.SHEETS.EMPLOYEES);
          var empData = empSheet.getDataRange().getValues();

          for (var j = 1; j < empData.length; j++) {
            // Match by EmployeeID first, then by email as fallback
            if ((employeeID && empData[j][1] === employeeID) || empData[j][3] === userEmail) {
              employeeData = {
                EmployeeID: empData[j][1],
                EmployeeName: empData[j][2],
                Email: empData[j][3],
                Position: empData[j][4],
                BasicSalary: empData[j][5],
                Status: empData[j][6]
              };
              // Update employeeID if it was matched by email
              if (!employeeID && empData[j][3] === userEmail) {
                employeeID = empData[j][1];
              }
              break;
            }
          }
        }

        logActivity(userData[i][1], 'Login', 'User logged in successfully');
        return {
          success: true,
          username: userData[i][1],
          email: userEmail,
          role: userRole,
          employeeID: employeeID,
          employeeData: employeeData,
          message: 'Login successful'
        };
      }
    }

    return { success: false, message: 'Invalid credentials or inactive account' };
  } catch (error) {
    return { success: false, message: 'Authentication error: ' + error.message };
  }
}

// ============= ACCOUNT Pengurusan =============

function changePassword(email, oldPassword, newPassword) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.USERS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]) === String(email) && String(data[i][3]) === String(oldPassword)) {
        sheet.getRange(i + 1, 4).setValue(newPassword);
        logActivity(data[i][1], 'Password Change', 'Password changed successfully');
        return { success: true, message: 'Password changed successfully' };
      }
    }

    return { success: false, message: 'Current password is incorrect' };
  } catch (error) {
    return { success: false, message: 'Error changing password: ' + error.message };
  }
}
