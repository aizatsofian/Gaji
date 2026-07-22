// ============= EMPLOYEES Pengurusan =============

function getEmployees(user, role) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    var data = sheet.getDataRange().getValues();
    var employees = [];

    for (var i = 1; i < data.length; i++) {
      employees.push({
        ID: data[i][0],
        EmployeeID: data[i][1],
        EmployeeName: data[i][2],
        Email: data[i][3],
        Position: data[i][4],
        BasicSalary: data[i][5],
        Status: data[i][6],
        JoinDate: data[i][7] ? (typeof data[i][7] === 'string' ? data[i][7] : new Date(data[i][7]).toISOString()) : '',
        CreatedAt: new Date(data[i][8]).toISOString(),
        UpdatedAt: new Date(data[i][9]).toISOString()
      });
    }

    return { success: true, data: employees };
  } catch (error) {
    return { success: false, message: 'Error loading employees: ' + error.message };
  }
}

function addEmployee(empData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var empSheet = ss.getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    var userSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    var timestamp = new Date().toISOString();

    // Check if employee with same EmployeeID already exists
    var empData_existing = empSheet.getDataRange().getValues();
    for (var i = 1; i < empData_existing.length; i++) {
      if (empData_existing[i][1] === empData.EmployeeID) {
        return { success: false, message: 'Employee ID already exists: ' + empData.EmployeeID };
      }
    }

    // Check if user with same email already exists
    var userData = userSheet.getDataRange().getValues();
    for (var j = 1; j < userData.length; j++) {
      if (userData[j][2] === empData.Email) {
        return { success: false, message: 'User with this email already exists: ' + empData.Email };
      }
    }

    // Add Employee to Employees sheet
    var empLastRow = empSheet.getLastRow();
    var newEmpId = empLastRow > 0 ? empSheet.getRange(empLastRow, 1).getValue() + 1 : 1;

    empSheet.appendRow([
      newEmpId,
      empData.EmployeeID,
      empData.EmployeeName,
      empData.Email,
      empData.Position,
      parseFloat(empData.BasicSalary),
      empData.Status || 'Active',
      empData.JoinDate ? new Date(empData.JoinDate).toISOString() : '',
      timestamp,
      timestamp
    ]);

    // Automatically create User account for the employee
    var userLastRow = userSheet.getLastRow();
    var newUserId = userLastRow > 0 ? userSheet.getRange(userLastRow, 1).getValue() + 1 : 1;

    userSheet.appendRow([
      newUserId,
      empData.EmployeeName,        // Username = Employee Name
      empData.Email,               // Email
      empData.Email,               // Password = Email (default)
      'Employee',                  // Role = Employee
      'Active',                    // Status = Active (default)
      timestamp,                   // CreatedAt
      empData.EmployeeID           // EmployeeID (link to employee)
    ]);

    logActivity(empData.CreatedBy || 'Sistem', 'Add Employee', 'Added employee: ' + empData.EmployeeName + ' and created user account');
    return { success: true, message: 'Employee added successfully with user account created (Password: ' + empData.Email + ')' };
  } catch (error) {
    return { success: false, message: 'Error adding employee: ' + error.message };
  }
}

function updateEmployee(id, empData, user, role) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.getRange(i + 1, 2).setValue(empData.EmployeeID);
        sheet.getRange(i + 1, 3).setValue(empData.EmployeeName);
        sheet.getRange(i + 1, 4).setValue(empData.Email);
        sheet.getRange(i + 1, 5).setValue(empData.Position);
        sheet.getRange(i + 1, 6).setValue(parseFloat(empData.BasicSalary));
        sheet.getRange(i + 1, 7).setValue(empData.Status);
        sheet.getRange(i + 1, 8).setValue(empData.JoinDate ? new Date(empData.JoinDate).toISOString() : '');
        sheet.getRange(i + 1, 10).setValue(new Date().toISOString());

        logActivity(user, 'Update Employee', 'Updated employee: ' + empData.EmployeeName);
        return { success: true, message: 'Employee updated successfully' };
      }
    }

    return { success: false, message: 'Employee not found' };
  } catch (error) {
    return { success: false, message: 'Error updating employee: ' + error.message };
  }
}

function deleteEmployee(id, user, role) {
  try {
    if (role !== 'Admin') {
      return { success: false, message: 'Only admins can delete employees' };
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        logActivity(user, 'Delete Employee', 'Deleted employee ID: ' + id);
        return { success: true, message: 'Employee deleted successfully' };
      }
    }

    return { success: false, message: 'Employee not found' };
  } catch (error) {
    return { success: false, message: 'Error deleting employee: ' + error.message };
  }
}

function getEmployeeByID(employeeID) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === employeeID) {
        return {
          success: true,
          data: {
            ID: data[i][0],
            EmployeeID: data[i][1],
            EmployeeName: data[i][2],
            Email: data[i][3],
            Position: data[i][4],
            BasicSalary: data[i][5],
            Status: data[i][6],
            JoinDate: data[i][7] ? (typeof data[i][7] === 'string' ? data[i][7] : new Date(data[i][7]).toISOString()) : ''
          }
        };
      }
    }

    return { success: false, message: 'Employee ID not found' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.message };
  }
}
