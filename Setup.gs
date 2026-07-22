// ============= SETUP FUNCTIONS =============

function setupDemoData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timestamp = new Date().toISOString();

  // Step 1: Create a temporary sheet (required - can't delete all sheets)
  var tempSheet = ss.insertSheet('_TEMP_SETUP_');

  // Step 2: Delete all existing sheets except temp
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName() !== '_TEMP_SETUP_') {
      ss.deleteSheet(allSheets[i]);
    }
  }

  // Step 3: Create fresh sheets with demo data

  // ===== USERS SHEET =====
  var usersSheet = ss.insertSheet(CONFIG.SHEETS.USERS);
  usersSheet.appendRow(['ID', 'Username', 'Email', 'Password', 'Role', 'Status', 'CreatedAt', 'EmployeeID']);
  // Admin & User accounts
  usersSheet.appendRow([1, 'Admin', 'admin@demo.com', 'admin123', 'Admin', 'Active', timestamp, '']);
  usersSheet.appendRow([2, 'Manager', 'manager@demo.com', 'manager123', 'User', 'Active', timestamp, '']);
  // Employee Self-Service Accounts (linked to EmployeeID)
  usersSheet.appendRow([3, 'Employee 1', 'employee1@demo.com', 'emp123', 'Employee', 'Active', timestamp, 'EMP001']);
  usersSheet.appendRow([4, 'Employee 2', 'employee2@demo.com', 'emp123', 'Employee', 'Active', timestamp, 'EMP002']);
  usersSheet.appendRow([5, 'Employee 3', 'employee3@demo.com', 'emp123', 'Employee', 'Active', timestamp, 'EMP003']);

  // ===== EMPLOYEES SHEET =====
  var empSheet = ss.insertSheet(CONFIG.SHEETS.EMPLOYEES);
  empSheet.appendRow(['ID', 'EmployeeID', 'EmployeeName', 'Email', 'Position', 'BasicSalary', 'Status', 'JoinDate', 'CreatedAt', 'UpdatedAt']);
  empSheet.appendRow([1, 'EMP001', 'Employee 1', 'employee1@demo.com', 'Staff', 800, 'Active', new Date('2024-01-15').toISOString(), timestamp, timestamp]);
  empSheet.appendRow([2, 'EMP002', 'Employee 2', 'employee2@demo.com', 'Supervisor', 950, 'Active', new Date('2024-02-01').toISOString(), timestamp, timestamp]);
  empSheet.appendRow([3, 'EMP003', 'Employee 3', 'employee3@demo.com', 'Assistant', 700, 'Active', new Date('2024-03-10').toISOString(), timestamp, timestamp]);

  // ===== PayslipS SHEET =====
  var PayslipSheet = ss.insertSheet(CONFIG.SHEETS.PAYSLIPS);
  PayslipSheet.appendRow(['ID', 'EmployeeID', 'EmployeeName', 'PayMonth', 'PayYear', 'PayDate', 'PaidDays', 'LOPDays',
                          'Basic', 'ServiceCharge', 'Incentive', 'BillTips', 'SalesBonus', 'GrossEarnings',
                          'LOPDeduction', 'AdvancePayment', 'ShopCharges', 'TotalDeductions', 'NetPay',
                          'Status', 'CreatedBy', 'CreatedAt']);

  // Demo Payslips for Employee 1
  PayslipSheet.appendRow([
    1, 'EMP001', 'Employee 1', 'December', 2025, new Date('2025-12-31').toISOString(), 30, 0,
    800.00, 150.00, 50.00, 80.00, 0.00, 1080.00,
    0.00, 0.00, 100.00, 100.00, 980.00,
    'Paid', 'Admin', timestamp
  ]);
  PayslipSheet.appendRow([
    2, 'EMP001', 'Employee 1', 'January', 2026, new Date('2026-01-31').toISOString(), 30, 0,
    800.00, 160.00, 55.00, 90.00, 0.00, 1105.00,
    0.00, 0.00, 105.00, 105.00, 1000.00,
    'Paid', 'Admin', timestamp
  ]);

  // Demo Payslips for Employee 2
  PayslipSheet.appendRow([
    3, 'EMP002', 'Employee 2', 'December', 2025, new Date('2025-12-31').toISOString(), 30, 0,
    950.00, 180.00, 60.00, 100.00, 50.00, 1340.00,
    0.00, 0.00, 120.00, 120.00, 1220.00,
    'Paid', 'Admin', timestamp
  ]);
  PayslipSheet.appendRow([
    4, 'EMP002', 'Employee 2', 'January', 2026, new Date('2026-01-31').toISOString(), 28, 2,
    950.00, 175.00, 55.00, 95.00, 45.00, 1320.00,
    63.33, 0.00, 115.00, 178.33, 1141.67,
    'Pending', 'Admin', timestamp
  ]);

  // Demo Payslips for Employee 3
  PayslipSheet.appendRow([
    5, 'EMP003', 'Employee 3', 'December', 2025, new Date('2025-12-31').toISOString(), 30, 0,
    700.00, 120.00, 40.00, 60.00, 0.00, 920.00,
    0.00, 50.00, 80.00, 130.00, 790.00,
    'Paid', 'Admin', timestamp
  ]);
  PayslipSheet.appendRow([
    6, 'EMP003', 'Employee 3', 'January', 2026, new Date('2026-01-31').toISOString(), 30, 0,
    700.00, 125.00, 45.00, 65.00, 0.00, 935.00,
    0.00, 0.00, 85.00, 85.00, 850.00,
    'Pending', 'Admin', timestamp
  ]);

  // ===== LOGS SHEET =====
  var logsSheet = ss.insertSheet(CONFIG.SHEETS.LOGS);
  logsSheet.appendRow(['ID', 'Username', 'Action', 'Details', 'Timestamp']);
  logsSheet.appendRow([1, 'Sistem', 'Setup', 'Demo data initialized successfully', timestamp]);

  // Step 4: Delete the temporary sheet
  ss.deleteSheet(tempSheet);

  return { success: true, message: 'Demo data setup completed! All sheets recreated with fresh demo data.' };
}
