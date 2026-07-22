// ============= DASHBOARD STATS =============

function getDashboardStats() {
  try {
    var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    var PayslipSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PAYSLIPS);

    var empData = empSheet.getDataRange().getValues();
    var PayslipData = PayslipSheet.getDataRange().getValues();

    var totalEmployees = empData.length - 1;
    var activeEmployees = 0;
    var inactiveEmployees = 0;

    for (var i = 1; i < empData.length; i++) {
      if (empData[i][6] === 'Active') activeEmployees++;
      else inactiveEmployees++;
    }

    var totalPayslips = PayslipData.length - 1;
    var totalPayout = 0;
    var paidPayslips = 0;
    var pendingPayslips = 0;

    for (var i = 1; i < PayslipData.length; i++) {
      totalPayout += PayslipData[i][18];
      if (PayslipData[i][19] === 'Paid') paidPayslips++;
      else pendingPayslips++;
    }

    return {
      success: true,
      totalEmployees: totalEmployees,
      activeEmployees: activeEmployees,
      inactiveEmployees: inactiveEmployees,
      totalPayslips: totalPayslips,
      totalPayout: totalPayout.toFixed(2),
      paidPayslips: paidPayslips,
      pendingPayslips: pendingPayslips
    };
  } catch (error) {
    return { success: false, message: 'Error loading stats: ' + error.message };
  }
}

/**
 * Get dashboard statistics for an employee
 */
function getEmployeeDashboardStats(employeeID) {
  try {
    var PayslipSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PAYSLIPS);
    var PayslipData = PayslipSheet.getDataRange().getValues();

    var totalPayslips = 0;
    var paidPayslips = 0;
    var pendingPayslips = 0;
    var totalEarnings = 0;
    var lastPayslip = null;

    for (var i = 1; i < PayslipData.length; i++) {
      if (PayslipData[i][1] === employeeID) {
        totalPayslips++;
        if (PayslipData[i][19] === 'Paid') {
          paidPayslips++;
          totalEarnings += parseFloat(PayslipData[i][18]) || 0;
        } else {
          pendingPayslips++;
        }

        // Track latest Payslip
        if (!lastPayslip || PayslipData[i][4] > lastPayslip.PayYear ||
            (PayslipData[i][4] === lastPayslip.PayYear &&
             ['January', 'February', 'March', 'April', 'May', 'June',
              'July', 'August', 'September', 'October', 'November', 'December'].indexOf(PayslipData[i][3]) >
             ['January', 'February', 'March', 'April', 'May', 'June',
              'July', 'August', 'September', 'October', 'November', 'December'].indexOf(lastPayslip.PayMonth))) {
          lastPayslip = {
            PayMonth: PayslipData[i][3],
            PayYear: PayslipData[i][4],
            NetPay: PayslipData[i][18],
            Status: PayslipData[i][19]
          };
        }
      }
    }

    // Get employee info
    var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    var empData = empSheet.getDataRange().getValues();
    var employeeInfo = null;

    for (var j = 1; j < empData.length; j++) {
      if (empData[j][1] === employeeID) {
        employeeInfo = {
          EmployeeID: empData[j][1],
          EmployeeName: empData[j][2],
          Position: empData[j][4],
          BasicSalary: empData[j][5]
        };
        break;
      }
    }

    return {
      success: true,
      employee: employeeInfo,
      stats: {
        totalPayslips: totalPayslips,
        paidPayslips: paidPayslips,
        pendingPayslips: pendingPayslips,
        totalEarnings: totalEarnings.toFixed(2),
        lastPayslip: lastPayslip
      }
    };
  } catch (error) {
    return { success: false, message: 'Error loading dashboard: ' + error.message };
  }
}
