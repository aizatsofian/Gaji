// ============= EMPLOYEE Payslip HISTORY =============

function getEmployeePayslipHistory(employeeID) {
  try {
    var PayslipSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PAYSLIPS);
    var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    var PayslipData = PayslipSheet.getDataRange().getValues();
    var empData = empSheet.getDataRange().getValues();

    // Get employee info
    var employeeInfo = null;
    for (var i = 1; i < empData.length; i++) {
      if (empData[i][1] === employeeID) {
        employeeInfo = {
          EmployeeID: empData[i][1],
          EmployeeName: empData[i][2],
          Email: empData[i][3],
          Position: empData[i][4],
          BasicSalary: empData[i][5],
          Status: empData[i][6],
          JoinDate: empData[i][7]
        };
        break;
      }
    }

    if (!employeeInfo) {
      return { success: false, message: 'Employee not found' };
    }

    var history = [];
    var totalEarned = 0;

    for (var i = 1; i < PayslipData.length; i++) {
      if (PayslipData[i][1] === employeeID) {
        var Payslip = {
          ID: PayslipData[i][0],
          PayMonth: PayslipData[i][3],
          PayYear: PayslipData[i][4],
          PayDate: PayslipData[i][5] ? (typeof PayslipData[i][5] === 'string' ? PayslipData[i][5] : new Date(PayslipData[i][5]).toISOString()) : '',
          GrossEarnings: PayslipData[i][13],
          TotalDeductions: PayslipData[i][17],
          NetPay: PayslipData[i][18],
          Status: PayslipData[i][19]
        };
        history.push(Payslip);
        totalEarned += Payslip.NetPay;
      }
    }

    // Sort by year and month (newest first)
    history.sort(function(a, b) {
      if (b.PayYear !== a.PayYear) return b.PayYear - a.PayYear;
      var months = ['January', 'February', 'March', 'April', 'May', 'June',
                    'July', 'August', 'September', 'October', 'November', 'December'];
      return months.indexOf(b.PayMonth) - months.indexOf(a.PayMonth);
    });

    return {
      success: true,
      employee: employeeInfo,
      history: history,
      totalPayslips: history.length,
      totalEarned: totalEarned.toFixed(2)
    };
  } catch (error) {
    return { success: false, message: 'Error loading history: ' + error.message };
  }
}

// ============= EMPLOYEE SELF-SERVICE PORTAL =============

/**
 * Get employee's own profile data
 */
function getEmployeeOwnData(employeeID) {
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

    return { success: false, message: 'Employee not found' };
  } catch (error) {
    return { success: false, message: 'Error loading employee data: ' + error.message };
  }
}

/**
 * Get employee's own Payslips (for Employee Self-Service Portal)
 */
function getEmployeeOwnPayslips(employeeID) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PAYSLIPS);
    var data = sheet.getDataRange().getValues();
    var Payslips = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === employeeID) {
        Payslips.push({
          ID: data[i][0],
          EmployeeID: data[i][1],
          EmployeeName: data[i][2],
          PayMonth: data[i][3],
          PayYear: data[i][4],
          PayDate: data[i][5] ? (typeof data[i][5] === 'string' ? data[i][5] : new Date(data[i][5]).toISOString()) : '',
          PaidDays: data[i][6],
          LOPDays: data[i][7],
          Basic: data[i][8],
          ServiceCharge: data[i][9],
          Incentive: data[i][10],
          BillTips: data[i][11],
          SalesBonus: data[i][12],
          GrossEarnings: data[i][13],
          LOPDeduction: data[i][14],
          AdvancePayment: data[i][15],
          ShopCharges: data[i][16],
          TotalDeductions: data[i][17],
          NetPay: data[i][18],
          Status: data[i][19]
        });
      }
    }

    // Sort by year and month (newest first)
    Payslips.sort(function(a, b) {
      if (b.PayYear !== a.PayYear) return b.PayYear - a.PayYear;
      var months = ['January', 'February', 'March', 'April', 'May', 'June',
                    'July', 'August', 'September', 'October', 'November', 'December'];
      return months.indexOf(b.PayMonth) - months.indexOf(a.PayMonth);
    });

    return { success: true, data: Payslips };
  } catch (error) {
    return { success: false, message: 'Error loading Payslips: ' + error.message };
  }
}
