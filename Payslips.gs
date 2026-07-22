// ============= PayslipS Pengurusan =============

function getPayslips(user, role) {
  try {
    var PayslipSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PAYSLIPS);
    var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    var PayslipData = PayslipSheet.getDataRange().getValues();
    var empData = empSheet.getDataRange().getValues();

    // Create employee position lookup
    var employeePositions = {};
    for (var i = 1; i < empData.length; i++) {
      employeePositions[empData[i][1]] = empData[i][4]; // EmployeeID -> Position
    }

    var Payslips = [];

    for (var i = 1; i < PayslipData.length; i++) {
      var employeeID = PayslipData[i][1];
      Payslips.push({
        ID: PayslipData[i][0],
        EmployeeID: employeeID,
        EmployeeName: PayslipData[i][2],
        Position: employeePositions[employeeID] || '',
        PayMonth: PayslipData[i][3],
        PayYear: PayslipData[i][4],
        PayDate: PayslipData[i][5] ? (typeof PayslipData[i][5] === 'string' ? PayslipData[i][5] : new Date(PayslipData[i][5]).toISOString()) : '',
        PaidDays: PayslipData[i][6],
        LOPDays: PayslipData[i][7],
        Basic: PayslipData[i][8],
        ServiceCharge: PayslipData[i][9],
        Incentive: PayslipData[i][10],
        BillTips: PayslipData[i][11],
        SalesBonus: PayslipData[i][12],
        GrossEarnings: PayslipData[i][13],
        LOPDeduction: PayslipData[i][14],
        AdvancePayment: PayslipData[i][15],
        ShopCharges: PayslipData[i][16],
        TotalDeductions: PayslipData[i][17],
        NetPay: PayslipData[i][18],
        Status: PayslipData[i][19],
        CreatedBy: PayslipData[i][20],
        CreatedAt: new Date(PayslipData[i][21]).toISOString()
      });
    }

    return { success: true, data: Payslips };
  } catch (error) {
    return { success: false, message: 'Error loading Payslips: ' + error.message };
  }
}

function addPayslip(PayslipData, user) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PAYSLIPS);
    var data = sheet.getDataRange().getValues();

    // Check for duplicate Payslip (same employee, month, year)
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === PayslipData.EmployeeID &&
          data[i][3] === PayslipData.PayMonth &&
          data[i][4] == PayslipData.PayYear) {
        return {
          success: false,
          message: 'Duplicate Payslip! A Payslip for ' + PayslipData.EmployeeName + ' already exists for ' + PayslipData.PayMonth + ' ' + PayslipData.PayYear
        };
      }
    }

    var lastRow = sheet.getLastRow();
    var newId = lastRow > 0 ? sheet.getRange(lastRow, 1).getValue() + 1 : 1;

    // Calculate totals
    var grossEarnings = parseFloat(PayslipData.Basic) + parseFloat(PayslipData.ServiceCharge) +
                        parseFloat(PayslipData.Incentive) + parseFloat(PayslipData.BillTips) +
                        parseFloat(PayslipData.SalesBonus);

    var totalDeductions = parseFloat(PayslipData.LOPDeduction) + parseFloat(PayslipData.AdvancePayment) +
                          parseFloat(PayslipData.ShopCharges);

    var netPay = grossEarnings - totalDeductions;

    sheet.appendRow([
      newId,
      PayslipData.EmployeeID,
      PayslipData.EmployeeName,
      PayslipData.PayMonth,
      PayslipData.PayYear,
      PayslipData.PayDate ? new Date(PayslipData.PayDate).toISOString() : '',
      PayslipData.PaidDays,
      PayslipData.LOPDays,
      parseFloat(PayslipData.Basic),
      parseFloat(PayslipData.ServiceCharge),
      parseFloat(PayslipData.Incentive),
      parseFloat(PayslipData.BillTips),
      parseFloat(PayslipData.SalesBonus),
      grossEarnings,
      parseFloat(PayslipData.LOPDeduction),
      parseFloat(PayslipData.AdvancePayment),
      parseFloat(PayslipData.ShopCharges),
      totalDeductions,
      netPay,
      PayslipData.Status || 'Pending',
      user,
      new Date().toISOString()
    ]);

    logActivity(user, 'Add Payslip', 'Created Payslip for: ' + PayslipData.EmployeeName);
    return { success: true, message: 'Payslip created successfully', netPay: netPay };
  } catch (error) {
    return { success: false, message: 'Error creating Payslip: ' + error.message };
  }
}

function updatePayslip(id, PayslipData, user, role) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PAYSLIPS);
    var data = sheet.getDataRange().getValues();

    // Calculate totals
    var grossEarnings = parseFloat(PayslipData.Basic) + parseFloat(PayslipData.ServiceCharge) +
                        parseFloat(PayslipData.Incentive) + parseFloat(PayslipData.BillTips) +
                        parseFloat(PayslipData.SalesBonus);

    var totalDeductions = parseFloat(PayslipData.LOPDeduction) + parseFloat(PayslipData.AdvancePayment) +
                          parseFloat(PayslipData.ShopCharges);

    var netPay = grossEarnings - totalDeductions;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.getRange(i + 1, 2).setValue(PayslipData.EmployeeID);
        sheet.getRange(i + 1, 3).setValue(PayslipData.EmployeeName);
        sheet.getRange(i + 1, 4).setValue(PayslipData.PayMonth);
        sheet.getRange(i + 1, 5).setValue(PayslipData.PayYear);
        sheet.getRange(i + 1, 6).setValue(PayslipData.PayDate ? new Date(PayslipData.PayDate).toISOString() : '');
        sheet.getRange(i + 1, 7).setValue(PayslipData.PaidDays);
        sheet.getRange(i + 1, 8).setValue(PayslipData.LOPDays);
        sheet.getRange(i + 1, 9).setValue(parseFloat(PayslipData.Basic));
        sheet.getRange(i + 1, 10).setValue(parseFloat(PayslipData.ServiceCharge));
        sheet.getRange(i + 1, 11).setValue(parseFloat(PayslipData.Incentive));
        sheet.getRange(i + 1, 12).setValue(parseFloat(PayslipData.BillTips));
        sheet.getRange(i + 1, 13).setValue(parseFloat(PayslipData.SalesBonus));
        sheet.getRange(i + 1, 14).setValue(grossEarnings);
        sheet.getRange(i + 1, 15).setValue(parseFloat(PayslipData.LOPDeduction));
        sheet.getRange(i + 1, 16).setValue(parseFloat(PayslipData.AdvancePayment));
        sheet.getRange(i + 1, 17).setValue(parseFloat(PayslipData.ShopCharges));
        sheet.getRange(i + 1, 18).setValue(totalDeductions);
        sheet.getRange(i + 1, 19).setValue(netPay);
        sheet.getRange(i + 1, 20).setValue(PayslipData.Status);

        logActivity(user, 'Update Payslip', 'Updated Payslip for: ' + PayslipData.EmployeeName);
        return { success: true, message: 'Payslip updated successfully' };
      }
    }

    return { success: false, message: 'Payslip not found' };
  } catch (error) {
    return { success: false, message: 'Error updating Payslip: ' + error.message };
  }
}

function updatePayslipStatus(id, newStatus, user, role) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PAYSLIPS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        // Column 20 is Status (0-indexed: 19)
        sheet.getRange(i + 1, 20).setValue(newStatus);

        logActivity(user, 'Update Payslip Status', 'Updated Payslip ID ' + id + ' status to ' + newStatus);
        return { success: true, message: 'Status updated to ' + newStatus };
      }
    }

    return { success: false, message: 'Payslip not found' };
  } catch (error) {
    return { success: false, message: 'Error updating status: ' + error.message };
  }
}

function deletePayslip(id, user, role) {
  try {
    if (role !== 'Admin') {
      return { success: false, message: 'Only admins can delete Payslips' };
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PAYSLIPS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        logActivity(user, 'Delete Payslip', 'Deleted Payslip ID: ' + id);
        return { success: true, message: 'Payslip deleted successfully' };
      }
    }

    return { success: false, message: 'Payslip not found' };
  } catch (error) {
    return { success: false, message: 'Error deleting Payslip: ' + error.message };
  }
}

function getPayslipByEmployeeID(employeeID) {
  try {
    var PayslipSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PAYSLIPS);
    var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    var PayslipData = PayslipSheet.getDataRange().getValues();
    var empData = empSheet.getDataRange().getValues();

    // Get employee position
    var position = '';
    for (var j = 1; j < empData.length; j++) {
      if (empData[j][1] === employeeID) {
        position = empData[j][4];
        break;
      }
    }

    var Payslips = [];

    for (var i = 1; i < PayslipData.length; i++) {
      if (PayslipData[i][1] === employeeID) {
        Payslips.push({
          ID: PayslipData[i][0],
          EmployeeID: PayslipData[i][1],
          EmployeeName: PayslipData[i][2],
          Position: position,
          PayMonth: PayslipData[i][3],
          PayYear: PayslipData[i][4],
          PayDate: PayslipData[i][5] ? (typeof PayslipData[i][5] === 'string' ? PayslipData[i][5] : new Date(PayslipData[i][5]).toISOString()) : '',
          PaidDays: PayslipData[i][6],
          LOPDays: PayslipData[i][7],
          Basic: PayslipData[i][8],
          ServiceCharge: PayslipData[i][9],
          Incentive: PayslipData[i][10],
          BillTips: PayslipData[i][11],
          SalesBonus: PayslipData[i][12],
          GrossEarnings: PayslipData[i][13],
          LOPDeduction: PayslipData[i][14],
          AdvancePayment: PayslipData[i][15],
          ShopCharges: PayslipData[i][16],
          TotalDeductions: PayslipData[i][17],
          NetPay: PayslipData[i][18],
          Status: PayslipData[i][19]
        });
      }
    }

    if (Payslips.length > 0) {
      return { success: true, data: Payslips };
    }

    return { success: false, message: 'No Payslips found for this Employee ID' };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.message };
  }
}

// ============= MONTHLY REPORT =============

function getMonthlyReport(month, year) {
  try {
    var PayslipSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PAYSLIPS);
    var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    var PayslipData = PayslipSheet.getDataRange().getValues();
    var empData = empSheet.getDataRange().getValues();

    // Create employee position lookup
    var employeePositions = {};
    for (var i = 1; i < empData.length; i++) {
      employeePositions[empData[i][1]] = empData[i][4];
    }

    var reportData = [];
    var totalGross = 0;
    var totalDeductions = 0;
    var totalNet = 0;
    var paidCount = 0;
    var pendingCount = 0;

    for (var i = 1; i < PayslipData.length; i++) {
      if (PayslipData[i][3] === month && PayslipData[i][4] == year) {
        var Payslip = {
          ID: PayslipData[i][0],
          EmployeeID: PayslipData[i][1],
          EmployeeName: PayslipData[i][2],
          Position: employeePositions[PayslipData[i][1]] || '',
          Basic: PayslipData[i][8],
          ServiceCharge: PayslipData[i][9],
          Incentive: PayslipData[i][10],
          BillTips: PayslipData[i][11],
          SalesBonus: PayslipData[i][12],
          GrossEarnings: PayslipData[i][13],
          LOPDeduction: PayslipData[i][14],
          AdvancePayment: PayslipData[i][15],
          ShopCharges: PayslipData[i][16],
          TotalDeductions: PayslipData[i][17],
          NetPay: PayslipData[i][18],
          Status: PayslipData[i][19]
        };

        reportData.push(Payslip);
        totalGross += Payslip.GrossEarnings;
        totalDeductions += Payslip.TotalDeductions;
        totalNet += Payslip.NetPay;

        if (Payslip.Status === 'Paid') paidCount++;
        else pendingCount++;
      }
    }

    return {
      success: true,
      data: reportData,
      summary: {
        month: month,
        year: year,
        totalEmployees: reportData.length,
        totalGrossEarnings: totalGross.toFixed(2),
        totalDeductions: totalDeductions.toFixed(2),
        totalNetPay: totalNet.toFixed(2),
        paidCount: paidCount,
        pendingCount: pendingCount
      }
    };
  } catch (error) {
    return { success: false, message: 'Error generating report: ' + error.message };
  }
}
