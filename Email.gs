function getCompanyInfo() {
  return {
    success: true,
    data: CONFIG.COMPANY
  };
}

// ============= EMAIL Payslip =============

function sendPayslipEmail(PayslipId, user, role) {
  try {
    if (role !== 'Admin') {
      return { success: false, message: 'Only admins can send Payslip emails' };
    }

    var PayslipSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PAYSLIPS);
    var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    var data = PayslipSheet.getDataRange().getValues();

    // Find the Payslip
    var Payslip = null;
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == PayslipId) {
        Payslip = {
          ID: data[i][0],
          EmployeeID: data[i][1],
          EmployeeName: data[i][2],
          PayMonth: data[i][3],
          PayYear: data[i][4],
          PayDate: data[i][5],
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
        };
        rowIndex = i;
        break;
      }
    }

    if (!Payslip) {
      return { success: false, message: 'Payslip not found' };
    }

    // Get employee email and position
    var empData = empSheet.getDataRange().getValues();
    var employeeEmail = '';
    var position = '';
    for (var i = 1; i < empData.length; i++) {
      if (empData[i][1] === Payslip.EmployeeID) {
        employeeEmail = empData[i][3];
        position = empData[i][4];
        break;
      }
    }

    if (!employeeEmail) {
      return { success: false, message: 'Employee email not found' };
    }

    // Add Position to Payslip object
    Payslip.Position = position;

    // Format date for display
    var payDate = new Date(Payslip.PayDate);
    var formattedDate = Utilities.formatDate(payDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');

    // Generate HTML email content
    var htmlBody = generatePayslipHTML(Payslip, formattedDate);

    // Send email
    var subject = 'Payslip for ' + Payslip.PayMonth + ' ' + Payslip.PayYear + ' - ' + Payslip.EmployeeName;

    MailApp.sendEmail({
      to: employeeEmail,
      subject: subject,
      htmlBody: htmlBody
    });

    // Log activity
    logActivity(user, 'Send Payslip Email', 'Sent Payslip email to ' + Payslip.EmployeeName + ' (' + employeeEmail + ')');

    return {
      success: true,
      message: 'Payslip sent successfully to ' + employeeEmail
    };

  } catch (error) {
    return { success: false, message: 'Error sending email: ' + error.message };
  }
}

function generatePayslipHTML(Payslip, formattedDate) {
  var html = '<!DOCTYPE html>';
  html += '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">';
  html += '<style>';

  // Email wrapper styles
  html += 'body { font-family: -apple-Sistem, BlinkMacSistemFont, "Segoe UI", Roboto, sans-serif; margin: 0; padding: 0; background: #f0f2f5; }';
  html += '.email-wrapper { max-width: 700px; margin: 0 auto; background: #f0f2f5; padding: 30px 20px; }';

  // Email header styles
  html += '.email-header { background: linear-gradient(135deg, #001f3f 0%, #003366 100%); padding: 35px 30px; border-radius: 12px 12px 0 0; text-align: center; }';
  html += '.email-logo { width: 80px; height: 80px; border-radius: 50%; border: 3px solid rgba(255,255,255,0.3); margin-bottom: 15px; }';
  html += '.email-header h1 { color: white; font-size: 24px; margin: 0 0 5px 0; font-weight: 700; }';
  html += '.email-header p { color: rgba(255,255,255,0.8); font-size: 14px; margin: 0; }';

  // Email body styles
  html += '.email-body { background: white; padding: 35px 30px; }';
  html += '.greeting { font-size: 18px; color: #001f3f; margin-bottom: 15px; font-weight: 600; }';
  html += '.intro-text { font-size: 15px; color: #555; line-height: 1.7; margin-bottom: 25px; }';

  // Quick summary card
  html += '.quick-summary { background: linear-gradient(135deg, #e8f4fd 0%, #d4ecfa 100%); border-radius: 10px; padding: 25px; margin-bottom: 30px; border-left: 5px solid #0074D9; }';
  html += '.quick-summary-title { font-size: 12px; color: #666; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 15px; font-weight: 600; }';
  html += '.quick-summary-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }';
  html += '.summary-item { }';
  html += '.summary-item .label { font-size: 12px; color: #666; margin-bottom: 3px; }';
  html += '.summary-item .value { font-size: 16px; color: #001f3f; font-weight: 700; }';
  html += '.net-pay-highlight { background: linear-gradient(135deg, #001f3f 0%, #003366 100%); border-radius: 10px; padding: 20px; text-align: center; margin-top: 15px; }';
  html += '.net-pay-highlight .label { font-size: 11px; color: rgba(255,255,255,0.7); text-transform: uppercase; letter-spacing: 1px; margin-bottom: 5px; }';
  html += '.net-pay-highlight .amount { font-size: 36px; color: white; font-weight: 700; }';

  // Payslip container styles
  html += '.Payslip-section-title { font-size: 14px; color: #001f3f; font-weight: 700; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #e0e0e0; display: flex; align-items: center; gap: 8px; }';
  html += '.Payslip-container { background: #fafbfc; border: 1px solid #e5e5e5; border-radius: 10px; padding: 25px; margin-bottom: 25px; position: relative; overflow: hidden; }';
  html += '.Payslip-watermark { position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); opacity: 0.03; width: 300px; height: 300px; background-image: url(' + CONFIG.COMPANY.WATERMARK + '); background-size: contain; background-repeat: no-repeat; background-position: center; z-index: 0; pointer-events: none; }';
  html += '.Payslip-content { position: relative; z-index: 1; }';

  // Employee info
  html += '.employee-info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 20px; }';
  html += '.info-item { font-size: 13px; }';
  html += '.info-item .label { color: #888; }';
  html += '.info-item .value { color: #333; font-weight: 600; }';

  // Earnings/Deductions table
  html += '.ed-table { width: 100%; border-collapse: collapse; font-size: 13px; }';
  html += '.ed-table th { background: #001f3f; color: white; padding: 12px 15px; text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; }';
  html += '.ed-table th:last-child { text-align: right; }';
  html += '.ed-table td { padding: 10px 15px; border-bottom: 1px solid #eee; }';
  html += '.ed-table td:last-child { text-align: right; font-weight: 600; }';
  html += '.ed-table .total-row { background: #f5f5f5; font-weight: 700; }';
  html += '.ed-table .total-row td { border-top: 2px solid #001f3f; color: #001f3f; }';

  // Net payable bar
  html += '.net-payable-bar { background: linear-gradient(135deg, #27ae60 0%, #2ecc71 100%); padding: 18px 20px; border-radius: 8px; display: flex; justify-content: space-between; align-items: center; margin-top: 20px; }';
  html += '.net-payable-bar .label { color: white; font-size: 14px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }';
  html += '.net-payable-bar .amount { color: white; font-size: 26px; font-weight: 700; }';

  // CTA Button
  html += '.cta-section { text-align: center; margin: 30px 0; }';
  html += '.cta-text { font-size: 14px; color: #666; margin-bottom: 15px; }';
  html += '.cta-button { display: inline-block; background: linear-gradient(135deg, #001f3f 0%, #003366 100%); color: white; text-decoration: none; padding: 14px 35px; border-radius: 8px; font-size: 14px; font-weight: 600; }';

  // Email footer
  html += '.email-footer { background: #001f3f; padding: 30px; border-radius: 0 0 12px 12px; text-align: center; }';
  html += '.footer-logo { width: 50px; height: 50px; border-radius: 50%; margin-bottom: 15px; border: 2px solid rgba(255,255,255,0.2); }';
  html += '.footer-company { color: white; font-size: 16px; font-weight: 700; margin-bottom: 5px; }';
  html += '.footer-tagline { color: rgba(255,255,255,0.6); font-size: 12px; margin-bottom: 20px; }';
  html += '.footer-contact { margin-bottom: 20px; }';
  html += '.footer-contact a { color: rgba(255,255,255,0.8); text-decoration: none; font-size: 13px; margin: 0 10px; }';
  html += '.social-links { margin-bottom: 20px; }';
  html += '.social-link { display: inline-block; width: 36px; height: 36px; background: rgba(255,255,255,0.1); border-radius: 50%; margin: 0 5px; line-height: 36px; color: white; text-decoration: none; font-size: 14px; }';
  html += '.footer-disclaimer { color: rgba(255,255,255,0.5); font-size: 11px; line-height: 1.6; border-top: 1px solid rgba(255,255,255,0.1); padding-top: 20px; }';

  html += '</style></head><body>';

  // Email Wrapper Start
  html += '<div class="email-wrapper">';

  // Email Header
  html += '<div class="email-header">';
  html += '<img src="' + CONFIG.COMPANY.LOGO + '" alt="Logo" class="email-logo">';
  html += '<h1>' + CONFIG.COMPANY.NAME + '</h1>';
  html += '<p>' + CONFIG.COMPANY.SUBTITLE + '</p>';
  html += '</div>';

  // Email Body
  html += '<div class="email-body">';

  // Greeting
  html += '<div class="greeting">Hello ' + Payslip.EmployeeName.split(' ')[0] + ',</div>';
  html += '<p class="intro-text">Your Payslip for <strong>' + Payslip.PayMonth + ' ' + Payslip.PayYear + '</strong> is now available. Below is a summary of your earnings and deductions for this pay period. Please review the details carefully.</p>';

  // Quick Summary Card
  html += '<div class="quick-summary">';
  html += '<div class="quick-summary-title">Payment Summary</div>';
  html += '<div class="quick-summary-grid">';
  html += '<div class="summary-item"><div class="label">Pay Period</div><div class="value">' + Payslip.PayMonth + ' ' + Payslip.PayYear + '</div></div>';
  html += '<div class="summary-item"><div class="label">Pay Date</div><div class="value">' + formattedDate + '</div></div>';
  html += '<div class="summary-item"><div class="label">Gross Earnings</div><div class="value">$' + parseFloat(Payslip.GrossEarnings).toFixed(2) + '</div></div>';
  html += '<div class="summary-item"><div class="label">Total Deductions</div><div class="value">$' + parseFloat(Payslip.TotalDeductions).toFixed(2) + '</div></div>';
  html += '</div>';
  html += '<div class="net-pay-highlight">';
  html += '<div class="label">Your Net Pay</div>';
  html += '<div class="amount">$' + Payslip.NetPay.toFixed(2) + '</div>';
  html += '</div>';
  html += '</div>';

  // Detailed Payslip Section
  html += '<div class="Payslip-section-title">Detailed Payslip</div>';
  html += '<div class="Payslip-container">';
  html += '<div class="Payslip-watermark"></div>';
  html += '<div class="Payslip-content">';

  // Employee Info Grid
  html += '<div class="employee-info-grid">';
  html += '<div class="info-item"><span class="label">Employee: </span><span class="value">' + Payslip.EmployeeName + '</span></div>';
  html += '<div class="info-item"><span class="label">ID: </span><span class="value">' + Payslip.EmployeeID + '</span></div>';
  html += '<div class="info-item"><span class="label">Position: </span><span class="value">' + (Payslip.Position || 'N/A') + '</span></div>';
  html += '<div class="info-item"><span class="label">Paid Days: </span><span class="value">' + Payslip.PaidDays + ' days</span></div>';
  html += '</div>';

  // Earnings Table
  html += '<table class="ed-table" style="margin-bottom: 20px;">';
  html += '<thead><tr><th colspan="2">Earnings</th></tr></thead>';
  html += '<tbody>';
  html += '<tr><td>Basic Salary</td><td>$' + parseFloat(Payslip.Basic).toFixed(2) + '</td></tr>';
  html += '<tr><td>Service Charge</td><td>$' + parseFloat(Payslip.ServiceCharge).toFixed(2) + '</td></tr>';
  html += '<tr><td>Incentive</td><td>$' + parseFloat(Payslip.Incentive).toFixed(2) + '</td></tr>';
  html += '<tr><td>Bill Tips</td><td>$' + parseFloat(Payslip.BillTips).toFixed(2) + '</td></tr>';
  html += '<tr><td>Sales Bonus</td><td>$' + parseFloat(Payslip.SalesBonus).toFixed(2) + '</td></tr>';
  html += '<tr class="total-row"><td>Gross Earnings</td><td>$' + parseFloat(Payslip.GrossEarnings).toFixed(2) + '</td></tr>';
  html += '</tbody></table>';

  // Deductions Table
  html += '<table class="ed-table">';
  html += '<thead><tr><th colspan="2">Deductions</th></tr></thead>';
  html += '<tbody>';
  html += '<tr><td>LOP Deduction (' + Payslip.LOPDays + ' days)</td><td>$' + parseFloat(Payslip.LOPDeduction).toFixed(2) + '</td></tr>';
  html += '<tr><td>Advance Payment</td><td>$' + parseFloat(Payslip.AdvancePayment).toFixed(2) + '</td></tr>';
  html += '<tr><td>Shop Charges</td><td>$' + parseFloat(Payslip.ShopCharges).toFixed(2) + '</td></tr>';
  html += '<tr class="total-row"><td>Total Deductions</td><td>$' + parseFloat(Payslip.TotalDeductions).toFixed(2) + '</td></tr>';
  html += '</tbody></table>';

  // Net Payable Bar
  html += '<div class="net-payable-bar">';
  html += '<span class="label">Net Payable Amount</span>';
  html += '<span class="amount">$' + Payslip.NetPay.toFixed(2) + '</span>';
  html += '</div>';

  html += '</div></div>'; // Close Payslip-content and Payslip-container

  // CTA Section
  html += '<div class="cta-section">';
  html += '<p class="cta-text">Have questions about your Payslip? Contact HR for assistance.</p>';
  html += '</div>';

  // Disclaimer
  html += '<p style="font-size: 12px; color: #888; line-height: 1.6; margin-top: 20px; padding: 15px; background: #f9f9f9; border-radius: 6px;">This is a computer-generated Payslip and does not require a signature. Please review your Payslip carefully. Any discrepancies should be reported to the HR department within 7 working days.</p>';

  html += '</div>'; // Close email-body

  // Email Footer
  html += '<div class="email-footer">';
  html += '<img src="' + CONFIG.COMPANY.LOGO + '" alt="Logo" class="footer-logo">';
  html += '<div class="footer-company">' + CONFIG.COMPANY.NAME + '</div>';
  html += '<div class="footer-tagline">' + CONFIG.COMPANY.SUBTITLE + '</div>';
  html += '<div class="footer-contact">';
  html += '<a href="mailto:' + CONFIG.COMPANY.EMAIL + '">' + CONFIG.COMPANY.EMAIL + '</a>';
  html += '<a href="tel:' + CONFIG.COMPANY.PHONE + '">' + CONFIG.COMPANY.PHONE + '</a>';
  html += '</div>';
  html += '<div class="social-links">';
  html += '<a href="https://wa.me/923224083545" class="social-link" title="WhatsApp">W</a>';
  html += '<a href="https://www.youtube.com/@rameezimdad" class="social-link" title="YouTube">Y</a>';
  html += '</div>';
  html += '<div class="footer-disclaimer">';
  html += 'This email and any attachments are confidential and intended solely for the addressee. If you have received this email in error, please notify the sender immediately and delete this email.<br><br>';
  html += '&copy; ' + new Date().getFullYear() + ' ' + CONFIG.COMPANY.NAME + '. All rights reserved.';
  html += '</div>';
  html += '</div>';

  html += '</div></body></html>'; // Close email-wrapper

  return html;
}

// ============= BULK EMAIL PayslipS =============

function sendBulkPayslipEmails(month, year, user, role) {
  try {
    if (role !== 'Admin') {
      return { success: false, message: 'Only admins can send bulk emails' };
    }

    var PayslipSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PAYSLIPS);
    var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.EMPLOYEES);
    var PayslipData = PayslipSheet.getDataRange().getValues();
    var empData = empSheet.getDataRange().getValues();

    // Create employee email lookup
    var employeeEmails = {};
    var employeePositions = {};
    for (var i = 1; i < empData.length; i++) {
      employeeEmails[empData[i][1]] = empData[i][3]; // EmployeeID -> Email
      employeePositions[empData[i][1]] = empData[i][4]; // EmployeeID -> Position
    }

    var sentCount = 0;
    var failedCount = 0;
    var failedNames = [];

    for (var i = 1; i < PayslipData.length; i++) {
      if (PayslipData[i][3] === month && PayslipData[i][4] == year) {
        var Payslip = {
          ID: PayslipData[i][0],
          EmployeeID: PayslipData[i][1],
          EmployeeName: PayslipData[i][2],
          Position: employeePositions[PayslipData[i][1]] || '',
          PayMonth: PayslipData[i][3],
          PayYear: PayslipData[i][4],
          PayDate: PayslipData[i][5],
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
          NetPay: PayslipData[i][18]
        };

        var employeeEmail = employeeEmails[Payslip.EmployeeID];

        if (employeeEmail) {
          try {
            var payDate = new Date(Payslip.PayDate);
            var formattedDate = Utilities.formatDate(payDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
            var htmlBody = generatePayslipHTML(Payslip, formattedDate);
            var subject = 'Payslip for ' + Payslip.PayMonth + ' ' + Payslip.PayYear + ' - ' + Payslip.EmployeeName;

            MailApp.sendEmail({
              to: employeeEmail,
              subject: subject,
              htmlBody: htmlBody
            });
            sentCount++;
          } catch (emailError) {
            failedCount++;
            failedNames.push(Payslip.EmployeeName);
          }
        } else {
          failedCount++;
          failedNames.push(Payslip.EmployeeName + ' (no email)');
        }
      }
    }

    logActivity(user, 'Bulk Email', 'Sent ' + sentCount + ' Payslip emails for ' + month + ' ' + year);

    if (sentCount === 0) {
      return { success: false, message: 'No Payslips found for ' + month + ' ' + year };
    }

    var message = 'Successfully sent ' + sentCount + ' email(s)';
    if (failedCount > 0) {
      message += '. Failed: ' + failedCount + ' (' + failedNames.join(', ') + ')';
    }

    return { success: true, message: message, sent: sentCount, failed: failedCount };
  } catch (error) {
    return { success: false, message: 'Error sending bulk emails: ' + error.message };
  }
}
