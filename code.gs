/**
 * Dibina oleh Dr Excel
 * WhatsApp: https://wa.me/60102314951 (Untuk Custom Projek)
 * YouTube: https://www.youtube.com/@draizatexcel (Subscribe Tau!)
 */

// Configuration
var CONFIG = {
  SHEETS: {
    USERS: 'Users',
    EMPLOYEES: 'Employees',
    PAYSLIPS: 'Payslips',
    LOGS: 'Logs'
  },
  COMPANY: {
    NAME: 'Dr Excel',
    SUBTITLE: 'Sistem Pengurusan Gaji',
    EMAIL: 'contact@rameezscripts.com',
    PHONE: '+92 322 4083545',
    LOGO: 'https://raw.githubusercontent.com/aizatsofian/peroduaimage/main/LOGO.jpg',
    WATERMARK: 'https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEhebmAUcYFeNJmlcMeG2LlswzbwmFOz4UkaekSdNaEo60TZKtE0uzXzfJGEEe7kT01PkBJoSe6zQrYQZrttYJ1H6Eczq8yl-PWqsF63i-t6ZbHKfPPbUNaZt5h5iVg5nbaeIpSu4XdRuiS4SDiIFXRPZduFIFz_0LDdwGEis9h-y7jpq9MMDwKInxYFcHTd/s920/c9872608-1902-4385-9a69-bba7de71fa08.jpg'
  }
};

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  return template
    .evaluate()
    .setTitle('Sistem Pengurusan Gaji')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Used by index.html to pull in Stylesheet/JS partials via <?!= include('Filename'); ?>
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
