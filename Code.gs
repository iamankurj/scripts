function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom')
      .addItem('Generate PDF', 'generateInvoice')
      .addToUi();
}

function generateInvoice() {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = activeSheet.getSheetByName("Template");

  var currentDate = new Date();

  updateInvoice(templateSheet, currentDate);

  var pdfFile = generateAndSavePDFFile(activeSheet, templateSheet, currentDate);

  // Update ledger and next invoice number
  var trackingSheet = activeSheet.getSheetByName("Tracking");
  var grossAmount = templateSheet.getRange("F25").getValue();
  updateLedger(activeSheet, trackingSheet, currentDate, grossAmount);
  updateNextInvoiceNumber(trackingSheet);

  // Show success message with link to PDF file
  showLink(pdfFile);
}

function updateInvoice(templateSheet, currentDate) {
  // Update invoice date
  var invoiceDateCell = templateSheet.getRange("G11");
  var invoiceDateString = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
  invoiceDateCell.setValue(invoiceDateString);

  // Wait for changes to be applied
  SpreadsheetApp.flush();
  do {
    Utilities.sleep(500);
  } while(invoiceDateCell.getValue() != invoiceDateString);
}

function generateAndSavePDFFile(activeSheet, templateSheet, currentDate) {
  var url = activeSheet.getUrl().replace(/edit$/, '') + 'export?exportFormat=pdf&format=pdf' +
    '&size=letter' +
    '&portrait=true' +
    '&fitw=true' +
    '&sheetnames=false&printtitle=false' +
    '&pagenumbers=false&gridlines=false' +
    '&fzr=false' +
    '&gid=' + templateSheet.getSheetId();
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });

  // Set name of PDF file to current month
  var invoiceNumber = templateSheet.getRange("G10").getValue();
  var clientNameCell = templateSheet.getRange("B11");
  var clientName = clientNameCell.getValue();

  var currentDateString = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyyMMdd");
  
  var blob = response.getBlob().setName('Invoice ' + invoiceNumber.toString() + ' - ' + clientName + ' - ' + currentDateString + '.pdf');

  // Create PDF file in Work / 14765567 Canada Inc. / Invoices
  var workFolder = DriveApp.getFoldersByName("Work").next();
  var myCompanyFolder = workFolder.getFoldersByName("14765567 Canada Inc.").next();
  var invoicesFolder = myCompanyFolder.getFoldersByName("Invoices").next();
  return invoicesFolder.createFile(blob);
}

function showLink(pdfFile) {
  var url = pdfFile.getUrl();
  var htmlString = '<base target="_blank">' + '<h3>Invoice generated successfully!</h3><a href="' + url + '">Download Invoice</a>';
  var html = HtmlService.createHtmlOutput(htmlString);
  SpreadsheetApp.getUi().showModalDialog(html, 'Invoice Download');
}

function updateNextInvoiceNumber(trackingSheet) {
  var nextInvoiceNumberCell = trackingSheet.getRange("B1");
  var nextInvoiceNumber = nextInvoiceNumberCell.getValue();
  nextInvoiceNumberCell.setValue(nextInvoiceNumber + 1);
}

function updateLedger(activeSheet, trackingSheet, invoiceDate, grossAmount) {
  nextRecordNumberCell = trackingSheet.getRange("B2");
  var nextRecordNumber = nextRecordNumberCell.getValue();
  var nextRecordRange = activeSheet.getSheetByName("Ledger").getRange("A"+nextRecordNumber+":B"+nextRecordNumber);
  nextRecordRange.setValues([[invoiceDate, grossAmount]]);
  nextRecordNumberCell.setValue(nextRecordNumber + 1);
}
