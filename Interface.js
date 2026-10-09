// ================================
// INTERFACE & UI TRIGGERS
// ================================

// Creating a menu in a worksheet upon opening
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('NGTL Email Sender')
      .addItem('Open the Dispatch Panel', 'showShippingPopup')
      .addToUi();
}

// Function to open a popup window
function showShippingPopup() {
  var html = HtmlService.createHtmlOutputFromFile('Popup')
      .setWidth(800)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dispatch Panel');
}


// Displays a native, clean message within a Google Sheet without technical URLs
function showNativeAlert(message) {
  SpreadsheetApp.getUi().alert('Sending status', message, SpreadsheetApp.getUi().ButtonSet.OK);
}