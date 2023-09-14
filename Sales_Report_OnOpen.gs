/** onOpen */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Reports')
    .addItem('Email...', 'menuItemRD1')
      .addItem('Refresh', 'menuItemRD2')
      .addSeparator()
      .addToUi();
}

/** menuItemRD1 - Email Report */
function menuItemRD1() {

  var debug = new Boolean(true);
  var recipientList = [];
  var spreadSheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  
  var columnValues = spreadSheet.getSheetByName('Configuration').getRange('A2:A').getValues();
  
  for ( var index in columnValues ) {
    if (columnValues[index] != '') {
      recipientList.push(columnValues[index]);
    }
  }
  
  recipientList.sort();
  
  // email confirmation
  var alertResult = ui.alert('Send Report to Recipient List:',
                               recipientList.join('\n'),
                               ui.ButtonSet.OK_CANCEL);
  
  if (alertResult == ui.Button.OK) {
    Sales_Report_Email(recipientList.join(', '));
  }  
}

/** menuItemRD2 - Refresh Report */
function menuItemRD2() {

  var spreadSheet = SpreadsheetApp.getActive();
  
  var sheetReportSummary = spreadSheet.getSheetByName('Report - Summary');
  if (sheetReportSummary != null) {
    spreadSheet.deleteSheet(sheetReportSummary);
  }
  
  var sheetReportDetail = spreadSheet.getSheetByName('Report - Detail');
  if (sheetReportDetail != null) {
    spreadSheet.deleteSheet(sheetReportDetail);
  }
  
  Sales_Report_Detail_Main();
  Sales_Report_Summary_Main();

  spreadSheet.getSheetByName('Report - Detail').activate();
  
}