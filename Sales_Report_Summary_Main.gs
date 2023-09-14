/**
 *   FUNCTION: Sales_Report_Summary_Main()
 *    PURPOSE: Create and format report of open sales opportunities.
 *
 *    TRIGGER: 
 *      INPUT: 'CompanyPlaceholder_Report_Sales_Pipeline'.'Open Opportunities - Sales'
 *             --> ProsperWorks CRM
 *
 *  CONTAINER: CompanyPlaceholder_Report_Sales_Pipeline
 *    PROJECT: Sales_Report
 *
 *     AUTHOR: trevor.young@CompanyPlaceholder.com
 *       DATE: 2017-04-20
 *
 *        TOC:
 *          1) DECLARATION and INITIALIZATION
 *          2) BUILD - REPORT TITLE and COLUMN HEADERS
 *          3) BUILD - SALES REPRESENTATIVE OPPORTUNITIES
 *          4) BUILD - SUMMARY TOTALS
 *          5) FORMATTING
 *          6) WRITE - VALUES TO SHEET
 */
function Sales_Report_Summary_Main() {


  // ----------------------------------------------------------------------------
  // DECLARATION and INITIALIZATION
  // ----------------------------------------------------------------------------

  const NAME_DATA_SOURCE = 'Open Opportunities - Sales';
  const NAME_DATA_TARGET = 'Report - Summary';
  const NUMBER_OF_COLUMNS = 4;
  const RECORD_START = 4;
  
  // Misc
  var currentDate = new Date();
  var debug = new Boolean(true);
  var columnHeaders = [];
  var index = 0;
  var startRow = 1;
  var endRow = 1;
  var representativeEndRows = [];
  var representativeStartEnd = Object.create(null);
  var totalOpportunities = 0;
  
  var spreadSheet = SpreadsheetApp.getActive();
  var sheetDataSource = spreadSheet.getSheetByName(NAME_DATA_SOURCE);
  var sheetReportSummary = spreadSheet.getSheetByName(NAME_DATA_TARGET);
  
  // INIT - DATA SOURCE
  var arrayDataSourceValues = sheetDataSource.getDataRange().getValues().filter(
    function(item) {
      var rowID = item[0];
      if( rowID !== undefined && typeof(rowID) === 'number' && !isNaN(rowID) )
        return true;
      return false;
    } );

  arrayDataSourceValues.sort( 
    function(a, b) {
      // 'Account Owner'
      if ( a[2] < b[2] ) return -1;
      if ( a[2] > b[2] ) return 1;

      // 'Stage'
      // if ( a[5] < b[5] ) return 1;
      // if ( a[5] > b[5] ) return -1;

      // 'Close Date'
      if ( a[10] < b[10] ) return -1;
      if ( a[10] > b[10] ) return 1;
      
      // 'Opportunity'
      if ( a[1] < b[1] ) return -1;
      if ( a[1] > b[1] ) return 1;

      return 0;
    } );

  // INIT - DATA TARGET
  var arrayDataTargetValues = [];

  // Delete Sheet - 'Report - Summary'
  if (sheetReportSummary != null) {
    spreadSheet.deleteSheet(sheetReportSummary);
  }
  
  // Create Sheet - 'Report - Summary'
  spreadSheet.insertSheet(NAME_DATA_TARGET, 2);
  sheetTarget = spreadSheet.getSheetByName(NAME_DATA_TARGET);
  sheetTarget.setTabColor("f26f21");
  
  // ----------------------------------------------------------------------------
  // BUILD - REPORT TITLE and COLUMN HEADERS
  // ----------------------------------------------------------------------------
  arrayDataTargetValues.push(['CompanyPlaceholder Sales Pipeline - Summary','','','']);
  arrayDataTargetValues.push(['Updated ' + currentDate.toUTCString(),'','','']);
  arrayDataTargetValues.push(
    [''
    ,''
    ,'Close Date'
    ,'Total Revenue']);

  // ----------------------------------------------------------------------------
  // BUILD - SALES REPRESENTATIVE OPPORTUNITIES
  // ----------------------------------------------------------------------------

  // get list of representatives
  var salesRepresentativeArray = new Array(); // distinct list

  for( index = 0; index < arrayDataSourceValues.length; index++ ) { 
    if( arrayDataSourceValues[index][2] != '' ) {
      if( salesRepresentativeArray.indexOf(arrayDataSourceValues[index][2]) === -1 ) {
        salesRepresentativeArray.push(arrayDataSourceValues[index][2]);
      }
    }
  }
  salesRepresentativeArray.sort();

  if (debug) { Logger.log(salesRepresentativeArray); }

  // foreach representative
  for( index = 0; index < salesRepresentativeArray.length; index++ ) {

    var salesRepresentative = salesRepresentativeArray[index];

    // get rows matching salesRepresentative
    var rows = arrayDataSourceValues.filter( 
      function(item) {
        if(item[2] === salesRepresentative)
          return true;
        return false;
      } );

    // append first representative row
    arrayDataTargetValues.push(
      [salesRepresentative + ' (' + rows.length + ')'
      ,rows[0][1] // 'Opportunity'
      ,rows[0][10] // 'Close Date'
      ,rows[0][4]  === '' ? 0 : rows[0][4]   // 'Total Revenue'
      ]);

    startRow = arrayDataTargetValues.length;  // row tracking

    // append remaining representative rows    
    for( var rowIndex = 1; rowIndex < rows.length; rowIndex++ ) {
      arrayDataTargetValues.push(
        [''                 //  blank beneath first row
        ,rows[rowIndex][1]  // 'Opportunity'
        ,rows[rowIndex][10] // 'Close Date'
        ,rows[rowIndex][4]  === '' ? 0 : rows[rowIndex][4]   // 'Total Revenue'
        ]);
    }

    var endRow = arrayDataTargetValues.length;  // row tracking

    // append representative total
    arrayDataTargetValues.push(
      [''
      ,''
      ,''
      ,'=SUM(D'+startRow+':D'+endRow+')'
      ]);
      
    // track for formatting
    representativeStartEnd[salesRepresentative] = [startRow, endRow];

    // track for easy Summary Totals row
    representativeEndRows.push(arrayDataTargetValues.length);
    totalOpportunities = totalOpportunities + rows.length;
    
  }  // END - foreach representative

  // ----------------------------------------------------------------------------
  // BUILD - SUMMARY TOTALS
  // ----------------------------------------------------------------------------

  // append - Total (<totalOpportunities>)
  arrayDataTargetValues.push(
  ['Total (' + totalOpportunities + ')'
  ,''
  ,''
  ,'=SUM(D' + representativeEndRows.join(',D') + ')'
  ]);

  if(debug) {
    for ( var item in arrayDataTargetValues ) { 
      Logger.log(arrayDataTargetValues[item]);
      }
  }
  
  // ----------------------------------------------------------------------------
  // FORMATTING
  // - Apply before range.setValue for minimal visibility on screen.
  // ----------------------------------------------------------------------------
  
  // -- CONSTANTS --
  
  // global
  const FONT = 'Tahoma';
  const FONT_SIZE = 8;
  const ALIGNMENT_VERTICAL = 'middle';
  const ROW_HEIGHT = 21;
  
  // title
  const HEADER_1_ALIGN_HORIZONTAL = 'center';
  const HEADER_1_BACKGROUND = '#e3efff';
  const HEADER_1_FONT_SIZE = 14;  
  const HEADER_1_FONT_WEIGHT = 'bold';
  const HEADER_1_ROW_HEIGHT = 40;
  
  // time stamp and column headers
  const HEADER_2_FONT_WEIGHT = 'bold';
  const HEADER_2_ROW_HEIGHT = 21;
  
  // -- APPLY FORMAT --
  
  // global - Apply first and override as needed.
  sheetTarget.getRange("A1:Z1000")
    .setFontFamily(FONT)
    .setFontSize(FONT_SIZE)
    .setVerticalAlignment(ALIGNMENT_VERTICAL)
    .setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID ); // Mask GridLines until API supports toggle.

  sheetTarget.setColumnWidth(1, 135);  // fit for Rep Name
  sheetTarget.setColumnWidth(2, 375);  // fit for Opportunity
  sheetTarget.setColumnWidth(3, 100); // 'Close Date'
  sheetTarget.setColumnWidth(4, 100); // 'Total Revenue'

  sheetTarget.deleteColumns(NUMBER_OF_COLUMNS+1, 26 - NUMBER_OF_COLUMNS);
  sheetTarget.deleteRows(arrayDataTargetValues.length + 1, 1000 - arrayDataTargetValues.length);
  
  // report border - must occur before internal vertical border formatting below
  sheetTarget.getRange(1, 1, arrayDataTargetValues.length, NUMBER_OF_COLUMNS)
    .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID);

  // global row height - Explicitly set for XLSX export.
  for (var row = 1; row < arrayDataTargetValues.length + 1; row++) {
    sheetTarget.setRowHeight(row, ROW_HEIGHT);
  }

  // title
  sheetTarget.setFrozenRows(1);
  sheetTarget.setRowHeight(1, HEADER_1_ROW_HEIGHT);
  sheetTarget.getRange(1, 1, 1, NUMBER_OF_COLUMNS)
    .mergeAcross()
    .setHorizontalAlignment(HEADER_1_ALIGN_HORIZONTAL)
    .setBackground(HEADER_1_BACKGROUND)
    .setFontSize(HEADER_1_FONT_SIZE)
    .setFontWeight(HEADER_1_FONT_WEIGHT)
    .setVerticalAlignment('middle');
  
  // time stamp
  sheetTarget.getRange(2, 1, 1, NUMBER_OF_COLUMNS)
    .mergeAcross()
    .setBackground('#4f81bd')
    .setFontColor('white')
    .setFontWeight(HEADER_1_FONT_WEIGHT)
    .setHorizontalAlignment('center');
    
  // column headers
  sheetTarget.getRange(3, 3, 1, 2)
    .setHorizontalAlignment('right')
    .setFontWeight(HEADER_2_FONT_WEIGHT);
    
  // representative rows
  for(var key in representativeStartEnd) {

    var rStartRow = representativeStartEnd[key][0];
    var rEndRow = representativeStartEnd[key][1];
    var rNumberOfRows = rEndRow - rStartRow + 1;

    Logger.log('Start: ' + rStartRow + ', End: ' + rEndRow);
    
    sheetTarget.getRange(rStartRow, 1, rNumberOfRows, NUMBER_OF_COLUMNS)
      .setBorder(true, true, true, true, true, null, 'black', SpreadsheetApp.BorderStyle.SOLID)
      .getCell(1, 1).setFontWeight('bold'); // representative name

    sheetTarget.getRange(rEndRow+1, NUMBER_OF_COLUMNS, 1).setFontWeight('bold');  // representative total
  }
  
  // Total/last row
  sheetTarget.getRange(arrayDataTargetValues.length, 1, 1, NUMBER_OF_COLUMNS)
    .setBorder(true, null, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.DASHED)
    .setFontWeight('bold');
  
  // column specific
  var numberOfRows = arrayDataTargetValues.length - RECORD_START + 1;
  
  sheetTarget.getRange(RECORD_START, 3, numberOfRows)
    .setHorizontalAlignment('right')
    //.setNumberFormat("m/dd/yyyy");
  
  sheetTarget.getRange(RECORD_START, 4, numberOfRows)
    .setHorizontalAlignment('right')
    .setNumberFormat('[$$-540A]#,##0');

  // ----------------------------------------------------------------------------
  // WRITE - VALUES TO TARGET SHEET
  // ----------------------------------------------------------------------------
  
  // write all arrayDataTargetValues
  sheetTarget
    .getRange(1, 1, arrayDataTargetValues.length, NUMBER_OF_COLUMNS)
    .setValues(arrayDataTargetValues);

  sheetTarget.setActiveSelection("A1:A1");

}