/**
 *   FUNCTION: Sales_Report_Detail_Main()
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
 *       DATE: 2017-04-03
 *
 *        TOC:
 *          1) DECLARATION and INITIALIZATION
 *          2) BUILD - REPORT TITLE and COLUMN HEADERS
 *          3) BUILD - SALES REPRESENTATIVE OPPORTUNITIES
 *          4) BUILD - SUMMARY TOTALS
 *          5) FORMATTING
 *          6) WRITE - VALUES TO SHEET
 */
function Sales_Report_Detail_Main() {

  // ----------------------------------------------------------------------------
  // DECLARATION and INITIALIZATION
  // ----------------------------------------------------------------------------

  const NAME_DATA_SOURCE = 'Open Opportunities - Sales';
  const NAME_DATA_TARGET = 'Report - Detail';
  const NUMBER_OF_COLUMNS = 19;

  // Misc
  var currentDate = new Date();
  var debug = new Boolean(true);
  var columnHeaders = [];
  var index = 0;
  var startRow = 1;
  var endRow = 1;
  var representativeSummaryRows = [];
  var representativeHeaderRows = [];
  var totalOpportunities = 0;
  
  var spreadSheet = SpreadsheetApp.getActive();
  var sheetDataSource = spreadSheet.getSheetByName(NAME_DATA_SOURCE);
  var sheetReportDetail = spreadSheet.getSheetByName(NAME_DATA_TARGET);
  
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
      if ( a[5] < b[5] ) return 1;
      if ( a[5] > b[5] ) return -1;

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

  // Delete Sheet - 'Report - Detail'
  if (sheetReportDetail != null) {
    spreadSheet.deleteSheet(sheetReportDetail);
  }
  
  // Create Sheet - 'Report - Detail'
  spreadSheet.insertSheet(NAME_DATA_TARGET, 1);
  sheetTarget = spreadSheet.getSheetByName(NAME_DATA_TARGET);
  sheetTarget.setTabColor("f26f21");
  
  // ----------------------------------------------------------------------------
  // BUILD - REPORT TITLE and COLUMN HEADERS
  // ----------------------------------------------------------------------------
  arrayDataTargetValues.push([' CompanyPlaceholder Sales Pipeline','','','','','','','','','','','','','','','','','','']);
  arrayDataTargetValues.push([' Updated ' + currentDate.toUTCString(),'','','','','','','','','','','','','','','','','','']);  
  arrayDataTargetValues.push(
    ['Account Owner'
    ,'Opportunity'
    ,'Modules'
    ,'Stage'
    ,'Close Date'
    ,'Total Potential Revenue'
    ,'OnSite License Revenue'
    ,'Annual M&S Revenue'
    ,'Services Revenue'
    ,'OnCloud Monthly Revenue'
    ,'OnCloud Terms'
    ,'OnCloud Total Revenue'
    ,'Platform'
    ,'ERP System'
    ,'ERP System Other'
    ,'Competition'
    ,'Company Annual Revenue'
    ,'Detail'
    ,'Lead Source']);
    
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

    // append representative header - <Rep Name> (Count: <numberOfOpportunities>)
    arrayDataTargetValues.push([salesRepresentative + ' (Count: ' + rows.length + ')','','','','','','','','','','','','','','','','','','']);
    
    // tracking for header row
    representativeHeaderRows.push(arrayDataTargetValues.length);

    // append opportunities, arrange columns, and standardize data for target
    startRow = arrayDataTargetValues.length+1;
    
    for( var rowIndex = 0; rowIndex < rows.length; rowIndex++ ) {
      arrayDataTargetValues.push(
        [''                 //  blank beneath representative header
        ,rows[rowIndex][1]  // 'Opportunity'
        ,rows[rowIndex][35].replace(/\;/g, ', ')  // 'Product Modules'
        ,rows[rowIndex][5]  // 'Stage'
        ,rows[rowIndex][10] // 'Close Date'
        ,rows[rowIndex][4]  === '' ? 0 : rows[rowIndex][4]   // 'Total Potential Revenue'
        ,rows[rowIndex][25] === '' ? 0 : rows[rowIndex][25]  // 'OnSite License Revenue'
        ,rows[rowIndex][24] === '' ? 0 : rows[rowIndex][24]  // 'Annual M&S Revenue'
        ,rows[rowIndex][26] === '' ? 0 : rows[rowIndex][26]  // 'Services Revenue'
        ,rows[rowIndex][27] === '' ? 0 : rows[rowIndex][27]  // 'OnCloud Monthly Revenue'
        ,rows[rowIndex][28] === '' ? 0 : rows[rowIndex][28]  // 'OnCloud Terms'
        ,rows[rowIndex][29] === '' ? 0 : rows[rowIndex][29]  // 'OnCloud Total Revenue'
        ,rows[rowIndex][31]  // 'Platform'
        ,rows[rowIndex][32]  // 'ERP System'
        ,rows[rowIndex][34]  // 'ERP System Other'
        ,rows[rowIndex][33].replace(/\;/g, ', ')  // 'Competition'
        ,rows[rowIndex][30]  // 'Company Annual Revenue'
        ,rows[rowIndex][9]   // 'Detail'
        ,rows[rowIndex][13]  // 'Lead Source'
        ]);
    }

    // append representative summary
    var endRow = arrayDataTargetValues.length;
    
    arrayDataTargetValues.push(
      ['Summary'
      ,''
      ,''
      ,''
      ,''
      ,'=SUM(F'+startRow+':F'+endRow+')'
      ,'=SUM(G'+startRow+':G'+endRow+')'
      ,'=SUM(H'+startRow+':H'+endRow+')'
      ,'=SUM(I'+startRow+':I'+endRow+')'
      ,'=SUM(J'+startRow+':J'+endRow+')'
      ,''
      ,'=SUM(L'+startRow+':L'+endRow+')'
      ,''
      ,''
      ,''
      ,''
      ,''
      ,''
      ,'']);

    // tracking for Summary Total row
    representativeSummaryRows.push(arrayDataTargetValues.length);
    totalOpportunities = totalOpportunities + rows.length;
    
  }  // END - foreach representative

  // ----------------------------------------------------------------------------
  // BUILD - SUMMARY TOTALS
  // ----------------------------------------------------------------------------

  // append header - All (Count: <totalOpportunities>)
  arrayDataTargetValues.push(['All (Count: ' + totalOpportunities + ')','','','','','','','','','','','','','','','','','','']);
  
  // track for format on 'All Count' Row
  representativeHeaderRows.push(arrayDataTargetValues.length);

  // append summary total
  arrayDataTargetValues.push(
    ['Summary Totals'
    ,''
    ,''
    ,''
    ,''
    ,'=SUM(F' + representativeSummaryRows.join(',F') + ')'
    ,'=SUM(G' + representativeSummaryRows.join(',G') + ')'
    ,'=SUM(H' + representativeSummaryRows.join(',H') + ')'
    ,'=SUM(I' + representativeSummaryRows.join(',I') + ')'
    ,'=SUM(J' + representativeSummaryRows.join(',J') + ')'
    ,''
    ,'=SUM(L' + representativeSummaryRows.join(',L') + ')'
    ,''
    ,''
    ,''
    ,''
    ,''
    ,''
    ,'']);

  // track for format on Summary Totals
  representativeSummaryRows.push(arrayDataTargetValues.length);

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
  const HEADER_1_ALIGN_HORIZONTAL = 'left';
  const HEADER_1_BACKGROUND = '#e3efff';
  const HEADER_1_FONT_SIZE = 14;  
  const HEADER_1_FONT_WEIGHT = 'bold';
  const HEADER_1_ROW_HEIGHT = 40;
  
  // column headers
  const COLUMN_HEADER_ROW_NUMBER = 3;
  const HEADER_2_ALIGN_HORIZONTAL = 'center';
  const HEADER_2_BACKGROUND = '#4f81bd';  
  const HEADER_2_FONT_COLOR = 'white';  
  const HEADER_2_FONT_LINE = 'underline';
  const HEADER_2_FONT_WEIGHT = 'bold';
  const HEADER_2_ROW_HEIGHT = 40;
  
  // representative records
  const RECORD_1_ROW_START = 4;
  
  // -- APPLY FORMAT --
  
  // global - Apply first and override as needed.
  sheetTarget.getRange("A1:Z1000")
    .setFontFamily(FONT)
    .setFontSize(FONT_SIZE)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, "white", SpreadsheetApp.BorderStyle.SOLID ); // Mask GridLines until API supports toggle.

  columnHeaders = columnHeaders.concat([''], arrayDataTargetValues[COLUMN_HEADER_ROW_NUMBER-1]); // shorten reference

  sheetTarget.setColumnWidth(columnHeaders.indexOf('Opportunity'), 275);  // 'Opportunity'
  sheetTarget.setColumnWidth(columnHeaders.indexOf('Stage'), 135);  // 'Stage'
  sheetTarget.setColumnWidth(columnHeaders.indexOf('Detail'), 275); // 'Detail'
  sheetTarget.setColumnWidth(columnHeaders.indexOf('Lead Source'), 150); // 'Lead Source'
  sheetTarget.hideColumns(columnHeaders.indexOf('ERP System Other')); // 'ERP System Other'
  sheetTarget.hideColumns(columnHeaders.indexOf('Detail')); // 'Detail'
  sheetTarget.deleteColumns(NUMBER_OF_COLUMNS+1, 26 - NUMBER_OF_COLUMNS);
  sheetTarget.deleteRows(arrayDataTargetValues.length + 1, 1000 - arrayDataTargetValues.length);
      
  // global row height - Explicitly set for XLSX export.
  for (var row = 1; row < arrayDataTargetValues.length + 1; row++) {
    sheetTarget.setRowHeight(row, ROW_HEIGHT);
  }
  
  // title
  sheetTarget.setRowHeight(1, HEADER_1_ROW_HEIGHT);
  sheetTarget.getRange(1, 1, 1, NUMBER_OF_COLUMNS)
    .mergeAcross()
    .setHorizontalAlignment(HEADER_1_ALIGN_HORIZONTAL)
    .setBackground(HEADER_1_BACKGROUND)
    .setFontSize(HEADER_1_FONT_SIZE)
    .setFontWeight(HEADER_1_FONT_WEIGHT);
  
  // time stamp
  sheetTarget.getRange(2, 1, 1, NUMBER_OF_COLUMNS)
    .mergeAcross()
    .setHorizontalAlignment(HEADER_1_ALIGN_HORIZONTAL)
    .setBackground(HEADER_1_BACKGROUND)
    .setFontWeight(HEADER_1_FONT_WEIGHT);
    
  // column headers
  sheetTarget.setFrozenRows(3)
  sheetTarget.setRowHeight(3, HEADER_2_ROW_HEIGHT);
  sheetTarget.getRange(3, 1, 1, NUMBER_OF_COLUMNS)
    .setHorizontalAlignment(HEADER_2_ALIGN_HORIZONTAL)
    .setBackground(HEADER_2_BACKGROUND)
    .setFontColor(HEADER_2_FONT_COLOR)
    .setFontLine(HEADER_2_FONT_LINE)
    .setFontWeight(HEADER_2_FONT_WEIGHT)
    .setWrap(true);

  // representative headers
  for(index in representativeHeaderRows) {
    sheetTarget.getRange(representativeHeaderRows[index], 1, 1, NUMBER_OF_COLUMNS)
      .mergeAcross()
      .setHorizontalAlignment('left')
      .setBackground('#95b3d7')
      .setFontColor(HEADER_2_FONT_COLOR)
      .setFontWeight(HEADER_2_FONT_WEIGHT);
  }

  // representative summaries
  for(index in representativeSummaryRows) {
    sheetTarget.getRange(representativeSummaryRows[index], 1, 1, NUMBER_OF_COLUMNS)
      .setFontWeight('bold')
      .setBorder(true, null, null, null, null, null, HEADER_2_BACKGROUND, SpreadsheetApp.BorderStyle.SOLID)
  }

  // column specific
  var numberOfRows = arrayDataTargetValues.length - RECORD_1_ROW_START + 1;
  
  sheetTarget.getRange(RECORD_1_ROW_START, columnHeaders.indexOf('Close Date'), numberOfRows)
    .setHorizontalAlignment('right');
  
  sheetTarget.getRange(RECORD_1_ROW_START, columnHeaders.indexOf('Total Potential Revenue'), numberOfRows, 5) //6-10
    .setHorizontalAlignment('right')
    .setNumberFormat('[$$-540A]#,##0');
    
  sheetTarget.getRange(RECORD_1_ROW_START, columnHeaders.indexOf('OnCloud Terms'), numberOfRows)
    .setHorizontalAlignment('right');
    
  sheetTarget.getRange(RECORD_1_ROW_START, columnHeaders.indexOf('OnCloud Total Revenue'), numberOfRows)
    .setHorizontalAlignment('right')
    .setNumberFormat('[$$-540A]#,##0');
    
  sheetTarget.getRange(RECORD_1_ROW_START, columnHeaders.indexOf('Platform'), numberOfRows)
    .setHorizontalAlignment('center');
    
  sheetTarget.getRange(RECORD_1_ROW_START, columnHeaders.indexOf('Company Annual Revenue'), numberOfRows)
    .setHorizontalAlignment('right')
    .setNumberFormat('[$$-540A]#,##0');
  
  // report border
  sheetTarget.getRange(1, 1, arrayDataTargetValues.length, NUMBER_OF_COLUMNS)
    .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID); 
  
  // ----------------------------------------------------------------------------
  // WRITE - VALUES TO SHEET
  // ----------------------------------------------------------------------------
  
  // write all arrayDataTargetValues
  sheetTarget
    .getRange(1, 1, arrayDataTargetValues.length, NUMBER_OF_COLUMNS)
    .setValues(arrayDataTargetValues);

  // add Notes on Column 'Opportunity' using Column Detail
  var noteValues = sheetTarget.getRange(RECORD_1_ROW_START, columnHeaders.indexOf('Detail'), numberOfRows).getValues();
  sheetTarget.getRange(RECORD_1_ROW_START, columnHeaders.indexOf('Opportunity'), numberOfRows).setNotes(noteValues);

  sheetTarget.setActiveSelection("A1:A1");

}