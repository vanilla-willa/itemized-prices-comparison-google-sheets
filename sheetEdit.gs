function onEdit(e) {
    // get sheet
    const sheet = e.source.getActiveSheet();
    // const sheetName = sheet.getSheetName();
    const sheetId = sheet.getSheetId();
    const currentRange = sheet.getActiveRange();
  
    if (sheetId == parseInt(PropertiesService.getScriptProperties().getProperty('dataSourceId'))) return;
    if (currentRange.isChecked() == null) return;
    
    const currentCell = currentRange.getA1Notation();
    const currentCol = currentCell.match(/[a-zA-Z]+/)[0];
    const currentRow = Number(currentCell.match(/\d+/)[0]);
    const numRows = sheet.getLastRow();
    const numCols = sheet.getLastColumn();
  
    // if A2 has background, then it's expanding row 
    if (sheet.getRange('A2').getBackground() != '#ffffff') {
      // Logger.log("Inside expanding row spreadsheet")
      // check if any rows or columns got added
      
      if (numRows != PropertiesService.getScriptProperties().getProperty('rNumRows')) {
        Logger.log("updating rows");
        PropertiesService.getScriptProperties().setProperty('rNumRows', numRows);
        // merge cells
        
        return;
      }
      if (numCols != PropertiesService.getScriptProperties().getProperty('rNumCols')) {
        Logger.log("updating columns")
        PropertiesService.getScriptProperties().setProperty('rNumCols', numCols);
        return;
      }
      ranges = JSON.parse(PropertiesService.getScriptProperties().getProperty('rCheckboxRange'));
      var start, end, columnCheckedData;
      for (var i = 0; i < ranges.length; i++) {
        [start, end] = ranges[i].split(':');
        if (currentCol <= col2letter(numCols) && currentCol >= 'B' && currentRow >= start && currentRow <= end) {
          columnCheckedData = sheet.getRange(currentCol + ':' + currentCol).getValues();
          getCheckedColumn(columnCheckedData, currentCol);
          break;
        }
      }
    }
    else if (sheet.getRange('B1').getBackground() != '#ffffff') {
      // Logger.log("Inside expanding column spreadsheet")
      // check if any rows or columns got added
      if (numRows != PropertiesService.getScriptProperties().getProperty('cNumRows')) {
        Logger.log("updating rows");
        PropertiesService.getScriptProperties().setProperty('cNumRows', numRows);
        return;
      }
      if (numCols != PropertiesService.getScriptProperties().getProperty('cNumCols')) {
        Logger.log("updating columns")
        PropertiesService.getScriptProperties().setProperty('cNumCols', numCols);
        return;
      }
      ranges = JSON.parse(PropertiesService.getScriptProperties().getProperty('cCheckboxRange'));
      var start, end, rowCheckedData;
      for (var i = 0; i < ranges.length; i++) {
        [start, end] = ranges[i].split(':');
        if (letter2col(currentCol) <= letter2col(end) && letter2col(currentCol) >= letter2col(start) && currentRow >= 2 && currentRow <= numRows) {
          rowCheckedData = sheet.getRange(currentRow + ':' + currentRow).getValues()[0];
          getCheckedRow(rowCheckedData, currentRow);
          break;
        }
      }
    }
  }
  
  function getCheckedColumn(column, currentCol) {
    const CB = SpreadsheetApp.DataValidationCriteria.CHECKBOX;
    const pricingData = JSON.parse(PropertiesService.getScriptProperties().getProperty('pricingData'));
    const rowTitles = JSON.parse(PropertiesService.getScriptProperties().getProperty('rowDict'));
    var sum = 0;
    for (var i = 0; i < column.length; i++) {
      if (column[i][0] != null) {
        if (column[i][0] == true) {
          rowTitle = rowTitles[i+1];
          sum += pricingData[rowTitle];
        }
      }
    }
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    sheet.getRange(currentCol +  lastRow)
      .setValue(sum)
      .setNumberFormat('"$"#,##0.00')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }
  
  function getCheckedRow(row, currentRow) {
    const CB = SpreadsheetApp.DataValidationCriteria.CHECKBOX;
    const pricingData = JSON.parse(PropertiesService.getScriptProperties().getProperty('pricingData'));
    const columnTitles = JSON.parse(PropertiesService.getScriptProperties().getProperty('columnDict'));
    var sum = 0;
    for (var i = 0; i < row.length; i++) {
      if (row[i] != null) {
        if (row[i] == true) {
          columnTitle = columnTitles[col2letter(i+1)];
          sum += pricingData[columnTitle];
        }
      }
    }
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastCol = sheet.getLastColumn();
    sheet.getRange(col2letter(lastCol) + currentRow)
      .setValue(sum)
      .setNumberFormat('"$"#,##0.00')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }