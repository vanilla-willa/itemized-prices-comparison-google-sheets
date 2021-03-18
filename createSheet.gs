/************************************
 *                                  *
 *        CREATE NEW SHEET          *
 *                                  *
 ************************************/

function createExpandColumnTemplate() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    const sheet = spreadsheet.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange('A:A');
    // only get data with values
    const data = range.getValues().filter(String);
    const backgrounds = range.getBackgrounds();
    
    // get category names in the spreadsheet
    var categoryRows = [];
    for (var i = 0; i < lastRow; i++) {
      // i is iterable starting at 0 since it's going through backgrounds array
      // but need to +1 since rows start at 1
      if (backgrounds[i] != "#ffffff" && backgrounds[i] != '#fff2cc') categoryRows.push(i+1)
    }
  
    // create new sheet with unique name
    var name = Utilities.formatDate(new Date(), "America/Los_Angeles", "yyyy-MM-dd HH:mm:ss");
    var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
  
    // populate spreadsheet with data as header
    var numCols = newSheet.getRange('1:1').getLastColumn();
    var numRows = newSheet.getRange('A:A').getLastRow();
    newSheet.deleteRows(1, numRows - 10)
    numRows = newSheet.getRange('A:A').getLastRow();
    if (numCols < data.length) newSheet.insertColumns(1, data.length - numCols)
    else newSheet.deleteColumns(1, numCols - data.length)
  
    // insert Name column at the beginning
    newSheet.insertColumnBefore(1);
    newSheet.getRange('A1').setValue('Name')
    numCols = newSheet.getRange('1:1').getLastColumn();
  
    // add 1 to the categoryRows array to account for Name column
    // subtract 11 to account for the additional image rows
    if (sheet.getRange('A1').getBackground() == '#fff2cc') {
      categoryRows = categoryRows.map(categoryRows => categoryRows -= 10)
    }
    else categoryRows = categoryRows.map(categoryRows => categoryRows += 1)
  
    // A column of values is represented in Apps Script as [['a'], ['b'], ['c']]. 
    // A row is represented as [['a', 'b', 'c']].
    newSheet.getRange('B1:1')
      .setValues(col2row(data))
    newSheet.getRange('1:1')
      .setFontWeight('bold')
      .setBackground('#F7EBEC') // light pink
      .setHorizontalAlignment('center');
  
    // add a Total column at end
    newSheet.insertColumnAfter(numCols);
    numCols = newSheet.getRange('1:1').getLastColumn();
    newSheet.getRange( col2letter(numCols) + '1' )
      .setValue('Total')
      .setBackground('#C3E6EA') // light sky blue
  
    // create dictionary mapping column ABC... to header
    //ie. {'A': 'Gas', 'B': 'Rent'...}
    const colTitles = newSheet.getRange('1:1').getValues()[0];
    var col2colTitles = {}
    for (var i = 0; i < colTitles.length; i++) {
      col2colTitles[col2letter(i+1)] = colTitles[i]
    }
    PropertiesService.getScriptProperties().setProperty('columnDict', JSON.stringify(col2colTitles));
    
    // clear background for Name column
    newSheet.getRange('A:A').setBackground(null);
  
    // format header and name column
    newSheet.getRange('1:1')
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    newSheet.getRange('A:A')
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
    // separate header names for categories
    var letter = '';
    for (var i = 0; i < categoryRows.length; i++) {
      letter = col2letter(categoryRows[i])
      newSheet.getRange(letter + ':' + letter)
        .setBackground('#E6E0EC') // light purple
        .mergeVertically();
    }
  
    // freeze column
    newSheet.setFrozenColumns(1);
    // add borders
    newSheet.getRange('1:1')
    .setBorder(null, null, true, null, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
    var checkboxRange = [];
    var checkbox = '';
    // group columns to allow expanding / collapsing
    for (var i = 0; i < categoryRows.length; i++) {
      var start = col2letter(categoryRows[i] + 1);
      if (i < categoryRows.length - 1) var end = col2letter(categoryRows[i+1] - 1)
      // exclude the Total column
      else var end = col2letter(numCols - 1);
      newSheet.getRange( start + ':' + end).shiftColumnGroupDepth(1);
      
      // add checkboxes
      newSheet.getRange(start + '2:' + end + String(numRows))
        .setFontColor('#000000')
        .setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(true)
        .requireCheckbox()
        .build());
      checkbox = start + ':' + end;
      checkboxRange.push(checkbox)
      // collapse grouped columns
      newSheet.getColumnGroup(categoryRows[i], 1).collapse();
    }
  
    PropertiesService.getScriptProperties().setProperty('cCheckboxRange', JSON.stringify(checkboxRange));
    PropertiesService.getScriptProperties().setProperty('cNumRows', numRows);
    PropertiesService.getScriptProperties().setProperty('cNumCols', numCols);
    pricingData();
  }
  
  function createExpandRowTemplate() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    const sheet = spreadsheet.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange('A:A');
    // only get data with values
    const data = range.getValues().filter(String);
    const backgrounds = range.getBackgrounds();
    
    // get category names in the spreadsheet
    var categoryRows = [];
    for (var i = 0; i < lastRow; i++) {
      // i is iterable starting at 0 since it's going through backgrounds array
      // but need to +1 since rows start at 1
      if (backgrounds[i] != "#ffffff" && backgrounds[i] != '#fff2cc') categoryRows.push(i+1)
    }
  
    // create new sheet with unique name
    var name = Utilities.formatDate(new Date(), "America/Los_Angeles", "yyyy-MM-dd HH:mm:ss");
    var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
  
    // populate spreadsheet with data as header
    var numCols = newSheet.getRange('1:1').getLastColumn();
    var numRows = newSheet.getRange('A:A').getLastRow();
    newSheet.deleteColumns(1, numCols - 10)
    numCols = newSheet.getRange('1:1').getLastColumn();
    if (numRows < data.length) newSheet.insertRows(1, data.length - numRows)
    else newSheet.deleteRows(1, numRows - data.length)
  
    // insert Name row at the beginning
    newSheet.insertRowBefore(1);
    newSheet.getRange('A1').setValue('Name')
    numRows = newSheet.getRange('A:A').getLastRow();
  
    // add 1 to the categoryRows array to account for Name column
    // subtract 11 to account for the additional image rows
    if (sheet.getRange('A1').getBackground() == '#fff2cc') {
      categoryRows = categoryRows.map(categoryRows => categoryRows -= 10)
    }
    else categoryRows = categoryRows.map(categoryRows => categoryRows += 10)
  
    // A column of values is represented in Apps Script as [['a'], ['b'], ['c']]. 
    // A row is represented as [['a', 'b', 'c']].
    newSheet.getRange('A2:A')
      .setValues(data)
    newSheet.getRange('A:A')
      .setFontWeight('bold')
      .setBackground('#F7EBEC') // light pink
      .setHorizontalAlignment('center');
  
    // add a Total row at end
    newSheet.insertRowAfter(numRows);
    numRows = newSheet.getRange('A:A').getLastRow();
    newSheet.getRange('A' + numRows)
      .setValue('Total')
      .setBackground('#C3E6EA') // light sky blue
  
    // create dictionary mapping row 123 to header
    // ie. {'1': 'Gas', '2': 'Rent'...}
    const rowTitles = newSheet.getRange('A:A').getValues();
    var row2rowTitles = {}
    for (var i = 0; i < rowTitles.length; i++) {
      row2rowTitles[i+1] = rowTitles[i][0]
    }
    PropertiesService.getScriptProperties().setProperty('rowDict', JSON.stringify(row2rowTitles));
    
    // clear background for Name column
    newSheet.getRange('1:1').setBackground(null);
  
    // format header and Name row
    newSheet.getRange('A:A')
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    newSheet.getRange('1:1')
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
    // separate header names for categories
    var letter = '';
    for (var i = 0; i < categoryRows.length; i++) {
      rowNum = categoryRows[i];
      newSheet.getRange(rowNum + ':' + rowNum)
        .setBackground('#E6E0EC') // light purple
        .mergeAcross();
    }
  
    //freeze row
    newSheet.setFrozenRows(1);
    // add borders
    newSheet.getRange('A:A')
      .setBorder(null, null, true, null, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
    var checkboxRange = [];
    var checkbox = '';
    // group columns to allow expanding / collapsing
    for (var i = 0; i < categoryRows.length; i++) {
      var start = categoryRows[i] + 1;
      if (i < categoryRows.length - 1) var end = categoryRows[i+1] - 1
      // exclude the Total row
      else var end = numRows - 1;
      newSheet.getRange( start + ':' + end )
        .activate().shiftRowGroupDepth(1);
      
      // add checkboxes
      newSheet.getRange('B' + start + ':' + col2letter(String(numCols)) + end)
        .setFontColor('#000000')
        .setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(true)
        .requireCheckbox()
        .build());
      checkbox = start + ':' + end;
      checkboxRange.push(checkbox)
      // collapse grouped rows
      newSheet.getRowGroup(categoryRows[i], 1).collapse();
    }
  
    PropertiesService.getScriptProperties().setProperty('rCheckboxRange', JSON.stringify(checkboxRange));
    PropertiesService.getScriptProperties().setProperty('rNumRows', numRows);
    PropertiesService.getScriptProperties().setProperty('rNumCols', numCols);
    pricingData();
  }