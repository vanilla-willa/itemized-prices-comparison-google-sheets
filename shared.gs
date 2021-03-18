/************************************
 *                                  *
 *        SHARED FUNCTIONS          *
 *                                  *
 ************************************/

function col2row(column) {
    //return [column.map(function(row) {return row[0];})];
    return [column.map(row => row[0])];
  } 
  
  function col2letter(column) {
    // A starts at 65 in ASCII but column A is column 1
    var temp, letter = '';
    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  }
  
  function letter2col(letter) {
    var column = 0, length = letter.length;
    for (var i = 0; i < length; i++)
    {
      column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column;
  }
  
  function pricingData() {
    const sheet = getSheetById(parseInt(PropertiesService.getScriptProperties().getProperty('dataSourceId')))
    const lastRow = sheet.getLastRow();
    const range = sheet.getDataRange();
    const data = range.getValues();
    const backgrounds = range.getBackgrounds();
  
    var pricing = {};
    if (sheet.getRange('A1').getBackground() == '#fff2cc') var i = 11;
    else var i = 0;
    for (i; i < lastRow; i++) {
      if (backgrounds[i][0] == "#ffffff") pricing[data[i][0]] = data[i][1];
    }
  
    PropertiesService.getScriptProperties().setProperty('pricingData', JSON.stringify(pricing))
  }
  
  function getSheetById(sheetId) {
    var sheet = SpreadsheetApp.getActive().getSheets().filter(
      function(s) {return s.getSheetId() === sheetId;}
    )[0];
    return sheet;
  }