function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Surprised Pikachu')
    .addItem('Show Sidebar', 'showSidebar')
    .addItem('Create Expanding Row Template', 'expandRowHelper')
    .addItem('Create Expanding Column Template', 'expandColumnHelper')
    .addToUi();
  //PropertiesService.getScriptProperties().deleteAllProperties();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Compare Itemized Pricing List')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function addShortcut() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  // check if shortcut cells have been added
  const backgrounds = sheet.getRange('A1:B11').getBackgrounds();
  if (sheet.getRange('A1').getBackground() == '#fff2cc') {
    ui.alert('Shortcut Already Exists', 'Seems like shortcuts already exist. Please click on the shortcut removal to remove and then try again.', ui.ButtonSet.OK)
    return;
  }
  sheet.insertRowsBefore(1, 11);
  sheet.getRange('1:11').merge().setBackground('#fff2cc')
  sheet.setColumnWidth(1, 400);
  sheet.setColumnWidth(2, 99);
  const expandingColumnImg = 'https://i.imgur.com/4wbSA9K.png';
  const expandingRowImg = 'https://i.imgur.com//Tz5jXsS.png';
  sheet.insertImage(expandingRowImg, 1, 1)
    .setAnchorCellXOffset(10)
    .setAnchorCellYOffset(15)
    .setHeight(180).setWidth(204)
    .assignScript('expandRowHelper');
  sheet.insertImage(expandingColumnImg, 1, 1)
    .setAnchorCellXOffset(230)
    .setAnchorCellYOffset(15)
    .setHeight(180).setWidth(240)
    .assignScript('expandColumnHelper');
}

function removeShortcut() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  button = sheet.getImages();
  button.map((b) => b.remove());
  sheet.deleteRows(1, 11)
}

function expandRowHelper() {confirmRunScript("row");}
function expandColumnHelper() {confirmRunScript("column");}

function confirmRunScript(type) {
  var ui = SpreadsheetApp.getUi();
  const sheetName = SpreadsheetApp.getActiveSpreadsheet().getSheetName();
  const valid = checkSheetValid(sheetName)
  if (!valid) return;

  var response = ui.alert('Confirm Action', `Are you sure you would like to run this script on the sheet named ${sheetName}?`, ui.ButtonSet.OK_CANCEL);
  // process input
  if (response == ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('dataSource', sheetName);
    if (type == "column") createExpandColumnTemplate();
    else if (type == "row") createExpandRowTemplate();
    else if (type == "add") addShortcut();
    else if (type == "remove") removeShortcut();
  }
  else ui.alert("Mission Aborted", "Running script cancelled.", ui.ButtonSet.OK)
}

function checkSheetValid(sheetName) {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var msg =  `Data format is invalid for sheet named ${sheetName}. Please run the script when you're in the correct sheet or format your sheet to have one column of items and one column of prices.`

  if (sheet.getRange('A1').getBackground() == '#fff2cc') {
    if (sheet.getRange('B13').getNumberFormat() != '"$"#,##0.00' || sheet.getRange('A12').getBackground() == "#ffffff") {
      ui.alert('Invalid Data Format', msg, ui.ButtonSet.OK);
      return false
    }
  }
  else {
    if (sheet.getRange('B2').getNumberFormat() != '"$"#,##0.00' || sheet.getRange('A1').getBackground() == "#ffffff") {
      ui.alert('Invalid Data Format', msg, ui.ButtonSet.OK);
      return false;
    }
  }
  return true;
}