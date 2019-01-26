// These columns are mandatory.
var 
  COL_NAME = 0,
  COL_URL = 1,
  COL_USERNAME = 2,
  COL_PASSWORD = 3;

// onOpen is executed when the sheet is opened.
// adds the Password Manager menu to the sheet.
function onOpen() {
 SpreadsheetApp.getUi()
   .createMenu('Password Manager')
   .addItem('Decrypt password', 'openDecryptUI')
   .addItem('Add new password', 'openAddNewEntryUI')
   .addToUi();
}

// openDecryptUI is executed when Decrypt password menu item is clicked.
// opens up a dialog box, rendering decryptPassword.html where user can 
// add a new password, encrypt it with a shared secret and save it.
function openDecryptUI() {
  var html = HtmlService.createTemplateFromFile('decryptPassword').evaluate()
      .setWidth(400)
      .setHeight(150);
  SpreadsheetApp.getUi()
      .showModalDialog(html, 'Decrypt password');
}

// openAddNewEntryUI is executed when Add new password menu item is clicked.
// opens up a dialod where user can enter the shared secret and decrypt the
// password in the row that is currently focused.
function openAddNewEntryUI() {
  var html = HtmlService.createTemplateFromFile('newEntry').evaluate()
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi()
      .showModalDialog(html, 'Add a new entry');
}

// createNewEntry gets a form object form the frontend dialog and saves it
// into a new row in the sheet. The encrypted password object (json) is
// base64 encoded.
function createNewEntry(form) {
  blob = Utilities.base64Encode(form.password); 
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow([form.name, form.url, form.username, blob]);
}

// getPassword returns the base64 decoded encrypted password object (json)
// from the current record.
function getPassword() {
  var data = getRecord()
  blob = bin2String(Utilities.base64Decode(data[COL_PASSWORD]));
  return blob;
}

// getRecord gets the row that is currently on focus.
function getRecord() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rowNum = sheet.getActiveCell().getRow();
  if (rowNum > data.length) return [];
  var record = [];
  for (var col=0;col<headers.length;col++) {
    record.push(data[rowNum-1][col]);
  }
  return record;
}

// bin2String converts a binary array to a string.
function bin2String(array) {
  var result = "";
  for (var i = 0; i < array.length; i++) {
    result += String.fromCharCode(parseInt(array[i], 10));
  }
  return result;
}

// include is a template helper function which renders HTML from a file.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}