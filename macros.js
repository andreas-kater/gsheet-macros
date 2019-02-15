
function number() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setNumberFormat('#,##0')
};

function gbp() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setNumberFormat('[$£]#,##0')
};

function usd() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setNumberFormat('"$"#,##0');
};

function eur() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setNumberFormat('[$€]#,##0')
};

function parameter() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground(null);
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#0000ff', SpreadsheetApp.BorderStyle.SOLID);
};

function link() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBorder(false, false, false, false, false, false);
  spreadsheet.getActiveRangeList().setBackground('#d9ead3');
};


function clear_formatting() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground(null);
  spreadsheet.getActiveRangeList().setBorder(false, false, false, false, false, false);
};

