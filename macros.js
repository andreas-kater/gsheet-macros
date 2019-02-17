
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

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = []
  menuEntries.push({name: "Trace Dependents", functionName: "traceDependents"});
  ss.addMenu("Detective", menuEntries);
}

function trace_dependents(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var regex = new RegExp("\\b" + ss.getActiveCell().getA1Notation() + "\\b");

  for (var s = 0; s < sheets.length ; s++ ) {
    var sheet = sheets[s];
    var range = sheet.getDataRange();
    var formulas = range.getFormulas();

    for (var i = 0; i < formulas.length; i++){
      var row = formulas[i];

      for (var j = 0; j < row.length; j++){
        var cellFormula = row[j].replace(/\$/g, "");
        if (regex.test(cellFormula)){
          var cell = range.getCell(i+1, j+1);
          var cellRef = cell.getA1Notation();
          sheet.getRange(cellRef).activate();
          return 
        }
      }
    }
  }
  SpreadsheetApp.getUi().alert('No dependencies found');
}

function trace_precedents(){
  console.log('this is under development');
  Logger.log('this is under development');
}
