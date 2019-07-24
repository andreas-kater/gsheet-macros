
function number() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setNumberFormat('#,##0')
};

function gbp() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setNumberFormat('[$£] #,##0')
};

function usd() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setNumberFormat('"$ "#,##0');
};

function eur() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setNumberFormat('[$€] #,##0')
};

function datapoint() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground(null);
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#188038', SpreadsheetApp.BorderStyle.SOLID);
};

function parameter() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBackground(null);
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#0000ff', SpreadsheetApp.BorderStyle.SOLID);
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
  menuEntries.push({ name: "Trace Dependents", functionName: "traceDependents" });
  ss.addMenu("Detective", menuEntries);
}

function trace_precedents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formula = ss.getActiveCell().getFormula();
  var regex = /[=(&]*('\w* *\w*'![A-Z]{1,3}[0-9]{1,4})|(\w* *\w*![A-Z]{1,3}[0-9]{1,4})|([A-Z]{1,3}[0-9]{1,4})[=)&]*/g;
  Logger.log(formula);
  if (formula.charAt(0) == "=") {
    formula = formula.replace("$", "");
    var matches = regex.exec(formula)
    if (matches.length > 0) {
      var goto = matches[1];
      if (!goto) {
        goto = matches[0];
      }
      ss.getRange(goto).activate();
      return
    }
  }
}

function trace_dependents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var regex = new RegExp("\\b" + ss.getActiveCell().getA1Notation() + "\\b");

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var range = sheet.getDataRange();
    var formulas = range.getFormulas();

    for (var i = 0; i < formulas.length; i++) {
      var row = formulas[i];

      for (var j = 0; j < row.length; j++) {
        var cellFormula = row[j].replace(/\$/g, "");
        if (regex.test(cellFormula)) {
          var cell = range.getCell(i + 1, j + 1);
          var cellRef = cell.getA1Notation();
          sheet.getRange(cellRef).activate();
          return
        }
      }
    }
  }
  SpreadsheetApp.getUi().alert('No dependencies found');
}

function sheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

function weekdayShort(cell) {
  return ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'][new Date(cell).getDay()]
}