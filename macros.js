
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
  var dependents = []
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentCell = ss.getActiveCell();
  var currentCellRef = currentCell.getA1Notation();
  var range = ss.getDataRange();

  var regex = new RegExp("\\b" + currentCellRef + "\\b");
  var formulas = range.getFormulas();

  for (var i = 0; i < formulas.length; i++){
    var row = formulas[i];

    for (var j = 0; j < row.length; j++){
      var cellFormula = row[j].replace(/\$/g, "");
      if (regex.test(cellFormula)){
        dependents.push([i,j]);
      }
    }
  }

  var dependentRefs = [];
  for (var k = 0; k < dependents.length; k ++){
    var rowNum = dependents[k][0] + 1;
    var colNum = dependents[k][1] + 1;
    var cell = range.getCell(rowNum, colNum);
    var cellRef = cell.getA1Notation();
    dependentRefs.push(cellRef);
  }
  var output = "Dependents: ";
  if(dependentRefs.length > 0){
    output += dependentRefs.join(", ");
  } else {
    output += " None";
  }
  currentCell.setNote(output);
}
