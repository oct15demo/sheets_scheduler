/** @OnlyCurrentDoc */

function green() {
  var activeSheet = SpreadsheetApp.getActive();
  var selection = activeSheet.getSelection();
  selection.getActiveRangeList().setBackground('#6aa84f');
};

function alternate() {
  var activeSheet = SpreadsheetApp.getActive();
  var selection = activeSheet.getSelection();
  selection.getActiveRangeList().setBackground('#d9ead3');
};

// adapted from https://stackoverflow.com/questions/52085571/looping-through-cells-in-a-range-in-google-sheets-with-google-script
function clear() {
  var activeSheet = SpreadsheetApp.getActive();
  var selection = activeSheet.getSelection();
  var range = selection.getActiveRange()
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var startRow = range.getRow();
  var startCol = range.getColumn();

  for (var i = 0; i < numRows; i++) {
    for (var j = 0; j < numCols; j++) {
      row = startRow + i
      col = startCol + i
      if(row%2 ==0){
        range.getCell(i+1,j+1).setBackground('#f3f3f3')
      } else {
        range.getCell(i+1,j+1).setBackground('white')
      }
      
    }
  }
}

function setPermissions(){
  debug = true 
  if (debug) {
    persons = ['red','green','orange']
  } else {
    persons = ['person1@google.com','person2@google.com','person3@gmail.com',]
  }
  colsPerDay = persons.length
  if(debug)Browser.msgBox(colsPerDay, Browser.Buttons.OK_CANCEL);
  daysPerWeek = 7
  weeks = 2
  rowsBetweenWeeks = 28
  var activeSheet = SpreadsheetApp.getActive();
  var sheet = activeSheet.getSheets()[0];
  startRow = 8
  startCol = 3
  rowsPerDay = 22
  for (var wknum = 0; wknum < weeks; wknum++){
    for (var day = 0; day < daysPerWeek; day++){
      for (var perIdx = 0; perIdx < colsPerDay; perIdx++){ //person index of person array
        rangeRow = startRow + (wknum * rowsBetweenWeeks)
        rangeCol = startCol + (colsPerDay * day) + perIdx
        range = sheet.getRange(rangeRow, rangeCol, rowsPerDay, 1)
        if (debug) {
          range.setBackground(persons[perIdx]);
        } else {
          rangeProtected = range.protect()
          rangeProtected.addEditor(persons[perIdx]);
        }
      }
    }
  }
}
//https://davidmeindl.com/google-sheets-convert-column-index-to-column-letter/
//=SUBSTITUTE(ADDRESS(1, 10, 4), 1, "")
//https://stackoverflow.com/questions/66769617/why-am-i-getting-this-error-the-parameters-string-number-number-number-dont

