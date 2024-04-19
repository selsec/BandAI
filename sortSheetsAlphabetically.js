//This function will sort all the tabs into alpha order. Everything is based off Last Name, First Name
function sortSheetsAlphabetically() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var sortedSheets = sheets.sort(function(a, b) {
      return a.getName().localeCompare(b.getName());
    });
  
    sortedSheets.forEach(function(sheet, index) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
    });
  }
  
  // Run the function to sort the sheets
  sortSheetsAlphabetically();