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
    // Check if the sheets exist and put them first
    var sheetNames = ['Master', 'Income/Expense', 'Bus Roster', 'Uniform Order'];
    var existingSheets = [];

    sheetNames.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        existingSheets.push(sheet);
      }
    });

    existingSheets.forEach(function(sheet) {
      var index = sortedSheets.indexOf(sheet);
      if (index !== -1) {
        sortedSheets.splice(index, 1);
      }
    });

    sortedSheets.unshift(...existingSheets);
    sortedSheets.forEach(function(sheet) {
      var sheetName = sheet.getName();
      if (sheetName.startsWith('Sheet') && /^\d+$/.test(sheetName.slice(5))) {
        ss.deleteSheet(sheet);
      }
    });
  }
  
  // Run the function to sort the sheets
  