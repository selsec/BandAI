function calculateStartDues() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  
    // Iterate over each sheet and update the starting dues
    sheets.forEach(function(sheet) {
      // Assuming student sheets have names that are not 'roster' or other utility sheet names
      if (sheet.getName() !== 'roster' && sheet.getName() !== 'anotherUtilitySheetName') {
        // Check if "Fair Share" and "Uniform Fee" already exist in the sheet
        var transactionsRange = sheet.getRange('B17:B' + sheet.getLastRow());
        var transactions = transactionsRange.getValues();
        var fairShareExists = transactions.some(function(row) { return row[0] === 'Fair Share'; });
        var uniformFeeExists = transactions.some(function(row) { return row[0] === 'Uniform Fee'; });
  
        // If neither "Fair Share" nor "Uniform Fee" exist, add them
        if (!fairShareExists && !uniformFeeExists) {
          // Set the current date in A17 and A18
          sheet.getRange('A17:A18').setValues([[currentDate], [currentDate]]);
          
          // Set the starting dues information
          sheet.getRange('B17').setValue('Fair Share');
          sheet.getRange('C17').setValue(400);
          sheet.getRange('D17').setValue('Debt');
          
          sheet.getRange('B18').setValue('Uniform Fee');
          sheet.getRange('C18').setValue(50);
          sheet.getRange('D18').setValue('Debt');
          
          // Format the cells if needed
          sheet.getRange('C17:C18').setNumberFormat('$#,##0.00');
        }
      }
    });
  }
  function updateDuesBasedOnGradeAndInstrument() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  
    sheets.forEach(function(sheet) {
      // Assuming the first sheet is not a student sheet
      if (sheet.getName().toLowerCase() !== 'roster') {
        var grade = sheet.getRange('C2').getValue();
        var instrument = sheet.getRange('E2').getValue();
  
        // Check if the grade is 9 and add "Shoes" and "Bibbers"
        if (grade === 9) {
          sheet.getRange('A19:A20').setValues([[currentDate], [currentDate]]);
          sheet.getRange('B19').setValue('Shoes');
          sheet.getRange('B20').setValue('Bibbers');
          sheet.getRange('C19').setValue(60); // Assuming the cost for Shoes is $400
          sheet.getRange('C20').setValue(100);  // Assuming the cost for Bibbers is $50
          sheet.getRange('D19:D20').setValues([['Debt'], ['Debt']]);
        }
  
        // Check the checkboxes in cells A9 and B9
        sheet.getRange('A9').insertCheckboxes().check();
        sheet.getRange('B9').insertCheckboxes().check();
  
        // Check if the instrument is "percussion" and add "Percussion Fee"
        if (instrument.toLowerCase() === 'percussion') {
          sheet.getRange('A21').setValue(currentDate);
          sheet.getRange('B21').setValue('Percussion Fee');
          sheet.getRange('C21').setValue(150); // Assuming the cost for Percussion Fee is $150
          sheet.getRange('D21').setValue('Debt');
        }
      }
    });
  }
  
  // Run the function to update dues based on grade and instrument
  updateDuesBasedOnGradeAndInstrument();
  