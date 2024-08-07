function calculateStartDues() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  
    // Iterate over each sheet and update the starting dues
    sheets.forEach(function(sheet) {
      var sheetName = sheet.getName();
      // Assuming student sheets have names that are not 'roster' or other utility sheet names
      if (sheetName !== "Master" && sheetName !== "Bus Roster" && sheetName !== "Period Roster" && sheetName !== "Attendance" && sheetName !== "Uniform Order"){
        // Check if "Fair Share" and "Uniform Fee" already exist in the sheet
        var transactionsRange = sheet.getRange('B17:B18' + sheet.getLastRow());
        var transactions = transactionsRange.getValues();
        var fairShareExists = transactions.some(function(row) { return row[0] === 'Band Fee'; });
        var uniformFeeExists = transactions.some(function(row) { return row[0] === 'Uniform Fee'; });
        
  
        // If neither "Fair Share" nor "Uniform Fee" exist, add them
        if (!fairShareExists && !uniformFeeExists) {
          // Set the current date in A17 and A18
          sheet.getRange('A17:A19').setValues([[currentDate], [currentDate], [currentDate]]);
          
          // Set the starting dues information
          sheet.getRange('B17').setValue('Band Fee');
          sheet.getRange('C17').setValue(400);
          sheet.getRange('D17').setValue('Debt');
          
          sheet.getRange('B18').setValue('Uniform Fee');
          sheet.getRange('C18').setValue(50);
          sheet.getRange('D18').setValue('Debt');

          sheet.getRange('B19').setValue('Marching Fee');
          sheet.getRange('C19').setValue(200);
          sheet.getRange('D19').setValue('Debt');
          
          // Format the cells if needed
          sheet.getRange('C17:C50').setNumberFormat('$#,##0.00');
          // Perform additional operations
          var grade = sheet.getRange('C2').getValue();
          var instrument = sheet.getRange('E2').getValue();
          // Check if the grade is 9 and add "Shoes" and "Bibbers"
          if (grade === 9) {
            sheet.getRange('A20:A21').setValues([[currentDate], [currentDate]]);
            sheet.getRange('B20').setValue('Shoes');
            sheet.getRange('B21').setValue('Bibbers');
            sheet.getRange('C20').setValue(60); 
            sheet.getRange('C21').setValue(100);  
            sheet.getRange('D20:D21').setValues([['Debt'], ['Debt']]);
            // Check the checkboxes in cells A9 and B9
            sheet.getRange('A9').insertCheckboxes().check();
            sheet.getRange('B9').insertCheckboxes().check();
            sheet.getRange('H10').insertCheckboxes().check();
          }
        
          // Check if the instrument is "percussion" and add "Percussion Fee"
          if (instrument.toLowerCase() === 'percussion') {
          var lastRow = sheet.getLastRow(); // Get the last row with content
          var nextAvailableRow = lastRow + 1; // Calculate the next available row

          // Ensure the next available row is at least 18
          nextAvailableRow = Math.max(nextAvailableRow, 18);

          // Set the current date and other information in the next available row
          sheet.getRange('A' + nextAvailableRow).setValue(currentDate);
          sheet.getRange('B' + nextAvailableRow).setValue('Percussion Fee');
          sheet.getRange('C' + nextAvailableRow).setValue(150); 
          sheet.getRange('D' + nextAvailableRow).setValue('Debt');
          }

        }
      }
    });
  }

  


