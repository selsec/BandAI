function importRoster() {
    //This will be the heavy lifter function. It will import a list of student information and the create a new tab for each student, and format it with basic information
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rosterSheet = ss.getSheetByName('roster');
    var rosterData = rosterSheet.getDataRange().getValues();
    
    // Skip the header row and start from the second row
    for (var i = 1; i < rosterData.length; i++) {
      var studentName = rosterData[i][0];
      var grade = rosterData[i][1];
      var instrument = rosterData[i][2];
      var ensembles = rosterData[i][3];
      var studentEmail = rosterData[i][4];
      var parentName = rosterData[i][5];
      var parentEmail = rosterData[i][6];
      
      // Create a new sheet for each student
      var studentSheet = ss.insertSheet(studentName);
      formatStudentSheet(studentSheet); //format the sheet with the format function
      
      // Set values in the student's sheet
      studentSheet.getRange('A2').setValue(studentName);
      studentSheet.getRange('C2').setValue(grade);
      studentSheet.getRange('E2').setValue(instrument);
      studentSheet.getRange('G2').setValue(ensembles);
      studentSheet.getRange('A5').setValue(studentEmail);
      studentSheet.getRange('C5').setValue(parentName);
      studentSheet.getRange('E5').setValue(parentEmail);
    }
}

function formatStudentSheet(sheet) {
    //Formats the student sheet, adds headers, and creates the transaction area
    //set up headers
    sheet.getRange('A1').setValue('Student Name');
    sheet.getRange('A4').setValue('Student Email');
    sheet.getRange('C1').setValue('Grade');
    sheet.getRange('C4').setValue('Parent Name');
    sheet.getRange('E1').setValue('Instrument');
    sheet.getRange('E4').setValue('Parent Email');
    sheet.getRange('G1').setValue('Ensembles');
    sheet.getRange('A8').setValue('Total Debt');
    sheet.getRange('C8').setValue('Total Paid');
    sheet.getRange('E8').setValue('Total Fundraised');
    sheet.getRange('G8').setValue('Balance Remaining');
    //format headers
    sheet.getRange('A1:G1').setFontWeight('bold');
    sheet.getRange('A4:G4').setFontWeight('bold');
    sheet.getRange('A8:G8').setFontWeight('bold');
    sheet.getRange('A1:G1').setHorizontalAlignment('center');
    sheet.getRange('A4:G4').setHorizontalAlignment('center');
    sheet.getRange('A8:G8').setHorizontalAlignment('center');
    //merge and center transaction history header
    var mergeRange =sheet.getRange('A15:G15');
    mergeRange.merge();
    mergeRange.setHorizontalAlignment('center');
    mergeRange.setValue('Transaction History');
    mergeRange.setFontWeight('bold');
    //set up transaction history headers
    sheet.getRange('A16').setValue('Date');
    sheet.getRange('B16').setValue('Description');
    sheet.getRange('C16').setValue('Amount');
    sheet.getRange('D16').setValue('Debt/Payment/Fundraised');
    sheet.getRange('E16').setValue('MyShoolBucks');
    sheet.getRange('F16').setValue('Check Number');
    sheet.getRange('G16').setValue('Reciept Number');
    //format transaction history headers
    sheet.getRange('A16:G16').setFontWeight('bold');
    sheet.getRange('A16:G16').setHorizontalAlignment('center');
    //format transactions area
    sheet.getRange('A17:A').setNumberFormat('MM/dd/yyyy');
    sheet.getRange('C17:C').setNumberFormat('$#,##0.00');
    sheet.getRange('A17:G').setHorizontalAlignment('center');
    //set formula to calculate total debt
    var formula1 = '=SUMIF(LOWER(D17:D), "debt", C17:C)';
    sheet.getRange('A9').setNumberFormat('$#,##0.00');
    sheet.getRange('A9').setFormula(formula1);
    sheet.getRange('A9').setHorizontalAlignment('center');
    //set formula to calculate total paid
    var formula2 = '=SUMIF(LOWER(D17:D), "payment", C17:C)';
    sheet.getRange('C9').setNumberFormat('$#,##0.00');
    sheet.getRange('C9').setFormula(formula2);
    sheet.getRange('C9').setHorizontalAlignment('center');
    //set formula to calculate total fundraised
    var formula3 = '=SUMIF(LOWER(D17:D), "fundraised", C17:C)';
    sheet.getRange('E9').setNumberFormat('$#,##0.00');
    sheet.getRange('E9').setFormula(formula3);
    sheet.getRange('E9').setHorizontalAlignment('center');  
    //set formula to calculate balance remaining
    var formula4 = '=A9-C9-E9';
    sheet.getRange('G9').setNumberFormat('$#,##0.00'); 
    sheet.getRange('G9').setFormula(formula4);
    sheet.getRange('G9').setHorizontalAlignment('center');
    sheet.getDataRange().setHorizontalAlignment('center');
    //set header color
    var headerRange = sheet.getRange('A1:G1');
    headerRange.setBackground('#213483');
    var headerRange2 = sheet.getRange('A4:G4');
    headerRange2.setBackground('#213483');
    var headerRange3 = sheet.getRange('A8:G8');
    headerRange3.setBackground('#213483');
    var headerRange4 = sheet.getRange('A15');
    headerRange4.setBackground('#213483');
}


