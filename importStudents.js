function importRoster() {
    //This will be the heavy lifter function. It will import a list of student information and the create a new tab for each student, and format it with basic information
    // Get the active spreadsheet and its parent folder
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var parentFolder = DriveApp.getFileById(activeSpreadsheet.getId()).getParents().next();

  // Search for the spreadsheet named 'roster' in the parent folder
  var files = parentFolder.searchFiles('title="roster" and mimeType="application/vnd.google-apps.spreadsheet" and trashed=false');
  
  if (!files.hasNext()) {
    throw new Error('Spreadsheet named "roster" not found.');
  }
  
  var file = files.next();
  var rosterSpreadsheet = SpreadsheetApp.openById(file.getId());
  var rosterSheet = rosterSpreadsheet.getSheets()[0]; // Assuming 'roster' has only one sheet
  var rosterData = rosterSheet.getDataRange().getValues();
    
    // Skip the header row and start from the second row
    for (var i = 1; i < rosterData.length; i++) {
    var studentName = rosterData[i][0] || '';
    var grade = rosterData[i][1] || '';
    var instrument = rosterData[i][2] || '';
    var ensembles = rosterData[i][3] || '';
    var studentEmail = rosterData[i][4] || '';
    var parentName = rosterData[i][5] || '';
    var parentEmail = rosterData[i][6] || '';

    
    // Create a new sheet for each student
    var studentSheet = activeSpreadsheet.insertSheet(studentName);
    formatStudentSheet(studentSheet); //format the sheet with the format function
    
    // Set values in the student's sheet
    studentSheet.getRange('A2').setValue(studentName || '');
    studentSheet.getRange('C2').setValue(grade || '');
    studentSheet.getRange('E2').setValue(instrument || '');
    studentSheet.getRange('G2').setValue(ensembles || '');
    studentSheet.getRange('A5').setValue(studentEmail || '');
    studentSheet.getRange('C5').setValue(parentName || '');
    studentSheet.getRange('E5').setValue(parentEmail || '');
      
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
    sheet.getRange('G4').setValue('Period');
    sheet.getRange('A12').setValue('Total Debt');
    sheet.getRange('C12').setValue('Total Paid');
    sheet.getRange('E12').setValue('Total Fundraised');
    sheet.getRange('G12').setValue('Balance Remaining');
    sheet.getRange('A7').setValue('Shoes');
    sheet.getRange('B7').setValue('Bibbers');
    sheet.getRange('C7').setValue('T-Shirt Size');
    sheet.getRange('D7').setValue('Jacket/Shako');
    sheet.getRange('E7').setValue('Chest');
    sheet.getRange('F7').setValue('Waist');
    sheet.getRange('G7').setValue('Hips');
    
    //insert check boxes for needed uniform items
    sheet.getRange('A9').insertCheckboxes();
    sheet.getRange('B9').insertCheckboxes();
    sheet.getRange('C9').insertCheckboxes();
    
    //format headers
    sheet.getRange('A1:G1').setFontWeight('bold');
    sheet.getRange('A4:G4').setFontWeight('bold');
    sheet.getRange('A7:G7').setFontWeight('bold');
    sheet.getRange('A12:G12').setFontWeight('bold');
    
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
    var formula1 = '=SUM(FILTER(C17:C, (D17:D <> "") * (LOWER(D17:D) = "debt")))';
    sheet.getRange('A13').setNumberFormat('$#,##0.00');
    sheet.getRange('A13').setFormula(formula1);
    sheet.getRange('A13').setHorizontalAlignment('center');
    
    //set formula to calculate total paid
    var formula2 = '=SUM(FILTER(C17:C, (D17:D <> "") * (LOWER(D17:D) = "payment")))';
    sheet.getRange('C13').setNumberFormat('$#,##0.00');
    sheet.getRange('C13').setFormula(formula2);
    sheet.getRange('C13').setHorizontalAlignment('center');
    
    //set formula to calculate total fundraised
    var formula3 = '=SUM(FILTER(C17:C, (D17:D <> "") * (LOWER(D17:D) = "fundraised")))';
    sheet.getRange('E13').setNumberFormat('$#,##0.00');
    sheet.getRange('E13').setFormula(formula3);
    sheet.getRange('E13').setHorizontalAlignment('center');  
    
    //set formula to calculate balance remaining
    var formula4 = '=A13-C13-E13';
    sheet.getRange('G13').setNumberFormat('$#,##0.00'); 
    sheet.getRange('G13').setFormula(formula4);
    sheet.getRange('G13').setHorizontalAlignment('center');
    sheet.getRange('A13:J13').setHorizontalAlignment('center');
    
    //set header color
    var headerRange = sheet.getRange('A1:G1');
    headerRange.setBackground('#213483');
    var headerRange2 = sheet.getRange('A4:G4');
    headerRange2.setBackground('#213483');
    var headerRange3 = sheet.getRange('A7:G7');
    headerRange3.setBackground('#213483');
    var headerRange4 = sheet.getRange('A15');
    headerRange4.setBackground('#213483');
    
    //set header text color
    var headerTextRange = sheet.getRange('A1:G1');
    headerTextRange.setFontColor('#ffffff');
    var headerTextRange2 = sheet.getRange('A4:G4');
    headerTextRange2.setFontColor('#ffffff');
    var headerTextRange3 = sheet.getRange('A7:G7');
    headerTextRange3.setFontColor('#ffffff');
    var headerTextRange4 = sheet.getRange('A15');
    headerTextRange4.setFontColor('#ffffff');

    

   
}
