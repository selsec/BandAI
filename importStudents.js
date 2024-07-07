function importRoster() {
    //this will import a list of student information and the create a new tab for each student, and format it with basic information
    //get the active spreadsheet and its parent folder
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var parentFolder = DriveApp.getFileById(activeSpreadsheet.getId()).getParents().next();

  //search for the spreadsheet named 'roster' in the parent folder
  var files = parentFolder.searchFiles('title="roster" and mimeType="application/vnd.google-apps.spreadsheet" and trashed=false');
  
  if (!files.hasNext()) {
    throw new Error('Spreadsheet named "roster" not found.');
  }
  
  var file = files.next();
  var rosterSpreadsheet = SpreadsheetApp.openById(file.getId());
  var rosterSheet = rosterSpreadsheet.getSheets()[0]; // Assuming 'roster' has only one sheet
  var rosterData = rosterSheet.getDataRange().getValues();
    
    //skip the header row and start from the second row
    for (var i = 1; i < rosterData.length; i++) {
    var studentName = rosterData[i][0] || '';
    var grade = rosterData[i][1] || '';
    var instrument = rosterData[i][2] || '';
    var ensembles = rosterData[i][3] || '';
    var studentEmail = rosterData[i][4] || '';
    var parentName = rosterData[i][5] || '';
    var parentEmail = rosterData[i][6] || '';
    var studentPeriod = rosterData[i][7] || '';  
    
    /*//create a new sheet for each student
    var studentSheet = activeSpreadsheet.insertSheet(studentName);
    formatStudentSheet(studentSheet); //format the sheet with the format function*/
    // Check if the sheet already exists
    if (sheetExists(studentName)) {
      throw new Error('Sheet already exists for student: ' + studentName);
    }

    // Create a new sheet for each student
    var studentSheet = activeSpreadsheet.insertSheet(studentName);
    formatStudentSheet(studentSheet); // Format the sheet with the format function

    // Set values in the student's sheet
    studentSheet.getRange('A2').setValue(studentName || '');
    studentSheet.getRange('C2').setValue(grade || '');
    studentSheet.getRange('E2').setValue(instrument || '');
    studentSheet.getRange('G2').setValue(ensembles || '');
    studentSheet.getRange('A5').setValue(studentEmail || '');
    studentSheet.getRange('C5').setValue(parentName || '');
    studentSheet.getRange('E5').setValue(parentEmail || '');
    studentSheet.getRange('B2').setValue(studentPeriod || '');

    // Function to check if the sheet already exists
    function sheetExists(sheetName) {
      var sheets = activeSpreadsheet.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getName() === sheetName) {
          return true;
        }
      }
      return false;
    }


    //set values in the student's sheet
    studentSheet.getRange('A2').setValue(studentName || '');
    studentSheet.getRange('C2').setValue(grade || '');
    studentSheet.getRange('E2').setValue(instrument || '');
    studentSheet.getRange('G2').setValue(ensembles || '');
    studentSheet.getRange('A5').setValue(studentEmail || '');
    studentSheet.getRange('C5').setValue(parentName || '');
    studentSheet.getRange('E5').setValue(parentEmail || '');
    studentSheet.getRange('B2').setValue(studentPeriod || '');
      
    }
    //place new functions here
    /*create these sheets
    Master
    Dashboard
    Income/Expense
    Bus Roster
    Uniform Order
    3rd Period
    5th Period
    */
  calculateStartDues();
   createAndFormatMasterSheet();
  sortSheetsAlphabetically();
  createIncomeExpense();
}

function formatStudentSheet(sheet) {
    //formats the student sheet, adds headers, and creates the transaction area
    //set up headers
    sheet.getRange('A1').setValue('Student Name');
    sheet.getRange('B1').setValue('Period');
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
    sheet.getRange('E7').setValue('Tie');
    sheet.getRange('F7').setValue('Concert Dress');
    sheet.getRange('G7').setValue('Chest');
    sheet.getRange('H7').setValue('Waist');
    sheet.getRange('I7').setValue('Hips');
    sheet.getRange('h9').setValue('Order Gloves');
    sheet.getRange('I9').setValue('Glove Size');
    sheet.getRange('K1').setValue('Band Fee ($300/$400)');
    sheet.getRange('K2').setValue('Uniform Fee ($50)');
    sheet.getRange('K3').setValue('Percussion Fee ($100)');
    sheet.getRange('K4').setValue('Marching Fee $200');
    sheet.getRange('K5').setValue('Bibbers $60');
    sheet.getRange('K6').setValue('Shoes $30');
    sheet.getRange('K7').setValue('Dress $70');
    sheet.getRange('K8').setValue('All County $10');
    sheet.getRange('K9').setValue('S&E $10');
    sheet.getRange('K10').setValue('State $10');
    sheet.getRange('K11').setValue('Indoor Winds');
    sheet.getRange('K12').setValue('Indoor Guard');
    sheet.getRange('K13').setValue('Leadership Chord');
    sheet.getRange('K14').setValue('Gloves');
    sheet.getRange('K15').setValue('Chaperone Shirt');
    sheet.getRange('K16').setValue('Extra Show Shirt');
    sheet.getRange('K17').setValue('Fundraising');
    sheet.getRange('K18').setValue('Senior Banners');
    
    //insert check boxes for needed uniform items
    sheet.getRange('A9').insertCheckboxes();
    sheet.getRange('B9').insertCheckboxes();
    sheet.getRange('C9').insertCheckboxes();
    sheet.getRange('E9').insertCheckboxes();
    sheet.getRange('F9').insertCheckboxes();
    sheet.getRange('h10').insertCheckboxes();
    
    //format headers
    sheet.getRange('A1:G1').setFontWeight('bold');
    sheet.getRange('A4:G4').setFontWeight('bold');
    sheet.getRange('A7:G7').setFontWeight('bold');
    sheet.getRange('A12:G12').setFontWeight('bold');
    sheet.getRange('H7:I7').setFontWeight('bold');
    sheet.getRange('H9:I9').setFontWeight('bold');
    sheet.getRange('K1:K19').setFontWeight('bold');
        
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
    sheet.getRange('K1:K19').setHorizontalAlignment('center');
    sheet.getRange('L1:L19').setHorizontalAlignment('center');
    sheet.getRange('L1:l19').setNumberFormat('$#,##0.00');
    
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
    var formula3 = '=IFERROR(SUM(FILTER(C17:C, (D17:D <> "") * (LOWER(D17:D) = "fundraised"))), "$0.00")';
    sheet.getRange('E13').setNumberFormat('$#,##0.00');
    sheet.getRange('E13').setFormula(formula3);
    sheet.getRange('E13').setHorizontalAlignment('center');  
    
    //set formula to calculate balance remaining
    var formula4 = '=A13-C13-E13';
    sheet.getRange('G13').setNumberFormat('$#,##0.00'); 
    sheet.getRange('G13').setFormula(formula4);
    sheet.getRange('G13').setHorizontalAlignment('center');
    sheet.getRange('A13:J13').setHorizontalAlignment('center');

    //set formula to put totals in column L
    sheet.getRange('L1').setValue('=sumif(B:B, "Band Fee", C:C)');
    sheet.getRange('L2').setValue('=sumif(B:B, "Uniform Fee", C:C)');
    sheet.getRange('L3').setValue('=sumif(B:B, "Percussion Fee", C:C)');
    sheet.getRange('L4').setValue('=sumif(B:B, "Marching Fee", C:C)');
    sheet.getRange('L5').setValue('=sumif(B:B, "Bibbers", C:C)');
    sheet.getRange('L6').setValue('=sumif(B:B, "Shoes", C:C)');
    sheet.getRange('L7').setValue('=sumif(B:B, "Dress", C:C)');
    sheet.getRange('L8').setValue('=sumif(B:B, "All County", C:C)');
    sheet.getRange('L9').setValue('=sumif(B:B, "S&E", C:C)');
    sheet.getRange('L10').setValue('=sumif(B:B, "State", C:C)');
    sheet.getRange('L11').setValue('=sumif(B:B, "Indoor Winds", C:C)');
    sheet.getRange('L12').setValue('=sumif(B:B, "Indoor Guard", C:C)');
    sheet.getRange('L13').setValue('=sumif(B:B, "Leadership Chord", C:C)');
    sheet.getRange('L14').setValue('=sumif(B:B, "Gloves", C:C)');
    sheet.getRange('L15').setValue('=sumif(B:B, "Chaperone Shirt", C:C)');
    sheet.getRange('L16').setValue('=sumif(B:B, "Extra Show Shirt", C:C)');
    sheet.getRange('L17').setValue('=sumif(B:B, "Fundraising", C:C)');
    sheet.getRange('L18').setValue('=sumif(B:B, "Senior Banners", C:C)');
    sheet.getRange('L1:L19').setNumberFormat('$#,##0.00');
    sheet.getRange('L1:L19').setHorizontalAlignment('center');
    

    
    //set header color
    var headerRange = sheet.getRange('A1:G1');
    headerRange.setBackground('#213483');
    var headerRange2 = sheet.getRange('A4:G4');
    headerRange2.setBackground('#213483');
    var headerRange3 = sheet.getRange('A7:I7');
    headerRange3.setBackground('#213483');
    var headerRange4 = sheet.getRange('A15');
    headerRange4.setBackground('#213483');
    var headerRange5 = sheet.getRange('h9:i9');
    headerRange5.setBackground('#213483');
    var headerRange6 = sheet.getRange('K1:K19');
    headerRange6.setBackground('#213483');

    
    //set header text color
    var headerTextRange = sheet.getRange('A1:G1');
    headerTextRange.setFontColor('#ffffff');
    var headerTextRange2 = sheet.getRange('A4:G4');
    headerTextRange2.setFontColor('#ffffff');
    var headerTextRange3 = sheet.getRange('A7:I7');
    headerTextRange3.setFontColor('#ffffff');
    var headerTextRange4 = sheet.getRange('A15');
    headerTextRange4.setFontColor('#ffffff');
    var headerTextRange5 = sheet.getRange('h9:i9');
    headerTextRange5.setFontColor('#ffffff');
    var headerTextRange6 = sheet.getRange('K1:K19');
    headerTextRange6.setFontColor('#ffffff');

    //center entire sheet
    sheet.getRange('A:Z').setHorizontalAlignment('center');
    
    //auto resize columns
    for (var i = 1; i <= 26; i++) {
      sheet.autoResizeColumn(i);
    }

}
