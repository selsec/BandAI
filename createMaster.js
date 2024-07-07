//this function will create the master sheet during the student import function importStudents.js
function createAndFormatMasterSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheetName = 'Master';
    var masterSheet = ss.getSheetByName(masterSheetName) || ss.insertSheet(masterSheetName, ss.getSheets().length);
  
    //headers for the master sheet
    var headers = [
      'Student Name', 'Grade', 'Period', 'Instrument', 'Balance Owed', 'Band Fee ($300/$400)', 'Uniform Fee ($50)',
      'Percussion Fee ($100)', 'Marching Fee ($200)', 'Bibbers ($60)', 'Shoes ($30)', 'Dress ($70)',
      'All County ($10)', 'S&E', 'State', 'Indoor Winds', 'Indoor Guard',
      'Leadership Chord', 'Gloves', 'Chaperone Shirt', 'Extra Show Shirt',
      'Fundraising', 'Senior Banners'
    ];
  
    //set the headers in row 1
    masterSheet.getRange('A1:W1').setValues([headers]);
    /*masterSheet.getRange('A1').setValue('Student Name');
    masterSheet.getRange('B1').setValue('Grade');
    masterSheet.getRange('C1').setValue('Period');
    masterSheet.getRange('D1').setValue('Instrument');
    masterSheet.getRange('E1').setValue('Balance');
    masterSheet.getRange('F1').setValue('Band Fee ($300/$400)');
    masterSheet.getRange('G1').setValue('Uniform Fee ($50)');
    masterSheet.getRange('H1').setValue('Percussion Fee ($100)');
    masterSheet.getRange('I1').setValue('Marching Fee ($200)');
    masterSheet.getRange('J1').setValue('Bibbers ($60)');
    masterSheet.getRange('K1').setValue('Shoes ($30)');
    masterSheet.getRange('L1').setValue('Dress ($70)');
    masterSheet.getRange('M1').setValue('All County ($10)');
    masterSheet.getRange('N1').setValue('S&E');
    masterSheet.getRange('O1').setValue('State');
    masterSheet.getRange('P1').setValue('Indoor Winds');
    masterSheet.getRange('Q1').setValue('Indoor Guard');
    masterSheet.getRange('R1').setValue('Leadership Chord');
    masterSheet.getRange('S1').setValue('Gloves');
    masterSheet.getRange('T1').setValue('Chaperone Shirt');
    masterSheet.getRange('U1').setValue('Extra Show Shirt');
    masterSheet.getRange('V1').setValue('Fundraising');
    masterSheet.getRange('W1').setValue('Senior Banners');
    */

  
    //formatting for the headers
    masterSheet.getRange('A1:W1').setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#213483')
      .setFontColor('#FFFFFF');
  
    //center align the entire sheet
    masterSheet.getRange('A:Z').setHorizontalAlignment('center');

    
    //loop through all sheets and set the student names in the master sheet
    ss.getSheets().forEach(function(sheet, index) {
      if (sheet.getName() !== masterSheetName && sheet.getName() !== 'Dashboard' && sheet.getName() !== 'IncomeExpense' && sheet.getName() !== 'Bus Roster' && sheet.getName() !== 'Uniform Order' && sheet.getName() !== '3rd Period' && sheet.getName() !== '5th Period' && sheet.getName() !== 'Master Roster' && sheet.getName() !== 'Attendance') {
               var studentName = sheet.getRange('A2').getValue();
               var studentGrade = sheet.getRange('C2').getValue();
               var studentPeriod = sheet.getRange('B2').getValue();
               var studentInstrument = sheet.getRange('E2').getValue();
        masterSheet.getRange('A' + (index + 1)).setValue(studentName);
        masterSheet.getRange('B' + (index + 1)).setValue(studentGrade);
        masterSheet.getRange('C' + (index + 1)).setValue(studentPeriod);
        masterSheet.getRange('D' + (index + 1)).setValue(studentInstrument);
        masterSheet.getRange('E' + (index + 1)).setValue('=INDIRECT("' + studentName + '!G13")');
        masterSheet.getRange('F' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L1")');
        masterSheet.getRange('G' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L2")');
        masterSheet.getRange('H' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L3")');
        masterSheet.getRange('I' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L4")');
        masterSheet.getRange('J' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L5")');
        masterSheet.getRange('K' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L6")');
        masterSheet.getRange('L' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L7")');
        masterSheet.getRange('M' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L8")');
        masterSheet.getRange('N' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L9")');
        masterSheet.getRange('O' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L10")');
        masterSheet.getRange('P' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L11")');
        masterSheet.getRange('Q' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L12")');
        masterSheet.getRange('R' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L13")');
        masterSheet.getRange('S' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L14")');
        masterSheet.getRange('T' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L15")');
        masterSheet.getRange('U' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L16")');
        masterSheet.getRange('V' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L17")');
        masterSheet.getRange('W' + (index + 1)).setValue('=INDIRECT("' + studentName + '!L18")');
        
      }
    });
    for (var i =1; i < 22; i++){
      masterSheet.autoResizeColumn(i);
    }
  }
  
  