//this function will create the master sheet during the student import function importStudents.js
function createAndFormatMasterSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheetName = 'Master';
    var masterSheet = ss.getSheetByName(masterSheetName) || ss.insertSheet(masterSheetName, ss.getSheets().length);
  
    //headers for the master sheet
    var headers = [
      'Student Name', 'Grade', 'Period', 'Instrument', 'Band Fee ($300/$400)', 'Uniform Fee ($50)',
      'Percussion Fee ($100)', 'Marching Fee ($200)', 'Bibbers ($60)', 'Shoes ($30)', 'Dress ($70)',
      'All County ($10)', 'S&E', 'State', 'Indoor Winds', 'Indoor Guard',
      'Leadership Chord', 'Gloves', 'Chaperone Shirt', 'Extra Show Shirt',
      'Fundraising', 'Senior Banners'
    ];
  
    //set the headers in row 1
    masterSheet.getRange('A1:V1').setValues([headers]);
    masterSheet.getRange('A1').setValue('Student Name');
    masterSheet.getRange('B1').setValue('Grade');
    masterSheet.getRange('C1').setValue('Period');
    masterSheet.getRange('E1').setValue('Instrument');
    masterSheet.getRange('F1').setValue('Band Fee/CG ($300/$400)');
    masterSheet.getRange('G1').setValue('Uniform Fee ($50)');
    masterSheet.getRange('H1').setValue('Percussion Fee ($100)');
    masterSheet.getRange('I1').setValue('Marching Fee ($200)');
    masterSheet.getRange('I1').setValue('Bibbers ($60)');
    masterSheet.getRange('J1').setValue('Shoes ($30)');
    masterSheet.getRange('K1').setValue('Dress ($70)');
    masterSheet.getRange('L1').setValue('All County ($10)');
    masterSheet.getRange('M1').setValue('S&E');
    masterSheet.getRange('N1').setValue('State');
    masterSheet.getRange('O1').setValue('Indoor Winds');
    masterSheet.getRange('P1').setValue('Indoor Guard');
    masterSheet.getRange('Q1').setValue('Leadership Chord');
    masterSheet.getRange('R1').setValue('Gloves');
    masterSheet.getRange('S1').setValue('Chaperone Shirt');
    masterSheet.getRange('T1').setValue('Extra Show Shirt');
    masterSheet.getRange('U1').setValue('Fundraising');
    masterSheet.getRange('V1').setValue('Senior Banners');

  
    //formatting for the headers
    masterSheet.getRange('A1:V1').setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#213483')
      .setFontColor('#FFFFFF');
  
    //center align the entire sheet
    masterSheet.getRange('A:Z').setHorizontalAlignment('center');

    
    //loop through all sheets and set the student names in the master sheet
    ss.getSheets().forEach(function(sheet, index) {
      if (sheet.getName() !== masterSheetName && sheet.getName() !== 'Dashboard' && sheet.getName() !== 'Income/Expense' && sheet.getName() !== 'Bus Roster' && sheet.getName() !== 'Uniform Order' && sheet.getName() !== '3rd Period' && sheet.getName() !== '5th Period' && sheet.getName() !== 'Master Roster' && sheet.getName() !== 'Attendance') {
               var studentName = sheet.getRange('A2').getValue();
               var studentGrade = sheet.getRange('C2').getValue();
               var studentPeriod = sheet.getRange('B2').getValue();
               var studentInstrument = sheet.getRange('E2').getValue();
        masterSheet.getRange('A' + (index + 2)).setValue(studentName);
        masterSheet.getRange('B' + (index + 2)).setValue(studentGrade);
        masterSheet.getRange('C' + (index + 2)).setValue(studentPeriod);
        masterSheet.getRange('E' + (index + 2)).setValue(studentInstrument);
      }
    });
    for (var i =1; i < 22; i++){
      masterSheet.autoResizeColumn(i);
    }
  }
  
  