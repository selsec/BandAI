//this function will create the master sheet during the student import function importStudents.js
function createAndFormatMasterSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheetName = 'Master Roster';
    var masterSheet = ss.getSheetByName(masterSheetName) || ss.insertSheet(masterSheetName, ss.getSheets().length);
  
    //headers for the master sheet
    var headers = [
      'Student Name', 'Grade', 'Period', 'Instrument', 'Band Fee ($300/$400)', 'Colorguard Fee ($400)', 'Uniform Fee ($50)',
      'Percussion Fee ($100)', 'Bibbers ($60)', 'Shoes ($30)', 'Dress ($70)',
      'All County ($10)', 'S&E', 'State', 'Indoor Winds', 'Indoor Guard',
      'Leadership Chord', 'Gloves', 'Chaperone Shirt', 'Extra Show Shirt',
      'Fundraising', 'Senior Banners'
    ];
  
    //set the headers in row 1
    masterSheet.getRange('A1:V1').setValues([headers]);
  
    //formatting for the headers
    masterSheet.getRange('A1:V1').setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#213483')
      .setFontColor('#FFFFFF');
  
    //center align the entire sheet
    masterSheet.getRange('A:Z').setHorizontalAlignment('center');

    //set the column widths
    masterSheet.setColumnWidth(1, 200);
    masterSheet.setColumnWidth(2, 200);
    masterSheet.setColumnWidth(3, 200);
    masterSheet.setColumnWidth(4, 200);
    masterSheet.setColumnWidth(5, 200);
    masterSheet.setColumnWidth(6, 200);
    masterSheet.setColumnWidth(7, 200);
    masterSheet.setColumnWidth(8, 200);
    masterSheet.setColumnWidth(9, 200);
    masterSheet.setColumnWidth(10, 200);
    masterSheet.setColumnWidth(11, 200);
    masterSheet.setColumnWidth(12, 200);
    masterSheet.setColumnWidth(13, 200);
    masterSheet.setColumnWidth(14, 200);
    masterSheet.setColumnWidth(15, 200);
    
    //loop through all sheets and set the student names in the master sheet
    ss.getSheets().forEach(function(sheet, index) {
      if (sheet.getName() !== masterSheetName && sheet.getName() !== 'Dashboard' && sheet.getName() !== 'Income/Expense' && sheet.getName() !== 'Bus Roster' && sheet.getName() !== 'Uniform Order' && sheet.getName() !== '3rd Period' && sheet.getName() !== '5th Period' && sheet.getName() !== 'Master Roster' && sheet.getName() !== 'Attendance') {
               var studentName = sheet.getRange('A2').getValue();
               var studentGrade = sheet.getRange('C2').getValue();
               var studentPeriod = sheet.getRange('B2').getValue();
        masterSheet.getRange('A' + (index + 2)).setValue(studentName);
        masterSheet.getRange('B' + (index + 2)).setValue(studentGrade);
        masterSheet.getRange('C' + (index + 2)).setValue(studentPeriod);
      }
    });
  }
  
  //run the function to create and format the master sheet
 // createAndFormatMasterSheet();