//this function will create the master sheet during the student import function importStudents.js
function createAndFormatMasterSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheetName = 'Master Roster';
    var masterSheet = ss.getSheetByName(masterSheetName) || ss.insertSheet(masterSheetName, ss.getSheets().length);
  
    //headers for the master sheet
    var headers = [
      'Student Name', 'Band Fee ($300/$400)', 'Colorguard Fee ($400)', 'Uniform Fee ($50)',
      'Percussion Fee ($100)', 'Bibbers ($60)', 'Shoes ($30)', 'Dress ($70)',
      'All County ($10)', 'S&E', 'State', 'Indoor Winds', 'Indoor Guard',
      'Leadership Chord', 'Gloves', 'Chaperone Shirt', 'Extra Show Shirt',
      'Fundraising', 'Senior Banners'
    ];
  
    //set the headers in row 1
    masterSheet.getRange('A1:S1').setValues([headers]);
  
    //formatting for the headers
    masterSheet.getRange('A1:S1').setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#213483')
      .setFontColor('#FFFFFF');
  
    //center align the entire sheet
    masterSheet.getRange('A:S').setHorizontalAlignment('center');
  }
  
  //run the function to create and format the master sheet
 // createAndFormatMasterSheet();