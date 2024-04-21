//the purpose of this function is the create the bus roster, placing brass and percussion on bus 1 and woodwinds and colorguard on bus 2
function createBusRoster() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var excludedSheets = ['Master', 'Bus Roster', 'Period Roster', 'Attendance', 'Uniform Order'];
    var bus1Instruments = ['Trombone', 'Baritone', 'Tuba', 'Trumpet', 'French Horn', 'Percussion'];
    var busRosterSheet = ss.getSheetByName('Bus Roster') || ss.insertSheet('Bus Roster', 1);
    busRosterSheet.clear(); // Clear any existing content
  
    var bus1Names = [];
    var bus2Names = [];
  
    ss.getSheets().forEach(function(sheet) {
      if (excludedSheets.indexOf(sheet.getName()) === -1) {
        var instrument = sheet.getRange('E2').getValue();
        var studentName = sheet.getName();
        if (bus1Instruments.indexOf(instrument) !== -1) {
          bus1Names.push(studentName);
        } else {
          bus2Names.push(studentName);
        }
      }
    });
  
    // Sort names alphabetically
    bus1Names.sort();
    bus2Names.sort();
  
    // Populate the Bus Roster sheet
    bus1Names.forEach(function(name, index) {
      busRosterSheet.getRange('A' + (index + 2)).setValue(name);
    });
    bus2Names.forEach(function(name, index) {
      busRosterSheet.getRange('D' + (index + 2)).setValue(name);
    });
  
    // Format the headers
    busRosterSheet.getRange('A1').setValue('Bus 1');
    busRosterSheet.getRange('D1').setValue('Bus 2');
    busRosterSheet.getRange('A1:D1').setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#213483')
      .setFontColor('#FFFFFF');
  }
  
  // Run the function to create the bus roster
  //createBusRoster();
  