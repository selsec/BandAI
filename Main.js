function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Start Here')
      .addItem('Create Tabs', 'importRoster')
      .addItem('Sort Tabs', 'sortSheetsAlphabetically')
      .addItem('Calculate Begining Dues', 'calculateStartDues')
      .addItem('Create/Update Uniform Order', 'uniformOrder')
      .addToUi();
  
}
var fluteColor = "Yellow";
var clarinetColor = "Red";
var saxophoneColor = "Blue";
var trumpetColor = "White";
var colorguardColor = "Pink";
var mellophoneColor = "Orange";
var lowBrassColor = "Teal";
var tubaColor = "Purple";
var percussionColor = "Green";


  