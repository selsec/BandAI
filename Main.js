function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Start Here')
      .addItem('Create Tabs', 'importRoster')
      .addItem('Sort Tabs', 'sortSheetsAlphabetically')
      .addItem('Calculate Begining Dues', 'calculateStartDues')
      .addToUi();
  
}
  