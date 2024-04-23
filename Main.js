function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Start Here')
      .addItem('Create Tabs', 'importRoster')
      .addItem('Sort Tabs', 'sortSheetsAlphabetically')
      .addItem('Calculate Begining Dues', 'calculateStartDues')
      .addItem('Create/Update Uniform Order', 'uniformOrder')
      .addItem('Create Bus Roster', 'createBusRoster')
      .addItem('Set Section Colors', 'setSectionColors')
      .addItem('Update Uniform Measurements', 'importUniformUpdates')
      .addToUi();
  
}


  