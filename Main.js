function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('BandAI Functions')
      .addItem('Import Your Roster', 'importRoster')
      .addItem('Calculate Begining Dues', 'calculateStartDues')
      .addItem('Create/Update Uniform Order', 'uniformOrder')
      .addItem('Create Bus Roster', 'createBusRoster')
      .addItem('Update Uniform Measurements', 'importUniformUpdates')
      .addSubMenu(ui.createMenu('Settings and Maintenance'))
        .addItem('Set Section Colors', 'setUniformColors')
        .addItem('Sort Sheets', 'sortSheetsAlphabetically')
      .addToUi();


  
}



  