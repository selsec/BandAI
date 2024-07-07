function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('BandAI Functions')
    .addSubMenu(ui.createMenu('Start Here'))
      .addItem('Import Your Roster', 'importRoster')
      .addItem('Create/Update Uniform Order', 'uniformOrder')
      .addItem('Create Bus Roster', 'createBusRoster')
      .addItem('Update Uniform Measurements', 'importUniformUpdates')
    .addSubMenu(ui.createMenu('Settings and Maintenance')
      .addItem('Set Section Colors', 'setUniformColors')
      .addItem('Sort Sheets', 'sortSheetsAlphabetically'))
    .addToUi();
}
