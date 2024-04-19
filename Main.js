function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Do The Magic!')
      .addItem('Create Tabs', 'importRoster')
      .addToUi();
  }