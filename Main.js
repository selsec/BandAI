function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Do The Magic!')
      .addItem('Create Tabs', 'makeTabs')
      .addItem('Add Student Tab', 'addStudentTab')
      .addItem('Update Master', 'transferDataToMaster')
      .addItem('Enter Student Details', 'showInputDialog')
      .addToUi();
  }