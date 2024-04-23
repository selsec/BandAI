function importUniformUpdates() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var parentFolder = DriveApp.getFileById(activeSpreadsheet.getId()).getParents().next();

  //search for the spreadsheet named 'uniform' in the parent folder
  var files = parentFolder.searchFiles('title="uniform" and mimeType="application/vnd.google-apps.spreadsheet" and trashed=false');
  
  if (!files.hasNext()) {
    throw new Error('Spreadsheet named "uniform" not found.');
  }
  
  var file = files.next();
  var uniformSpreadsheet = SpreadsheetApp.openById(file.getId());
  var uniformSheet = uniformSpreadsheet.getSheets()[0]; // Assuming 'uniform' has only one sheet
  var uniformData = uniformSheet.getDataRange().getValues();
    
    //skip the header row and start from the second row
    for (var i = 1; i < uniformData.length; i++) {
    var studentName = uniformData[i][0] || '';
    var uniformNumbner = uniformData[i][1] || '';
    var shirtSize = uniformData[i][2] || '';
    var HeadSize = uniformData[i][3] || '';
    var chest = uniformData[i][4] || '';
    var waist = uniformData[i][5] || '';
    var hips = uniformData[i][6] || '';
    var bibbers = uniformData[i][7] || '';  
    var gloves = uniformData[i][8] || '';
    var shoes = uniformData[i][9] || '';
    var orderBibbers = uniformData[i][10] || false;
    var orderShoes = uniformData[i][11] || false;
    var orderGloves = uniformData[i][12] || false;
    var orderConcert = uniformData[i][13] || false;
    var orderTie = uniformData[i][14] || false;

    //search for sheet name and fill data
    for (var j = 0; j < uniformSpreadsheet.getSheets().length; j++) {
        var sheet = uniformSpreadsheet.getSheets()[j];
        if (sheet.getName() === studentName) {
            sheet.getRange('A8').setValue(shoes);
            sheet.getRange('B8').setValue(bibbers);
            sheet.getRange('C8').setValue(shirtSize);
            sheet.getRange('D8').setValue(uniformNumbner);
            sheet.getRange('G8').setValue(chest);
            sheet.getRange('H8').setValue(waist);
            sheet.getRange('I8').setValue(hips);
            sheet.getRange('A9').setValue(orderShoes);
            sheet.getRange('B9').setValue(orderBibbers);
            sheet.getRange('E9').setValue(orderTie);
            sheet.getRange('F9').setValue(orderConcert);
            sheet.getRange('H10').setValue(orderGloves);
            sheet.getRange('I10').setValue(gloves);
            break;
        }
        }
    }
}