function createIncomeExpense() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var incomeExpenseSheet = ss.getSheetByName('Master');
    var masterIndex = ss.getSheets().indexOf(incomeExpenseSheet);
    var incomeExpenseSheet = ss.insertSheet('IncomeExpense', masterIndex + 1);
    var masterSheetName = 'Master';

    //headers for the sheet
    var headers = [
        'Student Name', 'Grade', 'Period', 'Instrument', 'Band Fee ($300/$400)', 'Colorguard Fee ($400)', 'Uniform Fee ($50)',
        'Percussion Fee ($100)', 'Bibbers ($60)', 'Shoes ($30)', 'Dress ($70)',
        'All County ($10)', 'S&E', 'State', 'Indoor Winds', 'Indoor Guard',
        'Leadership Chord', 'Gloves', 'Chaperone Shirt', 'Extra Show Shirt',
        'Fundraising', 'Senior Banners'
      ];
    
    //set the headers in row 1
    incomeExpenseSheet.getRange('A1:V1').setValues([headers]);
    incomeExpenseSheet.getRange('A1').setValue('Student Name');
    incomeExpenseSheet.getRange('B1').setValue('Grade');
    incomeExpenseSheet.getRange('C1').setValue('Period');
    incomeExpenseSheet.getRange('E1').setValue('Instrument');
    incomeExpenseSheet.getRange('F1').setValue('Band Fee/CG ($300/$400)');
    incomeExpenseSheet.getRange('G1').setValue('Uniform Fee ($50)');
    incomeExpenseSheet.getRange('H1').setValue('Percussion Fee ($100)');
    incomeExpenseSheet.getRange('I1').setValue('Bibbers ($60)');
    incomeExpenseSheet.getRange('J1').setValue('Shoes ($30)');
    incomeExpenseSheet.getRange('K1').setValue('Dress ($70)');
    incomeExpenseSheet.getRange('L1').setValue('All County ($10)');
    incomeExpenseSheet.getRange('M1').setValue('S&E');
    incomeExpenseSheet.getRange('N1').setValue('State');
    incomeExpenseSheet.getRange('O1').setValue('Indoor Winds');
    incomeExpenseSheet.getRange('P1').setValue('Indoor Guard');
    incomeExpenseSheet.getRange('Q1').setValue('Leadership Chord');
    incomeExpenseSheet.getRange('R1').setValue('Gloves');
    incomeExpenseSheet.getRange('S1').setValue('Chaperone Shirt');
    incomeExpenseSheet.getRange('T1').setValue('Extra Show Shirt');
    incomeExpenseSheet.getRange('U1').setValue('Fundraising');
    incomeExpenseSheet.getRange('V1').setValue('Senior Banners');
    
    //set the formula for row 2 to sum the corresponding header from the master sheet and place the value here
    headers.forEach(function(header, index) {
        var formula = '=SUM(' + masterSheetName + '!' + header + '2:' + header + ')';
        incomeExpenseSheet.getRange(2, index + 1).setFormula(formula);
    });
    
    //formatting for the headers
    incomeExpenseSheet.getRange('A1:V1').setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#213483')
      .setFontColor('#FFFFFF');
    
    //center align the entire sheet
    incomeExpenseSheet.getRange('A:Z').setHorizontalAlignment('center');

    // Format the cells directly under each header as currency
    for (var i = 2; i <= headers.length + 1; i++) {
    incomeExpenseSheet.getRange(2, i).setNumberFormat('$#,##0.00');
    }

    //setup dashboard
    incomeExpenseSheet.getRange('A5').setValue('Total Fees Paid');
    incomeExpenseSheet.getRange('C5').setValue('Total Income + Fees');
    incomeExpenseSheet.getRange('E5').setValue('Total Expenses');
    incomeExpenseSheet.getRange('G5').setValue('Balance');

    //formatting for the dashboard
    sheet.getRange("A5").setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBackground('#213483')
        .setFontColor('#FFFFFF');
    sheet.getRange("C5").setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBackground('#213483')
        .setFontColor('#FFFFFF');
    sheet.getRange("E5").setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBackground('#213483')
        .setFontColor('#FFFFFF');
    sheet.getRange("G5").setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBackground('#213483')
        .setFontColor('#FFFFFF');

    //set the formula for the dashboard
    var totalFeesPaid = '=SUM(A2:S2)';
    var totalIncomeFees = '=SUM(A6, C11)';
    var totalExpenses = '=SUM(E11:E)';
    var balance = '=C6-E6';
    incomeExpenseSheet.getRange('A6').setFormula(totalFeesPaid);
    incomeExpenseSheet.getRange('C6').setFormula(totalIncomeFees);
    incomeExpenseSheet.getRange('E6').setFormula(totalExpenses);
    incomeExpenseSheet.getRange('G6').setFormula(balance);
    


}