function createIncomeExpense() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var incomeExpense = ss.getSheetByName('Master');
    var masterIndex = ss.getSheets().indexOf(incomeExpense);
    var incomeExpense = ss.insertSheet('IncomeExpense', masterIndex + 1);
    var masterSheetName = 'Master';

    //headers for the sheet
    var headers = [
        'Band Fee ($300/$400)', 'Uniform Fee ($50)', 'Percussion Fee ($100)', 
        'Marching Fee ($200)', 'Bibbers ($60)', 'Shoes ($30)', 'Dress ($70)',
        'All County ($10)', 'S&E', 'State', 'Indoor Winds', 'Indoor Guard',
        'Leadership Chord', 'Gloves', 'Chaperone Shirt', 'Extra Show Shirt',
        'Fundraising', 'Senior Banners'
      ];
    
    //set the headers in row 1
    incomeExpense.getRange('A1:R1').setValues([headers]);
    /*
    incomeExpense.getRange('A1').setValue('Student Name');
    incomeExpense.getRange('B1').setValue('Grade');
    incomeExpense.getRange('C1').setValue('Period');
    incomeExpense.getRange('E1').setValue('Instrument');
    incomeExpense.getRange('F1').setValue('Band Fee/CG ($300/$400)');
    incomeExpense.getRange('G1').setValue('Uniform Fee ($50)');
    incomeExpense.getRange('H1').setValue('Percussion Fee ($100)');
    incomeExpense.getRange('I1').setValue('Bibbers ($60)');
    incomeExpense.getRange('J1').setValue('Shoes ($30)');
    incomeExpense.getRange('K1').setValue('Dress ($70)');
    incomeExpense.getRange('L1').setValue('All County ($10)');
    incomeExpense.getRange('M1').setValue('S&E');
    incomeExpense.getRange('N1').setValue('State');
    incomeExpense.getRange('O1').setValue('Indoor Winds');
    incomeExpense.getRange('P1').setValue('Indoor Guard');
    incomeExpense.getRange('Q1').setValue('Leadership Chord');
    incomeExpense.getRange('R1').setValue('Gloves');
    incomeExpense.getRange('S1').setValue('Chaperone Shirt');
    incomeExpense.getRange('T1').setValue('Extra Show Shirt');
    incomeExpense.getRange('U1').setValue('Fundraising');
    incomeExpense.getRange('V1').setValue('Senior Banners');
    */
    //set the formula for row 2 to sum the corresponding header from the master sheet and place the value here
    
    var headers1 =['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
    headers1.forEach(function(header, index) {
        var formula = '=SUM(' + masterSheetName + '!' + header + '2:' + header + ')';
        incomeExpense.getRange(2, index + 1).setFormula(formula);
    });
    
    //formatting for the headers
    incomeExpense.getRange('A1:R1').setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#213483')
      .setFontColor('#FFFFFF');
    
    //center align the entire sheet
    incomeExpense.getRange('A:Z').setHorizontalAlignment('center');

    // Format the cells directly under each header as currency
    for (var i = 1; i <= headers.length + 1; i++) {
    incomeExpense.getRange(2, i).setNumberFormat('$#,##0.00');
    }

    //setup dashboard
    incomeExpense.getRange('A5').setValue('Total Fees Paid');
    incomeExpense.getRange('C5').setValue('Total Income + Fees');
    incomeExpense.getRange('E5').setValue('Total Expenses');
    incomeExpense.getRange('G5').setValue('Balance');

    //formatting for the dashboard
    incomeExpense.getRange("A5").setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBackground('#213483')
        .setFontColor('#FFFFFF');
    incomeExpense.getRange("C5").setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBackground('#213483')
        .setFontColor('#FFFFFF');
    incomeExpense.getRange("E5").setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBackground('#213483')
        .setFontColor('#FFFFFF');
    incomeExpense.getRange("G5").setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBackground('#213483')
        .setFontColor('#FFFFFF');

    //set the formula for the dashboard
    var totalFeesPaid = '=SUM(A2:S2)';
    var totalIncomeFees = '=SUM(A6, C11)';
    var totalExpenses = '=SUM(E11:E)';
    var balance = '=C6-E6';
    incomeExpense.getRange('A6').setFormula(totalFeesPaid);
    incomeExpense.getRange('C6').setFormula(totalIncomeFees);
    incomeExpense.getRange('E6').setFormula(totalExpenses);
    incomeExpense.getRange('G6').setFormula(balance);

    //format the cells as currency
    incomeExpense.getRange('A6').setNumberFormat('$#,##0.00');
    incomeExpense.getRange('C6').setNumberFormat('$#,##0.00');
    incomeExpense.getRange('E6').setNumberFormat('$#,##0.00');
    incomeExpense.getRange('G6').setNumberFormat('$#,##0.00');

    //set up non-fee income/expense
    incomeExpense.getRange('A10').setValue('Non-Fee Income');
    incomeExpense.getRange('B10').setValue('Source');
    incomeExpense.getRange('C10').setValue('Total Non-Fee Income');

    //formatting for the non-fee income/expense
    incomeExpense.getRange('A10:C10').setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#213483')
      .setFontColor('#FFFFFF');
    incomeExpense.getRange('A11:A').setNumberFormat('$#,##0.00');
    incomeExpense.getRange('C11:C').setNumberFormat('$#,##0.00');
    var totanNonFeeIncome = '=SUM(A11:A)';
    incomeExpense.getRange('C11').setFormula(totanNonFeeIncome);

    //set up expenses
    incomeExpense.getRange('E10').setValue('Expenses');
    incomeExpense.getRange('F10').setValue('PO#');
    incomeExpense.getRange('G10').setValue('Description');
    incomeExpense.getRange('H10').setValue('Total Expenses');

    //formatting for the expenses
    incomeExpense.getRange('E10:H10').setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#213483')
      .setFontColor('#FFFFFF');
    incomeExpense.getRange('E11:E').setNumberFormat('$#,##0.00');
    incomeExpense.getRange('H11:H').setNumberFormat('$#,##0.00');
    var totalExpenses = '=SUM(E11:E)';
    incomeExpense.getRange('H11').setFormula(totalExpenses);

    //resize the columns
    for (var i =1; i < 22; i++){
      incomeExpense.autoResizeColumn(i);
    }
    
}