const d = new Date();
var $data = 'A4:H10000'
var fuelData = 'B3:J10000'


function test() {
  var ss = SpreadsheetApp.getActive().getActiveSheet();
  var data1 = ss.getRange('B7:B100').getValues();
  var blanks = 0;
  for (let i=0; i < data1.length;i++) {
    if (data1[i][0] === true || data1[i][0] === false) {
      var row = i+7

      ss.getRange(row, 5).activate(); // To Go
      ss.getRange(row, 3, 1, 3).activate(); // Sparkline
      ss.getRange(row, 7).activate(); // Activity
      ss.getRange(row, 9, 1, 6).activate(); // Bulk Formulas

      blanks = 0;
    } else {
      blanks++;
    }
    if (blanks > 3) {break}
  }
}


function equation(x, row=0) {
  var dict = {
    'assigned' : '=0' ,
    'activity' : '=SUMIFS(Expenses!$G$4:$G,Expenses!$E$4:$E,C'+row+',Expenses!$C$4:$C,">="&EOMonth($B$1,-1)+1,Expenses!$C$4:$C,"<="&EOMONTH($B$1,0))' ,
    'available' : '=ROUND(F'+row+' + G'+row+' + L'+row+', 3)' ,
    'targetSum' : '=IF(H'+(row+1)+' = "Monthly",H'+row+' + L'+row+', MAX(H'+row+', I'+row+'))' ,
    'remaining' : '=IF(ISBLANK(H'+(row+1)+'),H'+row+'-I'+row+',\nIF(H'+(row+1)+'="Monthly",H'+row+'-F'+row+',\n(H'+row+'-I'+row+'+F'+row+')/(MONTHS(H'+(row+1)+',$B$1)+1)-F'+row+'))' ,
    'lastMonth' : '=IFERROR(XLOOKUP(M'+row+', INDIRECT($L$4&"!C5:C"), INDIRECT($L$4&"!I5:I"),0))' ,
    'difference' : '=I'+row+'-L'+row ,
    'reference' : '=C' + row
  }

  return dict[x]
}


function updateFormulas() {
  var ss = SpreadsheetApp.getActiveSheet();
  var data1 = ss.getRange('B7:B100').getValues();
  var blanks = 0;
  for (let i=0; i < data1.length;i++) {
    if (data1[i][0] === true || data1[i][0] === false) {
      var row = i+7

      // ss.getRange('E5').copyTo(ss.getActiveSheet().getRange(row, 5), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false); // To Go
      // ss.getRange('C6:E6').copyTo(ss.getActiveSheet().getRange(row, 3, 1, 3), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false); // Sparkline
      // ss.getRange('G5').copyTo(ss.getActiveSheet().getRange(row, 7), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false); // Activity
      // ss.getRange('I5:N5').copyTo(ss.getActiveSheet().getRange(row, 9, 1, 6), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false); // Bulk Formulas
      ss.getRange('E5').copyTo(ss.getRange(row, 5)); // To Go
      ss.getRange('C6:E6').copyTo(ss.getRange(row+1, 3, 1, 3)); // Sparkline
      ss.getRange('G5').copyTo(ss.getRange(row, 7)); // Activity
      ss.getRange('I5:N5').copyTo(ss.getRange(row, 9, 1, 6)); // Bulk Formulas

      blanks = 0;
    } else {
      blanks++;
    }
    if (blanks > 3) {break}
  }
}


function newMonth() {
  var ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(ss.getSheets()[8]);

  // Duplicate the latest month sheet
  ss.duplicateActiveSheet();

  // get name of latest month
  // duplicate sheet
  // update name with month adjustment
  // reset contents
  
  // duplicate the latest month sheet and update name
  var title = ss.getRange('B1').getFormula().toString();
  var month;
  if(title.charAt(12)==','){
    month = parseInt(title.charAt(11))+1;
  } else month = parseInt(title.slice(11,13))+1;
  var year = parseInt(title.slice(6,10));
  if(month>12){
    month=1;
    year++;
  }
  ss.getRange('B1').setFormula('=Date('+year+','+month+',1)');
  ss.getActiveSheet().setName(monthAbr(month)+' '+year.toString().slice(2,4));
  ss.moveActiveSheet(9);

  // clear contents
  var checkbox = ss.getRange('B5:B100').getValues();
  var assigned = ss.getRange('F5:F100').getFormulas();
  var ref = ss.getRange('N5:N100').getFormulas();
  var blanks = 0;
  for (let i=0; i < checkbox.length;i++) {
    if (checkbox[i][0] === true || checkbox[i][0] === false) {
      var row = i+5

      checkbox[i][0] = false;
      assigned[i][0] = '=0';
      ref[i][0] = '=C'+row;
      blanks = 0;
    } else {
      blanks++;
    }
    if (blanks > 3) {break}
  }
  ss.getRange('B5:B100').setValues(checkbox);
  ss.getRange('F5:F100').setFormulas(assigned);
  ss.getRange('N5:N100').setFormulas(ref);
};


function newEntry() {
  // Quickly add a new line at the top of Expenses.
  var ss = SpreadsheetApp.getActive();
  //var month = d.getMonth(Date.now())+1;
  var date = Date.now()+1;
  var expenses = ss.getSheetByName('Expenses');
  sortExpenses();
  var data = ss.getRange($data).getValues();
  x = rowLoop(data, 2)+10;
  expenses.getRange('C'+x).setValue(date);
  sortExpenses()
  expenses.getRange('C4').setValue('');
  expenses.getRange('A4').activate();
};


function transactionProcess() { // process transaction
  // Place the input data into expenses sheet.
  var ss = SpreadsheetApp.getActive();
  var input = ss.getSheetByName('Input')
  var expenses = ss.getSheetByName('Expenses')
  sortExpenses();
  
  // gather user input data
  var data = input.getRange('B2:E9').getValues();

  var account = data[3][2]
  var flag = data[7][2]
  var date = data[4][2]
  var payee = data[1][2]
  var category = data[2][2]
  var memo = data[5][2]
  if (data[0][0]) {var flow = data[0][1]}
    else {var flow = -data[0][1]}
  var cleared = data[6][3]
  var split = false

  // check if it's income
  if (category == 'Ready to Assign' && flag == 'Green') {
    split = true
    flag = 'Green'
    var cTithing = '💵Donations'
    var cSavings = '🏛Savings'
    var mTithing = '(2/3) '+memo
    var mSavings = '(3/3) '+memo
    memo = '(1/3) '+memo
    var flow2 = Math.round(flow*10)/100
    flow -= flow2*2

    var extraData = [
      [account, flag, date, payee, cTithing, mTithing, flow2, cleared],
      [account, flag, date, payee, cSavings, mSavings, flow2, cleared]
    ]
  }

  var dataFormatted = [[account, flag, date, payee, category, memo, flow, cleared]]


  // find empty rows in expenses
  var dataSheet = expenses.getRange($data).getValues()
  var x = rowLoop(dataSheet, 2)+4;
  
  // paste data
  // try cleared first to verify if data locked
  expenses.getRange(x, 8, 1, 1).setValue(cleared)
  expenses.getRange(x, 1, 1, 8).setValues(dataFormatted)
  if (split) {
    expenses.getRange(x+1, 1, 2, 8).setValues(extraData)
  }
  sortExpenses();

  var defaultData = [ 
    [ false, 0, '', '' ],
    [ 'Payee', '', '', '' ],
    [ 'Category', '', '', '' ],
    [ 'Account', '', '', '' ],
    [ 'Date', '', '=TODAY()', '' ],
    [ 'Memo', '', '', '' ],
    [ 'Cleared', '', '', false ],
    [ 'Flag', '', '', '' ],
    [ '', '', '', ''],
    [ 'Record Transaction', '', '', false]
  ]
  input.getRange('B2:E11').setValues(defaultData)
};


function listProcess() { // process list
  // Place the input data into expenses sheet.
  var ss = SpreadsheetApp.getActive();
  var input = ss.getSheetByName('Input');
  var expenses = ss.getSheetByName('Expenses');
  
  // gather input data
  var sheet = input.getRange('I2:P1000').getValues();
  var data = [];

  // filter empty rows
  for(let i=0; i < sheet.length; i++){
    if(sheet[i][0] == '') {
      var z = i;
      break;
    }
    data[i] = sheet[i]
  };

  // find empty rows in expenses
  sortExpenses();
  var dataSheet = expenses.getRange($data).getValues();
  var x = rowLoop(dataSheet, 2)+4;
  
  Logger.log(data)

  // paste data
  // try cleared first to verify if data locked
  expenses.getRange(x, 8, 1, 1).setValue(false)
  expenses.getRange(x, 1, z, 8).setValues(data);
  sortExpenses();

  input.getRange('I2:O1000').clear({contentsOnly: true});
  input.getRange('P2:P1000').setValue(false);
};


function fuelProcess() { // process fuel stop
  // Place the input data into gas sheet.
  var ss = SpreadsheetApp.getActive();
  var input = ss.getSheetByName('Input');
  var fuel = ss.getSheetByName('Fuel');
  
  // gather user input data
  var data = input.getRange('B16:E24').getValues();

  var date = data[0][2]
  var price = data[1][2]
  if(data[2][2] > 0) {
    var total = data[2][2]
  } else {
    var total = ''
  }
  if (data[3][2] == 'Regular') {
    var grade = 'R'
  } else if (data[3][2] == 'Premium') {
    var grade = 'P'
  }
  var station = data[4][2]
  var location = data[5][2]
  var vehicle = data[6][2]
  var mileage = data[7][2]

  var dataFormatted = [[date, station, location, grade, price, total, vehicle, mileage]]

  // find empty rows in fuel
  sortFuel();
  var dataSheet = fuel.getRange(fuelData).getValues()
  var x = rowLoop(dataSheet)+4;
  
  // paste data
  fuel.getRange(x, 2, 1, 8).setValues(dataFormatted)
  sortFuel();

  var defaultData = [ 
    [ 'Date', '', '=TODAY()', '' ],
    [ '$/Gallon', '', '0', '' ],
    [ 'Total $', '', '0', '' ],
    [ 'Grade', '', '', '' ],
    [ 'Station', '', '', '' ],
    [ 'Location', '', '', '' ],
    [ 'Vehicle', '', '', '' ],
    [ 'Mileage', '', '', '' ],
    [ 'Make Transaction', '', '', false],
    [ '', '', '', ''],
    [ 'Record Fuel Stop', '', '', false]
  ]
  input.getRange('B16:E26').setValues(defaultData)

  // record transaction if marked
  if (vehicle == 'Mini Cooper') {
    var memo = 'Mini Fuel'
  } else {
    var memo = vehicle + ' Fuel'
  }
  var transactionData = [ 
    [ false, total, '', '' ],
    [ 'Payee', '', station, '' ],
    [ 'Category', '', '🚗Transportation', '' ],
    [ 'Account', '', '', '' ],
    [ 'Date', '', date, '' ],
    [ 'Memo', '', memo, '' ],
    [ 'Cleared', '', '', false ],
    [ 'Flag', '', '', '' ],
    [ '', '', '', ''],
    [ 'Record Transaction', '', '', false]
  ]
  if (data[8][3]) {
    input.getRange('B2:E11').setValues(transactionData)
  }
};


function transferProcess() { // process transfer
  // Place the input data into expenses sheet.
  var ss = SpreadsheetApp.getActive();
  var input = ss.getSheetByName('Input')
  var expenses = ss.getSheetByName('Expenses')
  
  // gather user input data
  var data = input.getRange('B31:E39').getValues();

  var account1 = data[1][2]
  var account2 = data[2][2]
  var flag = data[6][2]
  var date = data[3][2]
  var payee1 = 'Transfer: ' + account2
  var payee2 = 'Transfer: ' + account1
  var category = ''
  var memo = data[4][2]
  var flow = data[0][0]
  var cleared = data[5][3]

  var dataFormatted = [
    [account1, flag, date, payee1, category, memo, -flow, cleared],
    [account2, flag, date, payee2, category, memo, flow, cleared]
  ]

  // find empty rows in expenses
  sortExpenses();
  var dataSheet = expenses.getRange($data).getValues()
  var x = rowLoop(dataSheet, 2)+4;
  
  // paste data
  // try cleared first to verify if data locked
  expenses.getRange(x, 8, 1, 1).setValue(cleared)
  expenses.getRange(x, 1, 2, 8).setValues(dataFormatted)
  sortExpenses();

  var defaultData = [ 
    [ 0, '', '', '' ],
    [ 'From', '', '', '' ],
    [ 'To', '', '', '' ],
    [ 'Date', '', '=TODAY()', '' ],
    [ 'Memo', '', '', '' ],
    [ 'Cleared', '', '', false ],
    [ 'Flag', '', '', '' ],
    [ '', '', '', ''],
    [ 'Record Transfer', '', '', false]
  ]
  input.getRange('B31:E39').setValues(defaultData)
};


function reconcile() {
  var ss = SpreadsheetApp.getActive();
  var expenses = ss.getSheetByName('Expenses');
  var rule = SpreadsheetApp.newDataValidation().requireTextEqualTo('🔒').setAllowInvalid(false)

  // gather data
  sortExpenses();
  var data = expenses.getRange($data).getValues();
  var x = rowLoop(data, 2)+4;

  // check account
  // check date with reconciled date
  
  // change checkbox to lock and set data validation
  for (let i=0; i < x; i++) {
    if (data[i][2]) {
      if (data[i][7] === true) {
        expenses.getRange('H'+(i+4)).setValue('🔒').setDataValidation(rule).setHorizontalAlignment('center')
      }
      if (data[i][7] === '') {
        expenses.getRange('H'+(i+4)).setValue(false)
        Logger.log(expenses.getRange('H'+(i+4)).getDataValidations())
      }
    }
  }
  // change data validation
}


function sortExpenses() {
  var ss = SpreadsheetApp.getActive();
  var expenses = ss.getSheetByName('Expenses').getRange($data);
  expenses.sort({column: 3, ascending: false});
}


function sortFuel() {
  var ss = SpreadsheetApp.getActive();
  var fuel = ss.getSheetByName('Fuel').getRange(fuelData);
  fuel.sort([{column: 2, ascending: false}, {column: 5, ascending: false}]);
}


function borders() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveRangeList().setBorder(true, null, true, null, null, null, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID)
};



function UntitledMacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('I7').activate();
  spreadsheet.getRange('I5:N5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('O12').activate();
};

function updateSheet() {
  var ss = SpreadsheetApp.getActiveSheet();
  ss.insertColumnsAfter(12, 1);
  var data = [['Difference',], [,], ['=I5-L5',]]
  ss.getRange('M3:M5').setValues(data);
  ss.getRange('M3:M5').activate();
  updateFormulas();

  // var ss = SpreadsheetApp.getActive();
  // ss.getRange('L:L').activate();
  // ss.getActiveSheet().insertColumnsAfter(ss.getActiveRange().getLastColumn(), 1);
  // ss.getActiveRange().offset(0, ss.getActiveRange().getNumColumns(), ss.getActiveRange().getNumRows(), 1).activate();
  // ss.getRange('M3').activate();
  // ss.getRange('\'Nov 23\'!M3:M5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  };
