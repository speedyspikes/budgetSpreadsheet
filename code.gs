var run = true

function onOpen() {
  var ui = SpreadsheetApp.getUi()
  ui.createMenu('Functions')
  .addItem('New Entry', 'newEntry')
  .addItem('Sort Expenses', 'sortExpenses')
  .addItem('Reconcile', 'reconcile')
  .addSeparator()
  .addItem('Sort Fuel', 'sortFuel')
  .addSeparator()
  .addItem('New Month', 'newMonth')
  .addItem('Update Formulas', 'updateFormulas')
  .addItem('test', 'test')
  // .addItem('Update Questions', 'updateQuestions')
  .addToUi();

  resetNuke();
};


function onEdit() {
  // When an edit is made, check to see if any of these checkboxes have changed.
  // If so, run the corresponding function.
  var input = SpreadsheetApp.getActive().getSheetByName('Input');
  var transactionBoxR = 'E11'
  var listBoxR = 'R3'
  var fuelBoxR = 'E26'
  var transferBoxR = 'E39'

  if (run) {
    run = false
    var transactionBox = input.getRange(transactionBoxR).getValue();
    if (transactionBox !== true && transactionBox !== false) {
      throw new Error("transactionBox is not pointing to checkbox in Input sheet");
    } else if (transactionBox) {
      input.getRange(transactionBoxR).setValue(false);
      transactionProcess();
      Logger.log("Transaction Processed");
    }

    var listBox = input.getRange(listBoxR).getValue();
    if (listBox !== true && listBox !== false) {
      throw new Error("listBox is not pointing to checkbox in Input sheet");
    } else if (listBox) {
      input.getRange(listBoxR).setValue(false);
      listProcess();
      Logger.log("List Processed");
    }

    var fuelBox = input.getRange(fuelBoxR).getValue();
    if (fuelBox !== true && fuelBox !== false) {
      throw new Error("fuelBox is not pointing to checkbox in Input sheet");
    } else if (fuelBox) {
      input.getRange(fuelBoxR).setValue(false);
      fuelProcess();
      Logger.log("Fuel Processed");
    }

    var transferBox = input.getRange(transferBoxR).getValue();
    if (transferBox !== true && transferBox !== false) {
      throw new Error("transferBox is not pointing to checkbox in Input sheet");
    } else if (transferBox) {
      input.getRange(transferBoxR).setValue(false);
      transferProcess();
      Logger.log("Transfer Processed");
    }
    run = true
  }
};


function nightly() {
  var ss = SpreadsheetApp.getActive();
  ss.getRange('Summary!B6').setFormula('=EOMONTH(TODAY(),-11)');
}


function monthly() {
  // Create a new month sheet unless the month is already created.
  var ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(ss.getSheets()[8]);
  sheet = ss.getSheetName()
  var month = monthNum(sheet.slice(0, 3));
  var monthNow = d.getMonth(Date.now())+2;
  Logger.log('Latest Sheet = ' + sheet
    + '\n' + 'Month = ' + month
    + '\n' + 'Month Now = ' + monthNow)
  // if (month == 1) month = 13;   This caused an issue in creating February
  if (monthNow > month) {
    newMonth();
    Logger.log('New month sheet created')
  } else {
    Logger.log('Month sheet already exists')
  }
};


function monthAbr(num) {
  // Replace month number with abreviation.
  if(num == 1) return 'Jan';
  else if(num == 2) return 'Feb';
  else if(num == 3) return 'Mar';
  else if(num == 4) return 'Apr';
  else if(num == 5) return 'May';
  else if(num == 6) return 'Jun';
  else if(num == 7) return 'Jul';
  else if(num == 8) return 'Aug';
  else if(num == 9) return 'Sep';
  else if(num == 10) return 'Oct';
  else if(num == 11) return 'Nov';
  else if(num == 12) return 'Dec';
  else return 'ERROR';
};


function monthNum(name) {
  // Replace month name with number.
  if(name == 'Jan') return 1;
  else if(name == 'Feb') return 2;
  else if(name == 'Mar') return 3;
  else if(name == 'Apr') return 4;
  else if(name == 'May') return 5;
  else if(name == 'Jun') return 6;
  else if(name == 'Jul') return 7;
  else if(name == 'Aug') return 8;
  else if(name == 'Sep') return 9;
  else if(name == 'Oct') return 10;
  else if(name == 'Nov') return 11;
  else if(name == 'Dec') return 12;
  else return 'ERROR';
};


function rowLoop(data, column = 0, search = '') {
  // Loop through data rows. Return when row is empty.
  for(let i=0; i < data.length; i++){
    if(data[i][column] == search){
      return i;
    };
  };
  Logger.log('rowLoop: search criteria not found')
};


function findEnd(data) {
  // Loop through data rows. Find space between empty row and END.
  for(let i=0;i < data.length; i++){
    if(data[i][0] == 'END'){
      end = i;
      break;
    };
  }

  for(let i = end; i > 0; i--){
    if(data[i][0] != '')
      empty = i+1
      break;
  };
  return(empty, end);
};




