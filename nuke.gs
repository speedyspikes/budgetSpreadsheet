function nukeLaunchSequence() {
  var ss = SpreadsheetApp.getActive();
  var backEnd = ss.getSheetByName('Back End');
  var code = nukeCode();
  var box = backEnd.getRange('X23');
  var message = backEnd.getRange('X25');
  var rand = Math.floor(Math.random() * 10000)
  Logger.log('Launch Code: ' + rand)
  
  message.setValue('WAITING FOR LAUNCH CODE')
  x = 0;
  Utilities.sleep(5000)
  while (code != rand) {
    x++;
    if (x % 3 == 0) {
      message.setValue('WAITING FOR LAUNCH CODE .');
    } else if (x % 3 == 1) {
      message.setValue('WAITING FOR LAUNCH CODE ..');
    } else if (x % 3 == 2) {
      message.setValue('WAITING FOR LAUNCH CODE ...');
    }
    Utilities.sleep(1000)
    code = nukeCode()
  }
  message.setValue('LAUNCHING NUKE')
  nukeData()

  Logger.log(code)
  message.setValue('NUKE IMPACTED')
  for (i = 0; i < 10; i++) {
    nukeCode();
  }
  
  resetNuke();
}


function nukeCode(x = false) {
  var ss = SpreadsheetApp.getActive();
  var backEnd = ss.getSheetByName('Back End');
  var code = backEnd.getRange('X23');
  code.activate();
  if (x) {
    code.setValue(x);
  } else {
    return code.getValue();
  }
}


function resetNuke() {
  var ss = SpreadsheetApp.getActive();
  var backEnd = ss.getSheetByName('Back End');
  backEnd.getRange('X23').setValue('\'0000');
  backEnd.getRange('X25').setValue('NUKE INACTIVE');
}


function nukeData() {
  var ss = SpreadsheetApp.getActive();
  var backEnd = ss.getSheetByName('Back End')
  var expenses = ss.getSheetByName('Expenses')
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox()

  // clear all expenses data
  expenses.getRange('A4:G').clear()
  expenses.getRange('H4:H').setDataValidation(rule).setValue(false)
  expenses.getRange('L4:M22').clear({contentsOnly: true})
  expenses.getRange('L4:L6').setValue('1/1/2024')
  backEnd.getRange('D12:D30').clear({contentsOnly: true})

  // create some dummy data
  backEnd.getRange('D12:D14').setValues([['Credit Card'],['Checking'],['Savings']])
  expenses.getRange('A4:H8').setValues([
    ['Credit Card', , '1/5/24', 'Chick-fil-A', '🍽Eat Out', 'Dinner', -12.56, false],
    ['Checking', 'Green', '1/2/24', 'Company LLC', 'Ready to Assign', 'Paycheck', 1522.82, true],
    ['Credit Card', , '1/1/24', , 'Ready to Assign', 'Starting Balance', -58.42, true],
    ['Checking', , '1/1/24', , 'Ready to Assign', 'Starting Balance', 123.45, true],
    ['Savings', , '1/1/24', , 'Ready to Assign', 'Starting Balance', 2015, true]
  ])
  expenses.getRange('M4:M6').setValues([[-58.42],[123.456],[2015]])
}
