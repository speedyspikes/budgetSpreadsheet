function nukeData() {
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


