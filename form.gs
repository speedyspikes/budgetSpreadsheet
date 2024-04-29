// Depricated
// var form = FormApp.openById('1xlIMl8Zq7hgCqKWEsZQzGpWjDS9S2PgmhBhXG5c697o')


function identifyQuestions() {
  var questions = form.getItems()
  questions.forEach(question=>{
    Logger.log(question.getTitle())
    Logger.log(question.getType())
    Logger.log(question.getId().toString())
  })
}


function updateQuestions() {
  var ss = SpreadsheetApp.getActive();
  var backEnd = ss.getSheetByName('Back End')
  var accountList = backEnd.getRange('A12:A30').getValues().filter(row=>row[0]).map(row=>row[0])
  Logger.log(accountList)
  var categList = backEnd.getRange('B33:B60').getValues().filter(row=>row[0]).map(row=>row[0])
  Logger.log(categList)

  var accounts = form.getItemById('849650745')
  accounts.asListItem().setChoiceValues(accountList)
  var categories = form.getItemById('2112236567')
  categories.asListItem().setChoiceValues(categList)
}


/* 4:11:39 PM	Info	Select
4:11:39 PM	Info	MULTIPLE_CHOICE
4:11:39 PM	Info	683510116
4:11:39 PM	Info	Transaction
4:11:39 PM	Info	PAGE_BREAK
4:11:39 PM	Info	628128790
4:11:39 PM	Info	Account
4:11:39 PM	Info	LIST
4:11:39 PM	Info	849650745
4:11:39 PM	Info	Date
4:11:39 PM	Info	DATE
4:11:39 PM	Info	1146961676
4:11:39 PM	Info	Category
4:11:39 PM	Info	LIST
4:11:39 PM	Info	2112236567
4:11:39 PM	Info	Description
4:11:39 PM	Info	TEXT
4:11:39 PM	Info	411759865
4:11:39 PM	Info	Price
4:11:39 PM	Info	TEXT
4:11:39 PM	Info	1575329696
4:11:39 PM	Info	Gas
4:11:39 PM	Info	PAGE_BREAK
4:11:39 PM	Info	826236105
4:11:39 PM	Info	Station
4:11:39 PM	Info	TEXT
4:11:39 PM	Info	2014868971
4:11:39 PM	Info	Location
4:11:39 PM	Info	TEXT
4:11:39 PM	Info	2135411101
4:11:39 PM	Info	Odometer
4:11:39 PM	Info	TEXT
4:11:39 PM	Info	2097594525
4:11:39 PM	Info	Price/Gal
4:11:39 PM	Info	TEXT
4:11:39 PM	Info	1921192165
4:11:39 PM	Info	Total Price
4:11:39 PM	Info	TEXT
4:11:39 PM	Info	859339297
4:11:39 PM	Info	Notes
4:11:39 PM	Info	PARAGRAPH_TEXT
4:11:39 PM	Info	607512413 */
