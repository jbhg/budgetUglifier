/*
* Note that, to shorten (by no means simplify) the code, I make extensive use of the ternary operator, that is:
* (booleanstatement) ? iftrue : iffalse
* example: var freeTimeToday = ((isTodaySaturday() || isTodaySunday()) ? 12 : 4;s
*/ 

function createAnnualSpending() {
  //Constants
  var MAX_FOREIGN_COL = 13;
  var SORT = true;
  var FRAC_EPSILON = 0.1; // If set to 0, no 'partial' transactions will be counted.
  var dateCol = "A"; // Date of transaction
  var tierCol = "B"; // Set if there is tiered spending.
  var typeCol = "C"; // Transaction description, eg. 'Rent'
  var catCol  = "D";
  var descCol = "E";
  var depositCol = "F";
  var withdrawCol = "G";
  var netCol = "H";
  var cumsumCol = "I";
  var repCol = "J";
  var iterCol = "K";
  var mintCategory = "L";
  var mintType = "M";
  
  var iNumItems = 8;
  var sFrequency = 3;
  var dDate = 0;
  
  // getExpenseId() and getIncomeId() each return Google Spreadsheet URLs 
  // that contain the respective spreadsheets.
  
  // Get source and destination sheets
  var startRow = 2;  // First row of data to process
  var expenseSheet = SpreadsheetApp.openById(_getExpenseId()).getSheetByName("Expenses");
  var incomeSheet  = SpreadsheetApp.openById(_getIncomeId() ).getSheetByName("Income");
  var transactionsSheet = _getTransactionsSheet();

  // Get source data
  var expenseRange = expenseSheet.getRange(startRow, 1, expenseSheet.getLastRow(), MAX_FOREIGN_COL);
  var incomeRange = incomeSheet.getRange(startRow, 1, incomeSheet.getLastRow(), MAX_FOREIGN_COL);
  var data = incomeRange.getValues()
      .concat(expenseRange.getValues())
      ;
  
  // Fetch values for each row in the Range.
  var currentRow = 2,
      startDate, // JS code adds to the date when there are multiple occurrences
      financialColumn, // separate cols for deposit and withdrawal
      bExtraFraction, // boolean; if true, there is a "partial instance" of this expense.
      row;
  
  for (i in data) { // For each budget item
    row = data[i];
    if(row[8] > 0){ // If there is at least one instance of this budget item (0s are possible in this sheet)
      startDate = new Date(row[0]); // Starting date; to be increased if the item repeats.
      financialColumn = row[6] > 0 ? depositCol : withdrawCol; // This item is either a deposit or a withdrawal
      bExtraFraction = (Math.abs(row[8]-Math.floor(row[8])) > FRAC_EPSILON) ? true : false; // Is there a 'fraction' at the end of iterating?
      for(j = 1; j <= Math.floor(row[8]) + (bExtraFraction?1:0); j++){ // Each instance of this expense of expense
        transactionsSheet.getRange(typeCol + currentRow).setValue(row[5]);
        transactionsSheet.getRange(catCol + currentRow).setValue(row[3]);
        transactionsSheet.getRange(descCol + currentRow).setValue(j);
        transactionsSheet.getRange(financialColumn + currentRow)
          .setValue(row[6] 
                    * (row[7]>0 ? row[7] : 1) // Only pay a fraction of the overall cost, eg. 25% of Hulu
                    * ((j>row[8]) ? Math.abs(row[8]-Math.floor(row[8])) : 1) // This is a "half" amount, eg. a half-paycheck for 1 of 2 weeks
                   );
        transactionsSheet.getRange(tierCol+currentRow).setValue(row[4]);
        transactionsSheet.getRange(netCol+currentRow).setFormula("=" + depositCol+currentRow + "+" + withdrawCol+currentRow);
        transactionsSheet.getRange(cumsumCol+currentRow).setFormula("=" + netCol+currentRow + "+" + cumsumCol + (currentRow-1));
        if(row[3] == "Biweekly" || row[3] == "Pay Period")
          transactionsSheet.getRange(dateCol+currentRow).setValue(new Date(startDate.getYear(), startDate.getMonth(), startDate.getDate()+14*j));
        else if(row[3] == "Weekly")
          transactionsSheet.getRange(dateCol+currentRow).setValue(new Date(startDate.getYear(), startDate.getMonth(), startDate.getDate()+7*j));
        else if(row[3] == "Quarterly")
          transactionsSheet.getRange(dateCol+currentRow).setValue(new Date(startDate.getYear(), startDate.getMonth()+3*j, startDate.getDate()));
        else if(row[3] == "Monthly" || row[3] == "Seasonal")
          transactionsSheet.getRange(dateCol+currentRow).setValue(new Date(startDate.getYear(), startDate.getMonth()+j, startDate.getDate()));
        else /*if(row[3] == "Annual" || row[3] == "One-Time")*/ {
          transactionsSheet.getRange(dateCol+currentRow).setValue( startDate );
        }
        currentRow++;
      }
    }
  }
  if(SORT) transactionsSheet.getRange(2, 1, transactionsSheet.getLastRow(), MAX_FOREIGN_COL).sort([1, 4]); // Sort by date, then priority.
  Browser.msgBox("Mischief managed.");
}
/*
Returns the sheet entitled "Transactions (Auto)".
If one does not yet exist, it will be created.
*/
function _getTransactionsSheet()
{
  // Find the right sheet. Clear it if it exists, create it if not.
  var transactionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transactions (Auto)");
  if (transactionsSheet != null)
    transactionsSheet.clearContents();
  else 
    transactionsSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Transactions (Auto)");
 
  var SINGLE_DIGIT = 30;
  var CURRENCY = 75;
  var DATE = 80;

  //Set column widths.
  transactionsSheet
    .setColumnWidth(1,DATE)
    .setColumnWidth(3,450)
    .setColumnWidth(4,SINGLE_DIGIT * 3)
    .setColumnWidth(5,SINGLE_DIGIT)
    .setColumnWidth(6,CURRENCY)
    .setColumnWidth(2,SINGLE_DIGIT)
    .setColumnWidth(7,CURRENCY)
    .setColumnWidth(8,CURRENCY)
    .setColumnWidth(9,CURRENCY)
    ;
  
  var dateCol = "A";
  var typeCol = "C";
  var catCol = "D";
  var descCol = "E";
  var depositCol = "F";
  var withdrawCol = "G";
  var tierCol = "B";
  var netCol = "H";
  var cumsumCol = "I";
  var repCol = "J";
  var iterCol = "K";
  var mintCategory = "L";
  var mintType = "M";
  
  //Set header row.
  transactionsSheet.getRange(dateCol + "1").setValue('Date');
  transactionsSheet.getRange(descCol + "1").setValue('Item');
  transactionsSheet.getRange(repCol + "1").setValue('Type?');
  transactionsSheet.getRange(iterCol + "1").setValue('Iter');
  transactionsSheet.getRange(depositCol + "1").setValue('Deposit');
  transactionsSheet.getRange(withdrawCol + "1").setValue('Withdrawal');
  transactionsSheet.getRange(tierCol + "1").setValue('Tier');  
  transactionsSheet.getRange(netCol + "1").setValue('Net');
  transactionsSheet.getRange(cumsumCol + "1").setValue(0);  
  transactionsSheet.getRange(typeCol + "1").setValue('Type');  
  transactionsSheet.getRange(catCol + "1").setValue('Category');
  transactionsSheet.setFrozenRows(1);
  return transactionsSheet;
}
