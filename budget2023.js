// Written by Chelly Compendio https://github.com/chellxchell


// For updating the monthly sheets -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function updateMonthlySpending(sheet) {
  var sheet = SpreadsheetApp.getActiveSheet();

  // Change the name of the top row (ex. "[Month Name] Budget" -> "November Budget")
  var budgetName = sheet.getRange(1,1).getDisplayValue().replace("[Month Name]", sheet.getName())
  if (sheet.getName != "Template") sheet.getRange(1,1).setValue(budgetName)

  // Change the names of all the charts (ex. "[Month Name] Spending" -> "November Spending")
  var charts = sheet.getCharts();
  for (var ch in charts) {
    var chart = charts[ch]
    var chartName = chart.getOptions().get("title").replace("[Month Name]", sheet.getName())
    chart = chart.modify()
    .setOption('title', chartName)
    .build();
    sheet.updateChart(chart);
  }

  // Get all the expenses information
  var categories = sheet.getRange(cellMapMonth.get("categoriesList").cell.row, cellMapMonth.get("categoriesList").cell.column, 12).getDisplayValues().map((value) => value[0]);
  var summed_expenses = Array.from(Array(categories.length), () => new Array(1).fill(0))

  // Add recurring, itemized, and travel expenses to the total
  for (var e of ["recurringExpenses","itemizedExpenses", "tripExpenses"]){
    // Get that expense type from the map
    var expenseType = cellMapMonth.get(e)
    // Get and separate expense data
    var expenses_data = sheet.getRange(
      expenseType.cell.row, 
      expenseType.cell.column,
      expenseType.numEntries, 
      3).getValues();
    var expenses_amounts = expenses_data.map(function(value,index) { return value[expenseType.amountCol]; }); // get just the amounts column
    var expenses_categories = expenses_data.map(function(value,index) { return value[expenseType.categoryCol]; }); // get just the categories column

    // Loop through itemized expenses
    for (var i = 0; i < expenses_amounts.length; i++) {
      var value = expenses_amounts[i];
      // Ignore 0 values
      if (value === 0){
        continue
      }
      // Stop when reach a cell with no value
      else if (value === '') {
        break;
      }
      // add to the total
      summed_expenses[categories.indexOf(expenses_categories[i])][0] += value
    }
  }

  sheet.getRange(cellMapMonth.get("summedExpenses").cell.row, cellMapMonth.get("summedExpenses").cell.column, categories.length).setValues(summed_expenses);
}

// For the Yearly Spending sheet -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function updateYearlySpending(sheet){
  var months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  var sheet = SpreadsheetApp.getActiveSheet();

  // expenses by month
  for (var i=0; i<months.length; i++){
    // skip if there's no sheet for that month
    let monthSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(months[i]);
    if (!monthSheet){
      continue
    }

    // overall spending
    let monthSpending = monthSheet.getRange(cellMapMonth.get("totalLeisureSpending").cell.row,cellMapMonth.get("totalLeisureSpending").cell.column).getDisplayValue();
    sheet.getRange(i+4,3).setValue(monthSpending)
    // for the goal monthly spending
    var goalRow = cellMapYear.get("monthlyGoalStart").cell.row
    var goalColumn = cellMapYear.get("monthlyGoalStart").cell.column
    for (var j=goalRow; j<goalRow+12; j++){
      var goalForMonth = monthSheet.getRange(cellMapMonth.get("totalSpendingGoal").cell.row, cellMapMonth.get("totalSpendingGoal").cell.column).getDisplayValue();
      sheet.getRange(j, goalColumn).setValue(goalForMonth)
    }

    // spending by category
    var currRow = i+21
    var currCol = 3
    var categoryCells = ["K6", "K7", "L22", "L23", "L24", "L25", "L26", "L27", "L28", "L29", "L30", "L31", "L32", "L33"]
    for (var c=0; c<categoryCells.length; c++){
      let catSpending = monthSheet.getRange(fromA1Notation(categoryCells[c]).row, fromA1Notation(categoryCells[c]).column).getDisplayValue();
      sheet.getRange(currRow,currCol).setValue(catSpending)

      // account for the two columns of information
      currRow = (c%2==1) ? currRow + 15 : currRow
      currCol = (c%2==1) ? 3 : 6
    }

    // spending for 50/30/20 rule + net worth progression
    currRow = currRow + 2
    currCol = 3
    var breakdownCells = [cellMapMonth.get("essentialsPercent"), cellMapMonth.get("savingsPercent"), cellMapMonth.get("wantsPercent"), cellMapMonth.get("netWorth")]
    for (var b=0; b<breakdownCells.length; b++){
      // test for gas
      let breakSpending = monthSheet.getRange(breakdownCells[b].cell.row, breakdownCells[b].cell.column).getDisplayValue();
      sheet.getRange(currRow,currCol).setValue(breakSpending)

      // account for the two columns of information
      currRow = (b%2==1) ? currRow + 15 : currRow
      currCol = (b%2==1) ? 3 : 6
    }
  }
}

/**-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 * @param {string} cell -  The cell address in A1 notation
 * @returns {object} The row number and column number of the cell (0-based)
 * @example
 *   fromA1Notation("A2") returns {row: 1, column: 3}
 */
const fromA1Notation = (cell) => {
  var [, columnName, row] = cell.toUpperCase().match(/([A-Z]+)([0-9]+)/);
  const characters = 'Z'.charCodeAt() - 'A'.charCodeAt() + 1;

  let column = 0;
  columnName.split('').forEach((char) => {
    column *= characters;
    column += char.charCodeAt() - 'A'.charCodeAt() + 1;
  });

  row = parseInt(row)
  return { row, column };
};

// Global Structures -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
// Map each category / label to a cell on the Monthly Budget sheet
var cellMapMonth = new Map()
cellMapMonth.set("categoriesList", {cell: fromA1Notation("K22")})
cellMapMonth.set("recurringExpenses", {cell: fromA1Notation("F5"), numEntries: 11, amountCol: 0, categoryCol: 2})
cellMapMonth.set("itemizedExpenses", {cell: fromA1Notation("F23"), numEntries: 31, amountCol: 0, categoryCol: 2})
cellMapMonth.set("tripExpenses", {cell: fromA1Notation("B60"), numEntries: 12, amountCol: 1, categoryCol: 0})
cellMapMonth.set("summedExpenses", {cell: fromA1Notation("L22")})
cellMapMonth.set("totalLeisureSpending", {cell: fromA1Notation("C32")})
cellMapMonth.set("totalSpendingGoal", {cell: fromA1Notation("B23")})
cellMapMonth.set("essentialsPercent", {cell: fromA1Notation("D30")})
cellMapMonth.set("savingsPercent", {cell: fromA1Notation("D31")})
cellMapMonth.set("wantsPercent", {cell: fromA1Notation("D32")})
cellMapMonth.set("netWorth", {cell: fromA1Notation("C87")})

// Map each category / label to a cell on the Yearly Spending
var cellMapYear = new Map()
cellMapYear.set("monthlyGoalStart", {cell: fromA1Notation("L4")})
