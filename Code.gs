function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

// ============ SERVER-SIDE FUNCTIONS ============

function addExpense(data) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Expenses");
    
    if (!sheet) {
      throw new Error("Expenses sheet not found");
    }
    
    // Validate data
    if (!data.date || !data.description || !data.notes || !data.amount) {
      throw new Error("All fields are required");
    }
    
    if (isNaN(data.amount) || data.amount <= 0) {
      throw new Error("Amount must be a positive number");
    }
    
    // Append row
    sheet.appendRow([
      data.date,
      data.description,
      data.notes,
      data.amount
    ]);
    
    return { success: true, message: "Expense added successfully! ✅" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function addCredit(data) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Credit");
    
    if (!sheet) {
      throw new Error("Credit sheet not found");
    }
    
    // Validate data
    if (!data.date || !data.description || !data.notes || !data.amount) {
      throw new Error("All fields are required");
    }
    
    if (isNaN(data.amount) || data.amount <= 0) {
      throw new Error("Amount must be a positive number");
    }
    
    // Append row
    sheet.appendRow([
      data.date,
      data.description,
      data.notes,
      data.amount
    ]);
    
    return { success: true, message: "Credit added successfully! ✅" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getDashboardData() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var expensesSheet = spreadsheet.getSheetByName("Expenses");
    var creditSheet = spreadsheet.getSheetByName("Credit");
    
    // Get all transactions
    var expenses = getSheetData(expensesSheet);
    var credits = getSheetData(creditSheet);
    
    // Calculate totals
    var totalExpenses = 0;
    for (var i = 0; i < expenses.length; i++) {
      totalExpenses += expenses[i].amount;
    }
    
    var totalCredits = 0;
    for (var i = 0; i < credits.length; i++) {
      totalCredits += credits[i].amount;
    }
    
    var balance = totalCredits - totalExpenses;
    var transactionCount = expenses.length + credits.length;
    
    // Get recent transactions (last 10)
    var allTransactions = [];
    for (var i = 0; i < expenses.length; i++) {
      var t = expenses[i];
      allTransactions.push({
        date: t.date,
        description: t.description,
        notes: t.notes,
        amount: t.amount,
        type: 'Expense'
      });
    }
    for (var i = 0; i < credits.length; i++) {
      var t = credits[i];
      allTransactions.push({
        date: t.date,
        description: t.description,
        notes: t.notes,
        amount: t.amount,
        type: 'Credit'
      });
    }
    
    // Sort by date (newest first)
    allTransactions.sort(function(a, b) {
      return new Date(b.date) - new Date(a.date);
    });
    
    var recentTransactions = allTransactions.slice(0, 10);
    
    return {
      totalExpenses: totalExpenses,
      totalCredits: totalCredits,
      balance: balance,
      transactionCount: transactionCount,
      recentTransactions: recentTransactions
    };
  } catch (error) {
    Logger.log('Error in getDashboardData: ' + error.toString());
    return {
      totalExpenses: 0,
      totalCredits: 0,
      balance: 0,
      transactionCount: 0,
      recentTransactions: []
    };
  }
}

function getHistory(filter) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var expensesSheet = spreadsheet.getSheetByName("Expenses");
    var creditSheet = spreadsheet.getSheetByName("Credit");
    
    var allTransactions = [];
    
    if (filter === 'all' || filter === 'expense') {
      var expenses = getSheetData(expensesSheet);
      for (var i = 0; i < expenses.length; i++) {
        var t = expenses[i];
        allTransactions.push({
          date: t.date,
          description: t.description,
          notes: t.notes,
          amount: t.amount,
          type: 'Expense'
        });
      }
    }
    
    if (filter === 'all' || filter === 'credit') {
      var credits = getSheetData(creditSheet);
      for (var i = 0; i < credits.length; i++) {
        var t = credits[i];
        allTransactions.push({
          date: t.date,
          description: t.description,
          notes: t.notes,
          amount: t.amount,
          type: 'Credit'
        });
      }
    }
    
    // Sort by date (newest first)
    allTransactions.sort(function(a, b) {
      return new Date(b.date) - new Date(a.date);
    });
    
    return allTransactions;
  } catch (error) {
    Logger.log('Error in getHistory: ' + error.toString());
    return [];
  }
}

function getReports() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var expensesSheet = spreadsheet.getSheetByName("Expenses");
    var creditSheet = spreadsheet.getSheetByName("Credit");
    
    var expenses = getSheetData(expensesSheet);
    var credits = getSheetData(creditSheet);
    
    var totalExpenses = 0;
    for (var i = 0; i < expenses.length; i++) {
      totalExpenses += expenses[i].amount;
    }
    
    var totalCredits = 0;
    for (var i = 0; i < credits.length; i++) {
      totalCredits += credits[i].amount;
    }
    
    var netBalance = totalCredits - totalExpenses;
    var transactionCount = expenses.length + credits.length;
    
    // Monthly summary
    var monthlyData = {};
    
    for (var i = 0; i < expenses.length; i++) {
      var t = expenses[i];
      var month = t.date.substring(0, 7); // YYYY-MM
      if (!monthlyData[month]) {
        monthlyData[month] = { expenses: 0, credits: 0 };
      }
      monthlyData[month].expenses += t.amount;
    }
    
    for (var i = 0; i < credits.length; i++) {
      var t = credits[i];
      var month = t.date.substring(0, 7);
      if (!monthlyData[month]) {
        monthlyData[month] = { expenses: 0, credits: 0 };
      }
      monthlyData[month].credits += t.amount;
    }
    
    var months = Object.keys(monthlyData);
    months.sort();
    
    var monthlySummary = [];
    for (var i = 0; i < months.length; i++) {
      var month = months[i];
      monthlySummary.push({
        month: month,
        expenses: monthlyData[month].expenses,
        credits: monthlyData[month].credits,
        net: monthlyData[month].credits - monthlyData[month].expenses
      });
    }
    
    return {
      totalExpenses: totalExpenses,
      totalCredits: totalCredits,
      netBalance: netBalance,
      transactionCount: transactionCount,
      monthlySummary: monthlySummary
    };
  } catch (error) {
    Logger.log('Error in getReports: ' + error.toString());
    return {
      totalExpenses: 0,
      totalCredits: 0,
      netBalance: 0,
      transactionCount: 0,
      monthlySummary: []
    };
  }
}

function getReport(fromDate, toDate, type) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var expensesSheet = spreadsheet.getSheetByName("Expenses");
    var creditSheet = spreadsheet.getSheetByName("Credit");
    
    var allTransactions = [];
    
    // Convert dates to comparable format (remove time)
    var fromDateObj = new Date(fromDate);
    fromDateObj.setHours(0, 0, 0, 0);
    var fromDateStr = fromDateObj.toISOString().split('T')[0];
    
    var toDateObj = new Date(toDate);
    toDateObj.setHours(23, 59, 59, 999);
    var toDateStr = toDateObj.toISOString().split('T')[0];
    
    // Get expenses if needed
    if (type === 'all' || type === 'expense') {
      var expenses = getSheetData(expensesSheet);
      for (var i = 0; i < expenses.length; i++) {
        var t = expenses[i];
        // Check if within date range
        if (t.date >= fromDateStr && t.date <= toDateStr) {
          allTransactions.push({
            date: t.date,
            description: t.description,
            notes: t.notes,
            amount: t.amount,
            type: 'Expense'
          });
        }
      }
    }
    
    // Get credits if needed
    if (type === 'all' || type === 'credit') {
      var credits = getSheetData(creditSheet);
      for (var i = 0; i < credits.length; i++) {
        var t = credits[i];
        // Check if within date range
        if (t.date >= fromDateStr && t.date <= toDateStr) {
          allTransactions.push({
            date: t.date,
            description: t.description,
            notes: t.notes,
            amount: t.amount,
            type: 'Credit'
          });
        }
      }
    }
    
    // Sort by date (newest first)
    allTransactions.sort(function(a, b) {
      return new Date(b.date) - new Date(a.date);
    });
    
    // Calculate total
    var total = 0;
    for (var i = 0; i < allTransactions.length; i++) {
      if (allTransactions[i].type === 'Expense') {
        total -= allTransactions[i].amount;
      } else {
        total += allTransactions[i].amount;
      }
    }
    
    return {
      transactions: allTransactions,
      total: total,
      count: allTransactions.length
    };
  } catch (error) {
    Logger.log('Error in getReport: ' + error.toString());
    return {
      transactions: [],
      total: 0,
      count: 0
    };
  }
}

// Helper function to get sheet data
function getSheetData(sheet) {
  if (!sheet) return [];
  
  try {
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    var result = [];
    // Assuming columns: Date, Description, Notes, Amount
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      // Check if row has data (at least date, description, and amount)
      if (row[0] && row[1] && row[3] !== undefined && row[3] !== '') {
        var amount = parseFloat(row[3]);
        if (!isNaN(amount)) {
          var dateStr = String(row[0]);
          // Try to format date properly
          try {
            var dateObj = new Date(row[0]);
            if (!isNaN(dateObj.getTime())) {
              dateStr = dateObj.toISOString().split('T')[0];
            }
          } catch(e) {
            // Keep original string
          }
          
          result.push({
            date: dateStr,
            description: String(row[1]),
            notes: String(row[2] || ''),
            amount: amount
          });
        }
      }
    }
    return result;
  } catch (error) {
    Logger.log('Error in getSheetData: ' + error.toString());
    return [];
  }
}
