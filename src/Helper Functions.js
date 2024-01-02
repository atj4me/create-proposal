// Function to sum the values of merged cells in a specified column if they match a specified value
function sumMergedCells(sheet, column, value) {
  
  // Initialize the sum
  var sum = 0;
  
  // Get the data from the sheet
  var data = sheet.getDataRange().getValues();
  
  // Initialize the merged value
  var mergedValue = "";
  
  // Loop through each row in the data
  for (var i = 0; i < data.length; i++) {
    
    // If the value in the specified column is not empty, update the merged value
    if (data[i][column] != "") {
      mergedValue = data[i][column];
    }

    // If the merged value matches the specified value and the value in column G is greater than 0, add it to the sum
    if (mergedValue == value && data[i][6] > 0 ) {
      sum += data[i][6]; // Change this line to specify column G (0-indexed)
    }
  }
  
  // Return the sum
  return sum;
}

// Function to calculate the sum of the values in a specified column that match a specified text
function calculateSum(text) {
  
  // Get the "Estimation" sheet from the active spreadsheet
  var sheet = SpreadsheetApp.getActive().getSheetByName("Estimation");
  
  // Calculate the sum of the values in column B that match the specified text
  var result = sumMergedCells(sheet, 1, text);
  
  // Return the result
  return result;
}

// Function to test the calculateSum function
function test() {
  
  // Log the result of the calculateSum function for the text "Development"
  console.log(calculateSum("Development"));
}

// Function to find the range of rows that contain a specified text
function findRange(text) {
  
  // Get the "Estimation" sheet from the active spreadsheet
  var sheet = SpreadsheetApp.getActive().getSheetByName("Estimation");
  
  // Get the range of column A
  var range = sheet.getRange("A:A");
  
  // Find the first row that contains the specified text
  var startRow = range.createTextFinder(text).matchEntireCell(true).findNext().getRow();
  
  // Find the first row that contains the text "END" after the start row
  var endRow = range.offset(startRow - 1, 0).createTextFinder("END").matchEntireCell(true).findNext().getRow();
  
  // Return the start and end rows
  return {
    'start' : startRow,
     'end' : endRow -1 
  }
}

// Function to get the start row of a range that contains a specified text
function getStartRow(text) {
  
  // Find the range that contains the specified text
  var result = findRange(text);
  
  // Return the start row of the range
  return result.start;
}

// Function to get the end row of a range that contains a specified text
function getEndRow(text) {
  
  // Find the range that contains the specified text
  var result = findRange(text);
  
  // Return the end row of the range
  return result.end;
}
