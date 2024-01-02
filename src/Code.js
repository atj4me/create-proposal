// Function to extract merged data from a Google Spreadsheet
function extractMergedData() {
  
  // Get the "Estimation" sheet from the active spreadsheet
  var sheet = SpreadsheetApp.getActive().getSheetByName("Estimation");
  
  // Define the range of rows and columns to extract data from
  var startRow = getStartRow("Detailed Estimation Split-up")+2;
  var endRow = getEndRow("Detailed Estimation Split-up");
  var startColumn = 2;
  var endColumn = 5;
  
  // Get the range from the sheet
  var range = sheet.getRange(startRow, startColumn, endRow, endColumn - startColumn + 1);
  
  // Get the "Proposal" sheet from the active spreadsheet
  var outputSheet = SpreadsheetApp.getActive().getSheetByName("Proposal");
  
  // Define the starting row for the output
  var outputRow = 5;
  
  // Clear the rows in the output sheet
  clearRows(outputSheet,2) ;
  
  // Set the value of the first cell in the output sheet
  outputSheet.getRange(outputRow, 1).setValue("Next Steps\n\nCopy the content below and paste it on any AI tool requesting three things\n1. Project Overview in detail\n2. Project Service Description\n3. Detailed Explanation of the following in Layman's terms\n\nThis will generate a content for the proposal, which can be used in the Proposal Document, after CAREFUL REVIEW\n\nThanks,\nAjith Thampi Joseph");
  
  // Apply the function to the cell that contains the text
  makeTextProminent(outputSheet.getRange(outputRow, 1));
  
  // Increment the output row
  outputRow++;
  
  // Initialize variables to store the last heading, subheading, and paragraph
  var lastHeading, lastSubHeading, lastParagraph;
  
  // Loop through each row in the range
  for (var row = startRow; row <= sheet.getLastRow(); row++) {
    
    // If the value of the cell in the first column is "END", break the loop
    if (sheet.getRange(row, 1).getValue() === "END") {
      break;
    }
    
    // Loop through each column in the range
    for (var column = startColumn; column <= endColumn; column++) {
      
      // Get the cell at the current row and column
      var cell = sheet.getRange(row, column);
      
      // Initialize a variable to store the value of the cell
      var value = '';
      
      // Check if the cell is part of a merged range
      var mergedRanges = range.getMergedRanges();
      for (var i = 0; i < mergedRanges.length; i++) {
        if (cell.getRow() >= mergedRanges[i].getRow() && cell.getRow() <= mergedRanges[i].getLastRow() &&
            cell.getColumn() >= mergedRanges[i].getColumn() && cell.getColumn() <= mergedRanges[i].getLastColumn()) {
          value = mergedRanges[i].getValue();
          break;
        }
      }
      
      // If the cell is not part of a merged range, get its value directly
      if (!value) {
        value = cell.getValue();
      }
      
      // Write the value to the output sheet
      if (value) {
        if (column === startColumn) { // Column B
          if (value !== lastHeading) {
            outputSheet.getRange(outputRow, 1).setValue("");
            outputRow++;
            outputSheet.getRange(outputRow, 1).setValue("# " + value);
            outputRow++;
            lastHeading = value;
          }
        } else if (column === startColumn + 1) { // Column C
          if (value !== lastSubHeading) {
            outputSheet.getRange(outputRow, 1).setValue("");
            outputRow++;
            outputSheet.getRange(outputRow, 1).setValue("## " + value);
            outputRow++;
            lastSubHeading = value;
          }
        } else if (column === startColumn + 2) { // Column D
          if (value !== lastParagraph) {
            outputSheet.getRange(outputRow, 1).setValue("");
            outputRow++;
            outputSheet.getRange(outputRow, 1).setValue("### " + value);
            outputRow++;
            lastParagraph = value;
          }
        } else if (column === startColumn + 3) { // Column E
          outputSheet.getRange(outputRow, 1).setValue(value);
          outputRow++;
        }
      }
    }
  }
}

// Function to handle edits to the spreadsheet
function onEdit(e) {
  
  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the "Additional Effort" and "Efforts" sheets from the spreadsheet
  var additionalEffortSheet = ss.getSheetByName("Additional Effort");
  var effortsSheet = ss.getSheetByName("Efforts");
  
  // Get the range of the edit
  var range = e.range;
  
  // If the edit was made in the "Additional Effort" sheet and in cell A1
  if (range.getSheet().getName() == "Additional Effort" && range.getA1Notation() == "A1") {
    
    // Get the value of cell A1 in the "Additional Effort" sheet
    var a1Value = additionalEffortSheet.getRange("A1").getValue();
    
    // Get the values of the range A1:N10 in the "Efforts" sheet
    var effortsValues = effortsSheet.getRange("A1:N10").getValues();
    
    // Loop through each column in the range
    for (var i = 0; i < effortsValues[0].length; i += 3) {
      
      // If the value of the cell in the first row and current column is equal to the value of cell A1 in the "Additional Effort" sheet
      if (effortsValues[0][i] == a1Value) {
        
        // Initialize an array to store the new values
        var newValues = [];
        
        // Loop through each row in the range
        for (var j = 2; j < effortsValues.length; j++) {
          
          // Add the values of the cells in the current row and column, and the next column, to the array
          newValues.push([effortsValues[j][i], effortsValues[j][i + 1]]);
        }
        
        // Set the values of the range A3:B10 in the "Additional Effort" sheet to the new values
        additionalEffortSheet.getRange("A3:B10").setValues(newValues);
        
        // Break the loop
        break;
      }
    }
  }
}

// Function to clear the content of rows in a sheet
function clearRows(sheet, startRow) {
  
  // Calculate the number of rows to clear
  var numRows = sheet.getLastRow() - startRow + 1;
  
  // If the number of rows is less than 0, set it to 10
  if (numRows < 0) {
    numRows = 10;
  }
  
  // Get the number of columns in the sheet
  var numColumns = sheet.getLastColumn();
  
  // Get the range of cells to clear
  var range = sheet.getRange(startRow, 1, numRows, numColumns);
  
  // Clear the content of the cells in the range
  range.clearContent();
}

// Function to make the text in a cell prominent
function makeTextProminent(cell) {
  
  // Set the font size of the cell to 16
  cell.setFontSize(16);
  
  // Set the background color of the cell to cyan
  cell.setBackground('#00FFFF');
  
  // Set the font color of the cell to black
  cell.setFontColor('#000000');
  
  // Set the font style of the cell to italic
  cell.setFontStyle('italic');
}
