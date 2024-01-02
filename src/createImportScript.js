// Function to generate import range formulas
function generateImportRangeFormulas() {

  // Get the active spreadsheet and the active sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = spreadsheet.getActiveSheet();

  // Get the URL of the spreadsheet
  var url = spreadsheet.getUrl();
  
  // Get the name of the active sheet
  var sheet = outputSheet.getName();

  // Find the range of the "Detailed Estimation Split-up" table
  var estimationTable = findRange("Detailed Estimation Split-up");
  var countStart = estimationTable.start;
  var count = estimationTable.end;

  // Find the range of the "Requirement" table
  var requirementTable = findRange("Requirement");
  var requirementStart = requirementTable.start;
  var requirementEnd = requirementStart + 1;
  
  // Generate the import range formula for the "Requirement" table
  var requirements = '=IMPORTRANGE("' + url + '", "' + sheet + '!A' + requirementStart + ':D' + requirementEnd + '")';

  // Generate the import range formulas for the main table and the client copy table
  var mainTable = '={' +
                  'IMPORTRANGE("' + url + '", "' + sheet + '!A' + countStart + ':E' + count + '"),' +
                  'IMPORTRANGE("' + url + '", "' + sheet + '!G' + countStart + ':G' + count + '")' +
                  '}';
  var clientTable = '=IMPORTRANGE("' + url + '", "' + sheet + '!A' + countStart + ':E' + count + '")';

  // Find the range of the "Efforts" table and generate the import range formula
  var effortTable = findRange("Efforts");
  var effortStart = effortTable.start;
  var effortEnd = effortTable.end;
  var efforts = '=IMPORTRANGE("' + url + '", "' + sheet + '!A' + effortStart + ':D' + effortEnd + '")';

  // Find the range of the "Project Cost" table and generate the import range formula
  var costTable = findRange("Project Cost");
  var costStart = costTable.start;
  var costEnd = costTable.end;
  var cost = '=IMPORTRANGE("' + url + '", "' + sheet + '!A' + costStart + ':D' + costEnd + '")';

  // Find the range of the "Development Estimation Split-up" table and generate the import range formula
  var devEstimationTable = findRange("Development Estimation Split-up");
  var estimationStart = devEstimationTable.start;
  var estimationEnd = devEstimationTable.end;
  var estimation = '=IMPORTRANGE("' + url + '", "' + sheet + '!A' + estimationStart + ':D' + estimationEnd + '")';

  // Find the range of the "Timeline" table and generate the import range formula
  var timelineTable = findRange("Timeline");
  var timelineStart = timelineTable.start;
  var timelineStartDate = timelineStart+3;
  var timelineEnd = timelineTable.end;
  var timeline = '={' +
    'IMPORTRANGE("' + url + '", "'+ sheet +'!A' + timelineStart+':D'+timelineStart+'");' +
    'IMPORTRANGE("' + url+ '", "'+ sheet +'!A'+timelineStartDate+':D'+timelineEnd+'")' +
  '}';

  // Define the starting row and column for the output data
  var startRow = count + 3;
  var startCol = 1;

  // Clear the rows in the output sheet starting from the start row
  clearRows(outputSheet, startRow);

  // Set the value of the first cell in the output sheet and apply a border
  outputSheet.getRange(startRow, 1).setValue("Next Steps\n\n1. Make a copy of this TAB into a blank spreadsheet\n2. Remove the unnecessary columns and rows\n3. Clear all the values but retain the formula\n4. Copy these formulas to the corresponding table").setBorder(true, true, true, true, true, true);
  
  // Merge the cells in the first row
  outputSheet.getRange(startRow , startCol , 1, 7).mergeAcross();

  // Apply the function to the cell that contains the text
  makeTextProminent(outputSheet.getRange(startRow, 1));
  
  // Increment the start row
  startRow++;

  // Define the data to be added to the sheet
  var data = [
    ["Requirement Table", "'"+requirements.toString()],
    ["Efforts Table", "'"+ efforts.toString()],
    ["Cost Table", "'"+ cost.toString()],
    ["Estimation Table", "'"+ estimation.toString()],
    ["Timeline Table", "'"+ timeline.toString()]
    ["Main Table", "'"+ mainTable.toString()],
    ["Client Table", "'"+ clientTable.toString()],
  ];
  
  // Get the range to set the values in
  var range = outputSheet.getRange(startRow, startCol, data.length, data[0].length);

  // Set the wrap strategy to WRAP
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Merge the cells in the second column of each row
  for (var i = 0; i < data.length; i++) {
    outputSheet.getRange(startRow + i, startCol + 1, 1, 6).mergeAcross();
  }

  // Set the values in the range and apply a border
  range.setValues(data).setBorder(true, true, true, true, true, true);
}
