/*

CLASSROOM RUBRICMAKER v1.3
This will create a rubric for Google Classroom assignments based on the learning targets selected.
Created by Jonathan Kung

*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Rubric Tools")
    .addItem("Create Rubric", "copyCheckedSections")
    .addItem("Clear Checkboxes", "clearCheckedCheckboxes")
    .addItem("Extract Learning Targets", "copyRubricToLearningTargets")
    .addToUi();
}

function clearCheckedCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getActiveSheet();
  var lastRow = sourceSheet.getLastRow(); // Get last row with data

  for (var i = 3; i <= lastRow; i++) { // Start from row 3
    var checkboxCell = sourceSheet.getRange(i, 1); // Check Column A

    if (checkboxCell.getValue() === true) {
      checkboxCell.setValue(false); // Uncheck checkbox
    }
  }

  SpreadsheetApp.getUi().alert("Checked checkboxes have been cleared!");
}

function copyCheckedSections() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Select Learning Targets");
  var rubricSheet = ss.getSheetByName("Rubric");
  var lastSourceRow = sourceSheet.getLastRow(); // Get last row with data
  var destinationRow = 3; // Ensure we start writing at A3 in "Rubric"

  if (!rubricSheet) {
    rubricSheet = ss.insertSheet("Rubric"); // Create the sheet if it doesn't exist
  }

  // Clear only rows 3 and below to preserve rows 1 and 2
  var lastRow = rubricSheet.getLastRow();
  var atLeastOneChecked = false; // Flag to check if any checkbox is selected

  // Check if at least one checkbox is selected
  for (var i = 3; i <= lastSourceRow; i++) {
    if (sourceSheet.getRange(i, 1).getValue() === true) {
      atLeastOneChecked = true;
      break; // Exit loop early if we find a checked box
    }
  }

  // If none of the checkboxes are checked, exit
  if (!atLeastOneChecked) {
    SpreadsheetApp.getUi().alert("No learning targets selected!\nPlease check at least one checkbox before creating the rubric.");
    return; // Exit function early
  }


  if (lastRow > 2) {
    rubricSheet.getRange(3, 1, lastRow - 2, rubricSheet.getMaxColumns()).clearContent();
  }

  for (var i = 3; i <= lastSourceRow; i++) { // Start checking from row 3
    var checkbox = sourceSheet.getRange(i, 1).getValue(); // Check Column A

    if (checkbox === true) {
      var rangeToCopy = sourceSheet.getRange(i, 2, 4, 5); // Copy B3:F6 (4 rows, 5 columns)
      var valuesToCopy = rangeToCopy.getValues(); // Get values

      rubricSheet.getRange(destinationRow, 1, valuesToCopy.length, valuesToCopy[0].length).setValues(valuesToCopy);
      destinationRow += valuesToCopy.length; // Move down in "Rubric" without gaps
    }
  }

  clearCheckedCheckboxes();

  SpreadsheetApp.getUi().alert("Rubric created successfully!\nImport this Sheet using the Google Classroom Assignment Rubric section.\n\nBy Jonathan Kung");
}


function copyRubricToLearningTargets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Paste Rubric");
  var destinationSheet = ss.getSheetByName("Select Learning Targets");
  
  // Retrieve values from B1, C1, D1, E1
  var b1 = sourceSheet.getRange(1, 2).getValue(); // Column 2 = B
  var c1 = sourceSheet.getRange(1, 3).getValue(); // Column 3 = C
  var d1 = sourceSheet.getRange(1, 4).getValue(); // Column 4 = D
  var e1 = sourceSheet.getRange(1, 5).getValue(); // Column 5 = E

  // Get the number of rows in the source sheet (assuming data starts at row 2)
  var lastRow = sourceSheet.getLastRow();

  // Confirm with the user that they want to clear the destination sheet
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "Confirm Action", 
    "Are you sure you want to completely clear 'Select Learning Targets' and copy new data?", 
    ui.ButtonSet.YES_NO
  );

  if (response != ui.Button.YES) {
    ui.alert("Action canceled. No changes made.");
    return; // Exit the function if the user clicks "No"
  }

  // Remove all checkboxes from Column A
  var lastDestRow = destinationSheet.getLastRow();
  if (lastDestRow >= 3) {
    var checkboxRange = destinationSheet.getRange(3, 1, lastDestRow - 2);
    checkboxRange.clearDataValidations(); // Remove checkboxes
    checkboxRange.clearContent(); // Clear any remaining values
  }

  // Completely clear the "Select Learning Targets" sheet (including formatting, content, and checkboxes)
  destinationSheet.clear(); // This clears everything, including rows, content, and formatting

  // Set starting row for the destination sheet
  var destinationRow = 4;

  // Loop through the rows in the source sheet
  for (var i = 2; i <= lastRow; i++) {
    // Skip over every set of rows to create the empty rows in between
    if (i > 2) {
      destinationRow += 4; // Skip 4 rows (1 empty row + 3 for the current data block)
    }

    // Leave row 3, 7, 11, etc. completely empty except for the checkbox in column A
    destinationSheet.getRange(destinationRow, 1).clearContent();
    destinationSheet.getRange(destinationRow, 2, 1, destinationSheet.getMaxColumns() - 1).clearContent();

    // Step 2: Copy A (source) to B (destination)
    destinationSheet.getRange(destinationRow, 2).setValue(sourceSheet.getRange(i, 1).getValue()); // A -> B

    // Step 3: Write headers in the destination
    destinationSheet.getRange(destinationRow + 1, 3).setValue(e1);
    destinationSheet.getRange(destinationRow + 1, 4).setValue(d1);
    destinationSheet.getRange(destinationRow + 1, 5).setValue(c1);
    destinationSheet.getRange(destinationRow + 1, 6).setValue(b1);

    // Step 4: Copy E (source) to C (destination)
    destinationSheet.getRange(destinationRow + 2, 3).setValue(sourceSheet.getRange(i, 5).getValue()); // E -> C

    // Step 5: Copy D (source) to D (destination)
    destinationSheet.getRange(destinationRow + 2, 4).setValue(sourceSheet.getRange(i, 4).getValue()); // D -> D

    // Step 6: Copy C (source) to E (destination)
    destinationSheet.getRange(destinationRow + 2, 5).setValue(sourceSheet.getRange(i, 3).getValue()); // C -> E

    // Step 7: Copy B (source) to F (destination)
    destinationSheet.getRange(destinationRow + 2, 6).setValue(sourceSheet.getRange(i, 2).getValue()); // B -> F
  }

  // Add checkboxes in column A (A3, A7, A11, etc.)
  var checkboxRow = 3; // Start with A3
  for (var i = 2; i <= lastRow; i++) {
    if (checkboxRow <= destinationSheet.getLastRow()) {
      destinationSheet.getRange(checkboxRow, 1).insertCheckboxes(); // Insert checkbox at A3, A7, A11, etc.
      checkboxRow += 4; // Move to next row where checkbox should be inserted (A7, A11, etc.)
    }
  }

  // Notify the user the process is complete
  ui.alert("Rubric copied and checkboxes added to 'Select Learning Targets' successfully!");
}
