/**
 * The main script that will run when the checkbox is checked
 */
function script1() {
    Logger.log("Script started");

    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Clear the RangeToClean named range
    try {
        var rangeToClean1 = ss.getRangeByName("RangeToClean1");
        var rangeToClean2 = ss.getRangeByName("RangeToClean2");
        if (rangeToClean1 && rangeToClean2) {
            rangeToClean1.clearContent();
            rangeToClean2.clearContent();
            Logger.log("RangeToClean1 and RangeToClean2 have been cleared");
        } else {
            Logger.log("RangeToClean1 or RangeToClean2 named range not found");
        }
    } catch (error) {
        Logger.log("Error clearing RangeToClean1 or RangeToClean2: " + error);
    }
}

function script2() {
    Logger.log("Script 2 started");
    
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get the cell to increment
    try {
        var cellToIncrement = ss.getRangeByName("cellToIncrement");
        if (cellToIncrement) {
            // Get current value
            var currentValue = cellToIncrement.getValue();
            
            // Check if the current value is a number
            if (typeof currentValue === 'number') {
                // Increment by 1
                cellToIncrement.setValue(currentValue + 1);
                Logger.log("Cell value incremented from " + currentValue + " to " + (currentValue + 1));
            } else {
                // If not a number, set to 1
                cellToIncrement.setValue(1);
                Logger.log("Cell was not a number. Set to 1");
            }
        } else {
            Logger.log("cellToIncrement named range not found");
        }
    } catch (error) {
        Logger.log("Error incrementing cell: " + error);
    }
}

function script3() {
    Logger.log("Script 3 started");
    
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    try {
        // Get the cell to copy
        var cellToCopy = ss.getRangeByName("cellToCopy");
        if (!cellToCopy) {
            Logger.log("cellToCopy named range not found");
            return;
        }
        
        // Get the value to copy
        var valueToCopy = cellToCopy.getValue();
        
        // Get the sheet and column of the cell to copy
        var sheet = cellToCopy.getSheet();
        var column = cellToCopy.getColumn();
        
        // Find the first empty cell in the column
        var columnValues = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();
        var firstEmptyRow = -1;
        
        for (var i = 0; i < columnValues.length; i++) {
            if (columnValues[i][0] === "") {
                firstEmptyRow = i + 1; // +1 because array is 0-indexed but rows are 1-indexed
                break;
            }
        }
        
        // If no empty cell found, use the next row after the last row
        if (firstEmptyRow === -1) {
            firstEmptyRow = sheet.getLastRow() + 1;
        }
        
        // Paste the value to the first empty cell
        sheet.getRange(firstEmptyRow, column).setValue(valueToCopy);
        Logger.log("Value '" + valueToCopy + "' copied to row " + firstEmptyRow + ", column " + column);
        
    } catch (error) {
        Logger.log("Error in script3: " + error);
    }
}

/**
 * This script will check "rangeToClean2", and if a cell shows "Terminé", and the cell on the same row in the column "ErrorWithLine" is not empty, it will clear the cell in "rangeToClean2"
 */
function script4() {
    Logger.log("Script 4 started");
    
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    try {
        // Get the rangeToClean2
        var rangeToClean2 = ss.getRangeByName("rangeToClean2");
        if (!rangeToClean2) {
            Logger.log("rangeToClean2 named range not found");
            return;
        }
        
        // Get the sheet and range information
        var sheet = rangeToClean2.getSheet();
        var startRow = rangeToClean2.getRow();
        var startCol = rangeToClean2.getColumn();
        var numRows = rangeToClean2.getNumRows();
        var numCols = rangeToClean2.getNumColumns();
        
        // Get the ErrorWithLine column
        var errorWithLineRange = ss.getRangeByName("ErrorWithLine");
        if (!errorWithLineRange) {
            Logger.log("ErrorWithLine named range not found");
            return;
        }
        var errorCol = errorWithLineRange.getColumn();
        
        // Get all values from rangeToClean2
        var values = rangeToClean2.getValues();
        var modified = false;
        
        // Check each cell in rangeToClean2
        for (var i = 0; i < numRows; i++) {
            for (var j = 0; j < numCols; j++) {
                // Check if the cell contains "Terminé"
                if (values[i][j] === "Terminé") {
                    // Get the corresponding row in the sheet
                    var currentRow = startRow + i;
                    
                    // Check if the ErrorWithLine cell in the same row is not empty
                    var errorCellValue = sheet.getRange(currentRow, errorCol).getValue();
                    if (errorCellValue !== "") {
                        // Clear the cell in rangeToClean2
                        sheet.getRange(currentRow, startCol + j).clearContent();
                        Logger.log("Cleared cell at row " + currentRow + ", column " + (startCol + j) + " because ErrorWithLine is not empty");
                        modified = true;
                    }
                }
            }
        }
        
        if (!modified) {
            Logger.log("No cells were cleared in rangeToClean2");
        }
        
    } catch (error) {
        Logger.log("Error in script4: " + error);
    }
}
