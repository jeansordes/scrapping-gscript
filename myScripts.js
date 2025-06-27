/**
 * The main script that will run when the checkbox is checked
 */
function script1() {
    Logger.log("Script started");

    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Clear the RangeToClean named range
    try {
        var rangeToClean = ss.getRangeByName("RangeToClean");
        if (rangeToClean) {
            rangeToClean.clearContent();
            Logger.log("RangeToClean has been cleared");
        } else {
            Logger.log("RangeToClean named range not found");
        }
    } catch (error) {
        Logger.log("Error clearing RangeToClean: " + error);
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