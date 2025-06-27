/**
 * Function that runs when the spreadsheet is opened
 * Creates a custom menu in the spreadsheet
 */
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    
    // Create the main Script Actions menu
    ui.createMenu('Script Actions')
        .addItem('Set Current Cell as Trigger 1', 'setCurrentCellAsTrigger')
        .addItem('Set Current Cell as Trigger 2', 'setCurrentCellAsTrigger2')
        .addItem('Set Current Cell as Trigger 3', 'setCurrentCellAsTrigger3')
        .addSeparator()
        .addItem('Run Script 1 Manually', 'script1')
        .addItem('Run Script 2 Manually', 'script2')
        .addItem('Run Script 3 Manually', 'script3')
        .addSeparator()
        .addItem('Check Setup', 'checkSetup')
        .addItem('Show Tutorial', 'showTutorial')
        .addToUi();
}

/**
 * Automatically triggered when a user edits the spreadsheet
 * @param {Object} e The event object
 */
function onEdit(e) {
    // Get the active sheet and the edited range
    var range = e.range;
    var ss = e.source;

    // Use named ranges for trigger checkboxes
    var triggerRange1;
    var triggerRange2;
    var triggerRange3;
    try {
        triggerRange1 = ss.getRangeByName("TriggerCheckbox");
        triggerRange2 = ss.getRangeByName("TriggerCheckbox2");
        triggerRange3 = ss.getRangeByName("TriggerCheckbox3");
        
        if (!triggerRange1 && !triggerRange2 && !triggerRange3) {
            // Only log the error, don't show message box in onEdit trigger (would cause issues)
            Logger.log("No trigger checkboxes found. Please set up trigger checkboxes.");
            return;
        }
    } catch (error) {
        Logger.log("Error finding named ranges: " + error.toString());
        return;
    }

    // Check if the edited cell is one of our trigger cells
    if (triggerRange1 && range.getA1Notation() === triggerRange1.getA1Notation()) {
        handleCheckboxChange(range, 1);
    } else if (triggerRange2 && range.getA1Notation() === triggerRange2.getA1Notation()) {
        handleCheckboxChange(range, 2);
    } else if (triggerRange3 && range.getA1Notation() === triggerRange3.getA1Notation()) {
        handleCheckboxChange(range, 3);
    }
}

/**
 * Helper function to handle checkbox changes
 */
function handleCheckboxChange(range, triggerNumber) {
    // Check if the checkbox is checked
    if (range.getValue() === true) {
        try {
            // Run your script
            if (triggerNumber === 1) {
                script1();
            } else if (triggerNumber === 2) {
                script2();
            } else if (triggerNumber === 3) {
                script3();
            }
        } catch (error) {
            Logger.log("Error: " + error.toString());
            // Don't use Browser.msgBox in onEdit trigger as it can cause issues
            Logger.log("Error occurred: " + error.toString());
        } finally {
            // Uncheck the checkbox when done
            range.setValue(false);
        }
    }
}

/**
 * Helper function to prompt user to select a range
 */
function selectRangeWithPrompt(title, prompt) {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var result = ui.prompt(
        title,
        prompt,
        ui.ButtonSet.OK_CANCEL
    );

    // Check if the user clicked "OK"
    if (result.getSelectedButton() === ui.Button.OK) {
        // Get the active range (what the user has selected)
        var selectedRange = ss.getActiveRange();
        if (!selectedRange) {
            ui.alert('Error', 'No cell was selected. Please try again.', ui.ButtonSet.OK);
            return null;
        }
        return selectedRange;
    } else {
        ui.alert('Setup Cancelled', 'Named range setup was cancelled.', ui.ButtonSet.OK);
        return null;
    }
}

/**
 * Function to manually set up the trigger if needed
 * Run this once to set up the trigger if the simple onEdit trigger isn't sufficient
 */
function createEditTrigger() {
    // Delete existing triggers
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === "onEdit") {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }

    // Create new trigger
    ScriptApp.newTrigger("onEdit")
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onEdit()
        .create();

    var ui = SpreadsheetApp.getUi();
    ui.alert("Edit trigger has been set up!");
}

/**
 * Function to check if the named ranges are properly set up
 * Run this function to verify your configuration
 */
function checkSetup() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    
    var triggerRange1 = ss.getRangeByName("TriggerCheckbox");
    var triggerRange2 = ss.getRangeByName("TriggerCheckbox2");
    var triggerRange3 = ss.getRangeByName("TriggerCheckbox3");
    var cellToIncrement = ss.getRangeByName("cellToIncrement");
    var cellToCopy = ss.getRangeByName("cellToCopy");
    
    var message = "";
    
    if (!triggerRange1) {
        message += "❌ Trigger checkbox 1 is not set up.\n" +
                   "   Please use the 'Script Actions > Set Current Cell as Trigger 1' menu to configure it.\n\n";
    } else {
        message += "✅ Trigger checkbox 1 is set to cell " + triggerRange1.getA1Notation() + "\n\n";
    }
    
    if (!triggerRange2) {
        message += "❌ Trigger checkbox 2 is not set up.\n" +
                   "   Please use the 'Script Actions > Set Current Cell as Trigger 2' menu to configure it.\n\n";
    } else {
        message += "✅ Trigger checkbox 2 is set to cell " + triggerRange2.getA1Notation() + "\n\n";
    }
    
    if (!triggerRange3) {
        message += "❌ Trigger checkbox 3 is not set up.\n" +
                   "   Please use the 'Script Actions > Set Current Cell as Trigger 3' menu to configure it.\n\n";
    } else {
        message += "✅ Trigger checkbox 3 is set to cell " + triggerRange3.getA1Notation() + "\n\n";
    }
    
    if (!cellToIncrement) {
        message += "❌ 'cellToIncrement' named range is not set up.\n" +
                   "   Please create this named range for script2 to work properly.\n\n";
    } else {
        message += "✅ 'cellToIncrement' named range is set to cell " + cellToIncrement.getA1Notation() + "\n\n";
    }
    
    if (!cellToCopy) {
        message += "❌ 'cellToCopy' named range is not set up.\n" +
                   "   Please create this named range for script3 to work properly.\n\n";
    } else {
        message += "✅ 'cellToCopy' named range is set to cell " + cellToCopy.getA1Notation() + "\n\n";
    }
    
    ui.alert("Setup Status", message, ui.ButtonSet.OK);
}

/**
 * Sets the currently selected cell as the trigger and adds a checkbox to it
 */
function setCurrentCellAsTrigger() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var currentCell = ss.getActiveRange();
    
    // Check if a cell is selected
    if (!currentCell || currentCell.getNumRows() > 1 || currentCell.getNumColumns() > 1) {
        ui.alert(
            'Selection Error',
            'Please select a single cell before running this command.',
            ui.ButtonSet.OK
        );
        return;
    }
    
    // Check if named range already exists
    var existingRange = ss.getRangeByName("TriggerCheckbox");
    if (existingRange) {
        var response = ui.alert(
            'Trigger Already Set',
            'A trigger is already set at cell ' + existingRange.getA1Notation() + 
            '. Do you want to change it to ' + currentCell.getA1Notation() + '?',
            ui.ButtonSet.YES_NO
        );
        
        if (response !== ui.Button.YES) {
            ui.alert('Action Cancelled', 'The trigger cell was not changed.', ui.ButtonSet.OK);
            return;
        }
        
        // Remove existing named range
        ss.removeNamedRange("TriggerCheckbox");
    }
    
    // Create named range for the trigger
    ss.setNamedRange("TriggerCheckbox", currentCell);
    
    // Insert checkbox in the cell
    try {
        // This is a workaround as there's no direct API to insert a checkbox
        // We'll set the data validation to CHECKBOX type
        var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
        currentCell.setDataValidation(rule);
        
        // Set initial value to false (unchecked)
        currentCell.setValue(false);
        
        // Success message
        ui.alert(
            'Setup Complete',
            'Cell ' + currentCell.getA1Notation() + ' is now set as your trigger!\n\n' +
            'When you check this box, your script will run and then automatically uncheck itself.',
            ui.ButtonSet.OK
        );
    } catch (error) {
        ui.alert(
            'Error Setting Up Checkbox',
            'The cell was set as a trigger, but there was an error adding the checkbox: ' + error.toString() + '\n\n' +
            'Please manually add a checkbox to cell ' + currentCell.getA1Notation() + ' by selecting the cell and using Insert > Checkbox.',
            ui.ButtonSet.OK
        );
    }
}

/**
 * Sets the currently selected cell as the second trigger and adds a checkbox to it
 */
function setCurrentCellAsTrigger2() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var currentCell = ss.getActiveRange();
    
    // Check if a cell is selected
    if (!currentCell || currentCell.getNumRows() > 1 || currentCell.getNumColumns() > 1) {
        ui.alert(
            'Selection Error',
            'Please select a single cell before running this command.',
            ui.ButtonSet.OK
        );
        return;
    }
    
    // Check if named range already exists
    var existingRange = ss.getRangeByName("TriggerCheckbox2");
    if (existingRange) {
        var response = ui.alert(
            'Trigger Already Set',
            'A trigger 2 is already set at cell ' + existingRange.getA1Notation() + 
            '. Do you want to change it to ' + currentCell.getA1Notation() + '?',
            ui.ButtonSet.YES_NO
        );
        
        if (response !== ui.Button.YES) {
            ui.alert('Action Cancelled', 'The trigger cell was not changed.', ui.ButtonSet.OK);
            return;
        }
        
        // Remove existing named range
        ss.removeNamedRange("TriggerCheckbox2");
    }
    
    // Create named range for the trigger
    ss.setNamedRange("TriggerCheckbox2", currentCell);
    
    // Insert checkbox in the cell
    try {
        // This is a workaround as there's no direct API to insert a checkbox
        // We'll set the data validation to CHECKBOX type
        var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
        currentCell.setDataValidation(rule);
        
        // Set initial value to false (unchecked)
        currentCell.setValue(false);
        
        // Success message
        ui.alert(
            'Setup Complete',
            'Cell ' + currentCell.getA1Notation() + ' is now set as your trigger 2!\n\n' +
            'When you check this box, script2 will run and then automatically uncheck itself.',
            ui.ButtonSet.OK
        );
    } catch (error) {
        ui.alert(
            'Error Setting Up Checkbox',
            'The cell was set as a trigger, but there was an error adding the checkbox: ' + error.toString() + '\n\n' +
            'Please manually add a checkbox to cell ' + currentCell.getA1Notation() + ' by selecting the cell and using Insert > Checkbox.',
            ui.ButtonSet.OK
        );
    }
}

/**
 * Sets the currently selected cell as the third trigger and adds a checkbox to it
 */
function setCurrentCellAsTrigger3() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var currentCell = ss.getActiveRange();
    
    // Check if a cell is selected
    if (!currentCell || currentCell.getNumRows() > 1 || currentCell.getNumColumns() > 1) {
        ui.alert(
            'Selection Error',
            'Please select a single cell before running this command.',
            ui.ButtonSet.OK
        );
        return;
    }
    
    // Check if named range already exists
    var existingRange = ss.getRangeByName("TriggerCheckbox3");
    if (existingRange) {
        var response = ui.alert(
            'Trigger Already Set',
            'A trigger 3 is already set at cell ' + existingRange.getA1Notation() + 
            '. Do you want to change it to ' + currentCell.getA1Notation() + '?',
            ui.ButtonSet.YES_NO
        );
        
        if (response !== ui.Button.YES) {
            ui.alert('Action Cancelled', 'The trigger cell was not changed.', ui.ButtonSet.OK);
            return;
        }
        
        // Remove existing named range
        ss.removeNamedRange("TriggerCheckbox3");
    }
    
    // Create named range for the trigger
    ss.setNamedRange("TriggerCheckbox3", currentCell);
    
    // Insert checkbox in the cell
    try {
        // This is a workaround as there's no direct API to insert a checkbox
        // We'll set the data validation to CHECKBOX type
        var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
        currentCell.setDataValidation(rule);
        
        // Set initial value to false (unchecked)
        currentCell.setValue(false);
        
        // Success message
        ui.alert(
            'Setup Complete',
            'Cell ' + currentCell.getA1Notation() + ' is now set as your trigger 3!\n\n' +
            'When you check this box, script3 will run and then automatically uncheck itself.',
            ui.ButtonSet.OK
        );
    } catch (error) {
        ui.alert(
            'Error Setting Up Checkbox',
            'The cell was set as a trigger, but there was an error adding the checkbox: ' + error.toString() + '\n\n' +
            'Please manually add a checkbox to cell ' + currentCell.getA1Notation() + ' by selecting the cell and using Insert > Checkbox.',
            ui.ButtonSet.OK
        );
    }
}

/**
 * Shows a tutorial dialog explaining how to use the script
 */
function showTutorial() {
    var ui = SpreadsheetApp.getUi();
    
    var htmlOutput = HtmlService
        .createHtmlOutput(
            '<style>' +
            '  body { font-family: Arial, sans-serif; }' +
            '  h3 { color: #555; }' +
            '  p { color: #666; }' +
            '  ul { color: #666; }' +
            '  ol { color: #666; }' +
            '  li { color: #666; }' +
            '  code { color: #333; background-color: #f5f5f5; padding: 2px 4px; border-radius: 3px; }' +
            '</style>' +
            '<body>' +
            '<p>This script allows you to run actions automatically by checking checkboxes in your spreadsheet.</p>' +
            '<h3>Quick Setup</h3>' +
            '<ol>' +
            '  <li>Select the cell where you want to place your first trigger checkbox</li>' +
            '  <li>Click on <b>Script Actions > Set Current Cell as Trigger 1</b></li>' +
            '  <li>A checkbox will be added to the selected cell</li>' +
            '  <li>When you check this box, script1 will run and then automatically uncheck itself</li>' +
            '  <li>Repeat the process with <b>Script Actions > Set Current Cell as Trigger 2</b> for the second trigger</li>' +
            '  <li>Create a named range called "cellToIncrement" for script2 to work properly</li>' +
            '  <li>Repeat the process with <b>Script Actions > Set Current Cell as Trigger 3</b> for the third trigger</li>' +
            '  <li>Create a named range called "cellToCopy" for script3 to work properly</li>' +
            '</ol>' +
            '<h3>How It Works</h3>' +
            '<ul>' +
            '  <li>The script watches for changes to the checkbox cells</li>' +
            '  <li>When a checkbox is checked, it runs the corresponding script</li>' +
            '  <li>After running, it automatically unchecks the box</li>' +
            '  <li>Trigger 1 runs script1 which clears the "RangeToClean" named range</li>' +
            '  <li>Trigger 2 runs script2 which increments the value in the "cellToIncrement" named range</li>' +
            '  <li>Trigger 3 runs script3 which copies the value from "cellToCopy" to the first empty cell in its column</li>' +
            '</ul>' +
            '<h3>Customizing the Script</h3>' +
            '<p>To change what happens when the checkboxes are checked, edit the <code>script1()</code>, <code>script2()</code>, and <code>script3()</code> functions in the script editor.</p>' +
            '</body>'
        )
        .setWidth(450)
        .setHeight(500);
    
    ui.showModalDialog(htmlOutput, 'Checkbox Trigger Tutorial');
}