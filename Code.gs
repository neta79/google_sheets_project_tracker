/** @OnlyCurrentDoc */
/**
 * Google Sheets Requirements and Cross-Cutting Concerns Tracker
 * This solution allows managing requirements, cross-cutting concerns, and automatically
 * generates a comprehensive checklist while preserving state across updates.
 */

// Constants for sheet names
const SHEETS = {
  REQUIREMENTS: 'Requirements',
  CONCERNS: 'Cross-Cutting Concerns',
  CHECKLIST: 'Main Checklist',
  STATE_STORAGE: 'StateStorage',
  CONFIG: 'Config'
};

// Constants for status values
const STATUS = {
  EMPTY: '',
  DESIGN: 'Design',
  DEVEL: 'Devel',
  TESTABLE: 'Testable',
  SHIPPED: 'Shipped',
  REGRESSION: 'Regression',
  REJECTED: 'Rejected',
};

// Constants for item types
const ITEM_TYPES = {
  EMPTY: '',
  EPIC: 'Epic',
  FEATURE: 'Feature',
  TECHNICAL_REQUIREMENT: 'Technical Requirement',
  USER_STORY: 'User Story',
  TASK: 'Task',
  BUG: 'Bug',
  CORE: 'Core',
  SPIKE: 'Spike'
};

const DISCARDABLE_PREFIXES = [
  "REQ-",
  "CONCERN-",
  "CONC-",
  "CON-",
  "CCC-",
];


const DEFAULT_PROJECT_TEAM = [
  "Awesome1",
  "CoolDude2",
  "StellarFella",
  "Me",
  "ThatWizardGirl",
];

/**
 * Creates the menu when the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Requirements Tracker')
      .addItem('Setup Sheets', 'setupSheets')
      .addItem('Concern Picker', 'showConcernPicker')
      .addItem('(Re)Generate Checklist', 'generateChecklist')
      .addItem('Update Completion Percentages', 'updateCompletionAfterStatusChange')
      .addToUi();
}

/**
 * Set up all sheets with appropriate headers and configurations
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Requirements sheet if it doesn't exist
  let reqSheet = ss.getSheetByName(SHEETS.REQUIREMENTS);
  if (!reqSheet) {
    reqSheet = ss.insertSheet(SHEETS.REQUIREMENTS);
    reqSheet.getRange('A1:H1').setValues([['ID', 'Requirement Description', 'Item Type', 'Priority', 'Sources', 'Assigned to', '%', 'Concerns']]);
    reqSheet.getRange('A1:H1').setFontWeight('bold');
    reqSheet.setFrozenRows(1);

    // Hide the Concerns column (column H, index 8)
    reqSheet.hideColumns(8);
    
    // Add data validation for the item type column
    const itemTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(Object.values(ITEM_TYPES), true)
      .build();
    reqSheet.getRange('C2:C1000').setDataValidation(itemTypeRule);
    
    // Add data validation for the assigned to column
    const projectTeam = getProjectTeam();
    const assignedToRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([''].concat(projectTeam), true)
      .build();
    reqSheet.getRange('F2:F1000').setDataValidation(assignedToRule);
  }
  
  // Create Cross-Cutting Concerns sheet if it doesn't exist
  let concernsSheet = ss.getSheetByName(SHEETS.CONCERNS);
  if (!concernsSheet) {
    concernsSheet = ss.insertSheet(SHEETS.CONCERNS);
    concernsSheet.getRange('A1:D1').setValues([['ID', 'Description', 'Priority', 'Sources']]);
    concernsSheet.getRange('A1:D1').setFontWeight('bold');
    concernsSheet.setFrozenRows(1);
  }
  
  // Create Main Checklist sheet if it doesn't exist
  let checklistSheet = ss.getSheetByName(SHEETS.CHECKLIST);
  if (!checklistSheet) {
    checklistSheet = ss.insertSheet(SHEETS.CHECKLIST);
    checklistSheet.getRange('A1:J1').setValues([['Item ID', 'Assigned to', 'Status', 'Test OK', 'Item Type', 'Requirement Description', 'CCC Description', 'Requirement ID', 'CCC ID', 'Parent Requirement']]);
    checklistSheet.getRange('A1:J1').setFontWeight('bold');
    checklistSheet.setFrozenRows(1);
  }
  
  // Create hidden State Storage sheet if it doesn't exist
  let stateSheet = ss.getSheetByName(SHEETS.STATE_STORAGE);
  if (!stateSheet) {
    stateSheet = ss.insertSheet(SHEETS.STATE_STORAGE);
    stateSheet.getRange('A1:E1').setValues([['Item ID', 'Assigned to', 'Status', 'Test OK', 'Last Updated']]);
    stateSheet.getRange('A1:E1').setFontWeight('bold');
    stateSheet.setFrozenRows(1);
    stateSheet.hideSheet();
  }

  // Create Config sheet if it doesn't exist
  let configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) {
    configSheet = ss.insertSheet(SHEETS.CONFIG);
    configSheet.getRange('A1:B1').setValues([['Setting', 'Value']]);
    configSheet.getRange('A2:B2').setValues([['Last Generated', '']]);
    // Add discardable prefixes to config
    configSheet.getRange('A3:B3').setValues([['Discardable Prefixes', JSON.stringify(DISCARDABLE_PREFIXES)]]);
    // Add project team to config
    configSheet.getRange('A4:B4').setValues([['Project Team', JSON.stringify(DEFAULT_PROJECT_TEAM)]]);
    configSheet.getRange('A1:B1').setFontWeight('bold');
    configSheet.setFrozenRows(1);
    configSheet.hideSheet();
  }
  
  // Check if Discardable Prefixes setting exists, if not add it
  const configData = configSheet.getDataRange().getValues();
  let prefixesExist = false;
  let projectTeamExists = false;
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0] === 'Discardable Prefixes') {
      prefixesExist = true;
    }
    if (configData[i][0] === 'Project Team') {
      projectTeamExists = true;
    }
  }
  if (!prefixesExist) {
    const newRow = configSheet.getLastRow() + 1;
    configSheet.getRange(newRow, 1, 1, 2).setValues([['Discardable Prefixes', JSON.stringify(DISCARDABLE_PREFIXES)]]);
  }
  if (!projectTeamExists) {
    const newRow = configSheet.getLastRow() + 1;
    configSheet.getRange(newRow, 1, 1, 2).setValues([['Project Team', JSON.stringify(DEFAULT_PROJECT_TEAM)]]);
  }
  
  // Create data validation for the status column in checklist
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(STATUS), true)
    .build();
  
  // Create data validation for the assigned to column
  const projectTeam = getProjectTeam();
  const assignedToRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([''].concat(projectTeam), true)
    .build();
  
  // Create checkbox for Test OK column
  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  
  const lastRow = Math.max(checklistSheet.getLastRow(), 2);
  checklistSheet.getRange(2, 2, 1000).setDataValidation(assignedToRule); // Assigned to is now column 2
  checklistSheet.getRange(2, 3, 1000).setDataValidation(statusRule);     // Status is now column 3
  checklistSheet.getRange(2, 4, 1000).setDataValidation(checkboxRule);   // Test OK is now column 4
  
  setUpTriggers();
  
  SpreadsheetApp.getUi().alert('Sheets have been set up successfully!');
}

/**
 * Set up script triggers for the spreadsheet
 */
function setUpTriggers() {
  // Delete existing triggers to avoid duplicates
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  
  // Create an onEdit trigger
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
    
  // Create an onOpen trigger
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onOpen()
    .create();
  
  // Create a custom menu item to manually update percentages if needed
  ScriptApp.newTrigger('updateCompletionAfterStatusChange')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
    
  console.log('Triggers have been set up successfully');
}

/**
 * Handle edits to the spreadsheet
 */
function onEdit(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    
    // If editing the checklist status, test OK, or assigned to column
    if (sheetName === SHEETS.CHECKLIST && 
        (e.range.getColumn() === 2 || e.range.getColumn() === 3 || e.range.getColumn() === 4)) {
      const itemId = sheet.getRange(e.range.getRow(), 1).getValue();
      
      if (itemId) {
        const assignedTo = sheet.getRange(e.range.getRow(), 2).getValue();
        const status = sheet.getRange(e.range.getRow(), 3).getValue();
        const testOk = sheet.getRange(e.range.getRow(), 4).getValue();
        
        console.log(`Change detected for Item ${itemId}: Assigned to=${assignedTo}, Status=${status}, TestOK=${testOk}`);
        
        saveState(itemId, assignedTo, status, testOk);
        
        // For simple onEdit trigger, we need to create a separate function and call it
        // This is because onEdit has limitations on what it can do
        updateCompletionAfterStatusChange();
      }
    }
  } catch (error) {
    console.error('Error in onEdit:', error);
  }
}

/**
 * Updates completion percentages after a status change
 * This needs to be a separate function so it can be called from onEdit
 * but run with full permissions
 */
function updateCompletionAfterStatusChange() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get required data
    const reqData = ss.getSheetByName(SHEETS.REQUIREMENTS).getDataRange().getValues();
    const checklistData = ss.getSheetByName(SHEETS.CHECKLIST).getDataRange().getValues();
    const checklistItems = [];
    
    // Convert checklist data to item objects (skip header)
    for (let i = 1; i < checklistData.length; i++) {
      if (checklistData[i][0]) {
        checklistItems.push({
          itemId: checklistData[i][0],
          reqId: checklistData[i][7],
          status: checklistData[i][2],
          testOk: checklistData[i][3] || false
        });
      }
    }
    
    console.log(`Updating percentages for ${checklistItems.length} checklist items`);
    
    // Update percentages
    calculateCompletionPercentages(reqData, checklistItems);
    
    // Force UI update
    SpreadsheetApp.flush();
    
    console.log('Completion percentages updated successfully');
  } catch (error) {
    console.error('Error updating completion percentages:', error);
  }
}

/**
 * Get project team members from config or use fallback
 */
function getProjectTeam() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  
  if (!configSheet) return DEFAULT_PROJECT_TEAM;
  
  const configData = configSheet.getDataRange().getValues();
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0] === 'Project Team') {
      try {
        return JSON.parse(configData[i][1]);
      } catch (e) {
        console.error('Failed to parse project team:', e);
        return DEFAULT_PROJECT_TEAM;
      }
    }
  }
  
  return DEFAULT_PROJECT_TEAM;
}

/**
 * Get discardable prefixes from config or use fallback
 */
function getDiscardablePrefixes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  
  if (!configSheet) return DISCARDABLE_PREFIXES;
  
  const configData = configSheet.getDataRange().getValues();
  for (let i = 1; i < configData.length; i++) {
    if (configData[i][0] === 'Discardable Prefixes') {
      try {
        return JSON.parse(configData[i][1]);
      } catch (e) {
        console.error('Failed to parse discardable prefixes:', e);
        return DISCARDABLE_PREFIXES;
      }
    }
  }
  
  return DISCARDABLE_PREFIXES;
}

/**
 * Strip discardable prefixes from ID
 */
function stripPrefixes(id) {
  const prefixes = getDiscardablePrefixes();
  let result = id;
  
  prefixes.forEach(prefix => {
    if (result.startsWith(prefix)) {
      result = result.substring(prefix.length);
    }
  });
  
  return result;
}

/**
 * Show concern picker dialog
 */
function showConcernPicker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  
  // Only show the picker if we're in the Requirements sheet
  if (activeSheet.getName() !== SHEETS.REQUIREMENTS) {
    SpreadsheetApp.getUi().alert('Please select cells in the Requirements sheet.');
    return;
  }
  
  // Get all selected ranges (handles multiple selections via ctrl+click)
  const selectedRanges = activeSheet.getActiveRangeList();
  if (!selectedRanges) {
    SpreadsheetApp.getUi().alert('Please select at least one requirement row.');
    return;
  }
  
  // Get the headers to find the Concerns column and ID column
  const headers = activeSheet.getRange(1, 1, 1, activeSheet.getLastColumn()).getValues()[0];
  const concernsColIndex = headers.indexOf('Concerns') + 1;
  const reqIdColIndex = headers.indexOf('ID') + 1;
  
  if (concernsColIndex < 1) {
    SpreadsheetApp.getUi().alert('The Requirements sheet does not have a Concerns column. Please run Setup Sheets first.');
    return;
  }
  
  // Process all selected ranges to collect requirement IDs and concerns
  const selectedRows = new Set(); // Use Set to avoid duplicates
  const selectedReqIds = [];
  const currentConcernsMap = {};
  const reqRows = {}; // Map of reqId to row number
  
  // Get all the ranges from the range list
  const ranges = selectedRanges.getRanges();
  
  for (let r = 0; r < ranges.length; r++) {
    const range = ranges[r];
    const firstRow = range.getRow();
    const numRows = range.getNumRows();
    
    // Process each row in the current range
    for (let i = 0; i < numRows; i++) {
      const currentRow = firstRow + i;
      // Skip header row
      if (currentRow === 1) continue;
      
      // Skip rows we've already processed
      if (selectedRows.has(currentRow)) continue;
      
      // Add this row to our processed set
      selectedRows.add(currentRow);
      
      const reqId = activeSheet.getRange(currentRow, reqIdColIndex).getValue();
      if (!reqId) continue; // Skip rows without a requirement ID
      
      const concernsCell = activeSheet.getRange(currentRow, concernsColIndex).getValue() || '';
      const concerns = getConcernIDS(concernsCell).map(item => item.trim()).filter(item => item);
      
      selectedReqIds.push(reqId);
      currentConcernsMap[reqId] = concerns;
      reqRows[reqId] = currentRow; // Map the reqId to its row number
    }
  }
  
  if (selectedReqIds.length === 0) {
    SpreadsheetApp.getUi().alert('No valid requirements found in the selected rows.');
    return;
  }
  
  // Get all available concerns
  const concernsSheet = ss.getSheetByName(SHEETS.CONCERNS);
  if (!concernsSheet) {
    SpreadsheetApp.getUi().alert('Cross-Cutting Concerns sheet not found.');
    return;
  }
  
  const concernsData = concernsSheet.getDataRange().getValues();
  const concerns = [];
  
  // Skip header row
  for (let i = 1; i < concernsData.length; i++) {
    if (concernsData[i][0]) {
      concerns.push({
        id: concernsData[i][0],
        description: concernsData[i][1]
      });
    }
  }
  
  // Create template
  const template = HtmlService.createTemplateFromFile('ConcernPicker');
  template.concerns = concerns;
  template.selectedReqIds = selectedReqIds;
  template.selectedRowCount = selectedReqIds.length;
  template.currentConcernsMap = currentConcernsMap;
  template.reqRows = reqRows; // Pass the reqRows mapping to the template
  
  // Determine the checkbox state for each concern
  const checkboxStates = {};
  concerns.forEach(concern => {
    // Count how many selected requirements have this concern
    let count = 0;
    selectedReqIds.forEach(reqId => {
      if (currentConcernsMap[reqId].includes(concern.id)) {
        count++;
      }
    });
    
    if (count === 0) {
      checkboxStates[concern.id] = 'unchecked';
    } else if (count === selectedReqIds.length) {
      checkboxStates[concern.id] = 'checked';
    } else {
      checkboxStates[concern.id] = 'indeterminate';
    }
  });
  
  template.checkboxStates = checkboxStates;
  
  const html = template.evaluate()
    .setWidth(1000)
    .setHeight(600)
    .setTitle('Select Cross-Cutting Concerns');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Cross-Cutting Concerns');
}

/**
 * Update concerns for multiple requirements
 */
function updateConcerns(reqRows, concernStates) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reqSheet = ss.getSheetByName(SHEETS.REQUIREMENTS);
    
    // Find the concerns column
    const headers = reqSheet.getRange(1, 1, 1, reqSheet.getLastColumn()).getValues()[0];
    const concernsColIndex = headers.indexOf('Concerns') + 1;
    
    if (concernsColIndex < 1) return false;
    
    // For each row, update the concerns based on the provided states
    Object.keys(reqRows).forEach(reqId => {
      const row = reqRows[reqId];
      
      // Get current concerns for this requirement
      let currentConcerns = reqSheet.getRange(row, concernsColIndex).getValue() || '';
      currentConcerns = getConcernIDS(currentConcerns).map(c => c.trim()).filter(c => c);
      
      // Process each concern based on its state
      Object.keys(concernStates).forEach(concernId => {
        const state = concernStates[concernId];
        
        if (state === 'checked') {
          // Add the concern if not already present
          if (!currentConcerns.includes(concernId)) {
            currentConcerns.push(concernId);
          }
        } else if (state === 'unchecked') {
          // Remove the concern if present
          const index = currentConcerns.indexOf(concernId);
          if (index !== -1) {
            currentConcerns.splice(index, 1);
          }
        }
        // For 'indeterminate' state, do nothing
      });
      
      // Update the cell
      const cellRange = reqSheet.getRange(row, concernsColIndex);
      cellRange.setValue(currentConcerns.join(', '));
      
      // Provide visual feedback
      const originalBackground = cellRange.getBackground();
      cellRange.setBackground('#ffffcc'); // Light yellow highlight
      
      // Reset the background after a short delay (will be executed in batch)
      Utilities.sleep(200);
      cellRange.setBackground(originalBackground);
    });
    
    // Force the spreadsheet to update immediately
    SpreadsheetApp.flush();
    
    return true;
  } catch (error) {
    console.error('Error in updateConcerns:', error);
    return false;
  }
}

/**
 * Save item state to the state storage
 */
function saveState(itemId, assignedTo, status, testOk) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stateSheet = ss.getSheetByName(SHEETS.STATE_STORAGE);
  
  if (!stateSheet) return;
  
  // Find if the item already exists in state storage
  const stateData = stateSheet.getDataRange().getValues();
  let found = false;
  
  for (let i = 1; i < stateData.length; i++) {
    if (stateData[i][0] === itemId) {
      // Update existing state
      stateSheet.getRange(i + 1, 2).setValue(assignedTo);
      stateSheet.getRange(i + 1, 3).setValue(status);
      stateSheet.getRange(i + 1, 4).setValue(testOk);
      stateSheet.getRange(i + 1, 5).setValue(new Date());
      found = true;
      break;
    }
  }
  
  if (!found) {
    // Add new state
    const newRow = stateSheet.getLastRow() + 1;
    stateSheet.getRange(newRow, 1, 1, 5).setValues([[itemId, assignedTo, status, testOk, new Date()]]);
  }
}

/**
 * Save the state of all checklist items to the state storage
 * This ensures no state is lost even when onEdit triggers fail
 * @returns {boolean} True if the operation was successful, false otherwise
 */
function saveAllState() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const checklistSheet = ss.getSheetByName(SHEETS.CHECKLIST);
    const stateSheet = ss.getSheetByName(SHEETS.STATE_STORAGE);
    
    if (!checklistSheet || !stateSheet) {
      console.error('Required sheets not found');
      return false;
    }
    
    // Get all data from the checklist sheet
    const checklistData = checklistSheet.getDataRange().getValues();
    if (checklistData.length <= 1) {
      console.log('No checklist data to save');
      return true; // No data is not an error
    }
    
    // Create a map of item IDs to state information
    const stateMap = {};
    for (let i = 1; i < checklistData.length; i++) {
      const itemId = checklistData[i][0];
      const assignedTo = checklistData[i][1] || '';
      const status = checklistData[i][2] || '';
      const testOk = Boolean(checklistData[i][3]);
      
      // Only save items that have at least one meaningful attribute set
      if (itemId && (assignedTo || status || testOk)) {
        stateMap[itemId] = {
          assignedTo: assignedTo,
          status: status,
          testOk: testOk,
          lastUpdated: new Date()
        };
      }
    }
    
    // Clear existing state data (keep header)
    const lastStateRow = Math.max(stateSheet.getLastRow(), 1);
    if (lastStateRow > 1) {
      stateSheet.getRange(2, 1, lastStateRow - 1, 5).clearContent();
    }
    
    // Create state data rows
    const stateData = [];
    Object.keys(stateMap).forEach(itemId => {
      stateData.push([
        itemId,
        stateMap[itemId].assignedTo,
        stateMap[itemId].status,
        stateMap[itemId].testOk,
        stateMap[itemId].lastUpdated
      ]);
    });
    
    // Write all state data at once
    if (stateData.length > 0) {
      stateSheet.getRange(2, 1, stateData.length, 5).setValues(stateData);
    }
    
    console.log(`Saved state for ${stateData.length} checklist items`);
    return true;
  } catch (error) {
    console.error('Error saving all states:', error);
    return false;
  }
}

/**
 * Extract parent requirement ID from a hierarchical requirement ID
 * e.g. "1.2.3" has parent "1.2"
 */
function getParentId(reqId) {
  const parts = reqId.split('.');
  if (parts.length <= 1) {
    return ''; // No parent for top-level requirements
  }
  return parts.slice(0, parts.length - 1).join('.');
}

/**
 * Main function to generate the checklist based on requirements and concerns
 */
function generateChecklist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Confirm before proceeding
  const response = ui.alert(
    'Generate Checklist',
    'This will rebuild the checklist. State information will be preserved. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    // Save all current states first to ensure nothing is lost
    saveAllState();
    
    // Load requirements
    const reqSheet = ss.getSheetByName(SHEETS.REQUIREMENTS);
    if (!reqSheet) throw new Error('Requirements sheet not found');
    
    const reqData = reqSheet.getDataRange().getValues();
    if (reqData.length <= 1) throw new Error('No requirements data found');
    
    // Load cross-cutting concerns
    const concernsSheet = ss.getSheetByName(SHEETS.CONCERNS);
    if (!concernsSheet) throw new Error('Cross-Cutting Concerns sheet not found');
    
    const concernsData = concernsSheet.getDataRange().getValues();
    
    // Load existing state data
    const stateData = loadStateData();
    
    // Generate the checklist items
    const checklistItems = generateChecklistItems(reqData, concernsData);
    
    // Apply saved states
    restoreState(checklistItems, stateData);
    
    // Write to checklist sheet
    writeChecklistToSheet(checklistItems); 
    
    // Calculate completion percentages and update the Requirements sheet
    calculateCompletionPercentages(reqData, checklistItems);
    
    // Update config
    updateLastGeneratedTime();
    
    ui.alert('Success', 'Checklist has been generated, states have been restored, and completion percentages have been updated.', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'Failed to generate checklist: ' + error.message, ui.ButtonSet.OK);
  }
}

function getConcernIDS(concernsCellText) {
  // Check if we have a concerns cell value
  if (!concernsCellText) return [];
  
  // Handle both comma-separated and space-separated formats
  // First split by commas, then further split by spaces if needed
  const concerns = concernsCellText
    .split(',')                      // Split by commas first
    .flatMap(item => item.trim().split(/\s+/)) // Split each part by whitespace
    .filter(c => c);                 // Remove empty entries
  
  return concerns;
}


/**
 * Generate checklist items based on requirements and concerns
 */
function generateChecklistItems(reqData, concernsData) {
  const items = [];
  const requirements = [];
  const concerns = [];
  const reqConcernMap = {};
  const concernMap = {};
  
  // Find headers for columns
  const reqHeaders = reqData[0];
  const concernsColIndex = reqHeaders.indexOf('Concerns');
  const assignedToColIndex = reqHeaders.indexOf('Assigned to');
  
  // Process requirements (skip header row)
  for (let i = 1; i < reqData.length; i++) {
    const req = {
      id: reqData[i][0],
      description: reqData[i][1],
      itemType: reqData[i][2] || '',
      priority: reqData[i][3],
      sources: reqData[i][4] || '',
      assignedTo: assignedToColIndex >= 0 ? reqData[i][assignedToColIndex] || '' : '',
      concerns: concernsColIndex >= 0 ? getConcernIDS(reqData[i][concernsColIndex] || '').map(c => c.trim()).filter(c => c) : []
    };
    
    if (req.id) {
      requirements.push(req);
      reqConcernMap[req.id] = req.concerns;
    }
  }
  
  // Process concerns (skip header row)
  for (let i = 1; i < concernsData.length; i++) {
    const concern = {
      id: concernsData[i][0],
      description: concernsData[i][1],
      priority: concernsData[i][2],
      sources: concernsData[i][3] || ''
    };
    
    if (concern.id) {
      concerns.push(concern);
      concernMap[concern.id] = concern;
    }
  }
  
  // Build requirement hierarchy for concern propagation
  const reqMap = {};
  requirements.forEach(req => {
    reqMap[req.id] = req;
  });
  
  // Propagate concerns from parents to children
  propagateConcerns(reqMap);
  
  // For each requirement, create a base checklist item
  requirements.forEach(req => {
    const parentId = getParentId(req.id);
    let parentDesc = '';
    if (parentId && reqMap[parentId]) {
      parentDesc = reqMap[parentId].description;
    }
    
    // Add the base requirement item
    const strippedReqId = stripPrefixes(req.id);
    items.push({
      itemId: `ITEM-${strippedReqId}`,
      reqId: req.id,
      reqDesc: req.description,
      itemType: req.itemType,
      parentReq: parentDesc,
      concernId: '',
      concernDesc: '',
      assignedTo: req.assignedTo || '',
      status: '',
      testOk: false
    });
    
    // Get all applicable concerns for this requirement
    const applicableConcernIds = reqMap[req.id].propagatedConcerns || [];
    
    // Add items for each applicable cross-cutting concern
    concerns.forEach(concern => {
      if (applicableConcernIds.includes(concern.id)) {
        const strippedConcernId = stripPrefixes(concern.id);
        items.push({
          itemId: `ITEM-${strippedReqId}-C-${strippedConcernId}`,
          reqId: req.id,
          reqDesc: req.description,
          itemType: req.itemType,
          parentReq: parentDesc,
          concernId: concern.id,
          concernDesc: concern.description,
          assignedTo: req.assignedTo || '',
          status: '',
          testOk: false
        });
      }
    });
  });
  
  return items;
}

/**
 * Propagate concerns from parent requirements to children
 */
function propagateConcerns(reqMap) {
  // Get all requirements
  const reqIds = Object.keys(reqMap);
  
  // First pass: initialize propagatedConcerns with direct concerns
  reqIds.forEach(reqId => {
    reqMap[reqId].propagatedConcerns = [...(reqMap[reqId].concerns || [])];
  });
  
  // Second pass: propagate concerns from parents to children
  reqIds.forEach(reqId => {
    const parentId = getParentId(reqId);
    if (parentId && reqMap[parentId]) {
      // Add parent concerns to this requirement
      reqMap[reqId].propagatedConcerns = [
        ...new Set([
          ...reqMap[reqId].propagatedConcerns,
          ...(reqMap[parentId].propagatedConcerns || [])
        ])
      ];
    }
  });
}

/**
 * Load state data from the state storage sheet
 */
function loadStateData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stateSheet = ss.getSheetByName(SHEETS.STATE_STORAGE);
  
  if (!stateSheet) return {};
  
  const stateData = stateSheet.getDataRange().getValues();
  const stateMap = {};
  
  // Start from row 1 (skip header)
  for (let i = 1; i < stateData.length; i++) {
    const itemId = stateData[i][0];
    const assignedTo = stateData[i][1];
    const status = stateData[i][2];
    const testOk = stateData[i][3];
    
    if (itemId) {
      stateMap[itemId] = {
        assignedTo: assignedTo || '',
        status: status || '',
        testOk: Boolean(testOk) // Convert to boolean
      };
    }
  }
  
  return stateMap;
}

/**
 * Restore saved states to checklist items
 */
function restoreState(checklistItems, stateData) {
  checklistItems.forEach(item => {
    if (stateData[item.itemId]) {
      item.assignedTo = stateData[item.itemId].assignedTo;
      item.status = stateData[item.itemId].status;
      item.testOk = stateData[item.itemId].testOk;
    }
  });
  
  return checklistItems;
}

/**
 * Write checklist items to the checklist sheet
 */
function writeChecklistToSheet(checklistItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let checklistSheet = ss.getSheetByName(SHEETS.CHECKLIST);
  
  if (!checklistSheet) {
    checklistSheet = ss.insertSheet(SHEETS.CHECKLIST);
  }
  
  // Clear existing data (keep header)
  const lastRow = Math.max(checklistSheet.getLastRow(), 1);
  if (lastRow > 1) {
    checklistSheet.getRange(2, 1, lastRow - 1, 10).clearContent();
  }
  
  // Set header if not present
  if (checklistSheet.getRange('A1').getValue() !== 'Item ID') {
    checklistSheet.getRange('A1:J1').setValues([['Item ID', 'Assigned to', 'Status', 'Test OK', 'Item Type', 'Requirement Description', 'CCC Description', 'Requirement ID', 'CCC ID', 'Parent Requirement']]);
    checklistSheet.getRange('A1:J1').setFontWeight('bold');
    checklistSheet.setFrozenRows(1);
  }
  
  // Get project team for assignment dropdown
  const projectTeam = getProjectTeam();
  
  // Add data validation for assigned to column
  const assignedToRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([''].concat(projectTeam), true)
    .build();
  
  // Add data validation for status column
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(STATUS), true)
    .build();
  
  // Add data validation for Test OK column
  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  
  // Prepare data for writing - reordered columns
  const data = checklistItems.map(item => [
    item.itemId,       // Column 1: Item ID
    item.assignedTo,   // Column 2: Assigned to
    item.status,       // Column 3: Status
    item.testOk || false, // Column 4: Test OK 
    item.itemType,     // Column 5: Item Type
    item.reqDesc,      // Column 6: Requirement Description
    item.concernDesc,  // Column 7: CCC Description
    item.reqId,        // Column 8: Requirement ID
    item.concernId,    // Column 9: CCC ID
    item.parentReq     // Column 10: Parent Requirement
  ]);
  
  // Write data
  if (data.length > 0) {
    checklistSheet.getRange(2, 1, data.length, 10).setValues(data);
    checklistSheet.getRange(2, 2, data.length, 1).setDataValidation(assignedToRule); // Assigned to is column 2
    checklistSheet.getRange(2, 3, data.length, 1).setDataValidation(statusRule);     // Status is column 3
    checklistSheet.getRange(2, 4, data.length, 1).setDataValidation(checkboxRule);   // Test OK is column 4
  }
  
  // Format and optimize
  checklistSheet.autoResizeColumns(1, 10);
  checklistSheet.setFrozenColumns(1);
}

function calcBackgroundColorGradient(percentage) {
  // red for 0%, green for 100%, and yellow for 50%, plus a gradient in between
  let red, green;
  if (percentage <= 0) {
    // task is not being worked on
    // return null color
    return null;
  }
  if (percentage <= 50) {
    // From red (0%) to yellow (50%)
    red = 255;
    green = Math.round(255 * (percentage / 50));
  } else {
    // From yellow (50%) to green (100%)
    red = Math.round(255 * (1 - (percentage - 50) / 50));
    green = 255;
  }
  
  const blue = 0; // No blue component
  return `rgb(${red}, ${green}, ${blue})`;
}

/**
 * Calculate completion percentages for all requirements
 */
function calculateCompletionPercentages(reqData, checklistItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reqSheet = ss.getSheetByName(SHEETS.REQUIREMENTS);
  
  // Create a map of all requirements
  const requirements = [];
  const reqMap = {};
  
  // Skip header row
  for (let i = 1; i < reqData.length; i++) {
    const reqId = reqData[i][0];
    if (reqId) {
      requirements.push(reqId);
      reqMap[reqId] = {
        row: i + 1, // Add 1 to account for 0-based index
        children: []
      };
    }
  }
  
  // Build parent-child relationships
  requirements.forEach(reqId => {
    const parentId = getParentId(reqId);
    if (parentId && reqMap[parentId]) {
      reqMap[parentId].children.push(reqId);
    }
  });
  
  // Calculate percentages using recursive function
  requirements.forEach(reqId => {
    const percentage = calculateRequirementCompletion(reqId, reqMap, checklistItems);
    const percentTxt = `${percentage}%`;

    // Find percentage column index
    let percentColIndex = -1;
    const headerRow = reqData[0];
    for (let i = 0; i < headerRow.length; i++) {
      if (headerRow[i] === '%') {
        percentColIndex = i + 1; // Convert to 1-based index
        break;
      }
    }
    
    if (percentColIndex > 0) {
      reqSheet.getRange(reqMap[reqId].row, percentColIndex).setValue(percentTxt);
      reqSheet.getRange(reqMap[reqId].row, percentColIndex).setBackground(calcBackgroundColorGradient(percentage));
      // set foreground color to black for better visibility
      reqSheet.getRange(reqMap[reqId].row, percentColIndex).setFontColor('#000000');
    }
  });
}

/**
 * Recursively calculate completion percentage for a requirement and its children
 */
function calculateRequirementCompletion(reqId, reqMap, checklistItems) {
  // If this is a leaf requirement with no children
  if (reqMap[reqId].children.length === 0) {
    // Count items in the checklist related to this requirement
    let totalItems = 0;
    let shippedItems = 0;
    
    checklistItems.forEach(item => {
      if (item.reqId === reqId) {
        totalItems++;
        if (item.status === STATUS.SHIPPED 
            || item.status === STATUS.REJECTED
            || item.testOk
          ) {  
          shippedItems++;
        }
      }
    });
    
    return totalItems > 0 ? Math.round((shippedItems / totalItems) * 100) : 0;
  } else {
    // For requirements with children, calculate average of child percentages
    let totalPercentage = 0;
    reqMap[reqId].children.forEach(childId => {
      totalPercentage += calculateRequirementCompletion(childId, reqMap, checklistItems);
    });
    
    return Math.round(totalPercentage / reqMap[reqId].children.length);
  }
}

/**
 * Update the last generated timestamp in config
 */
function updateLastGeneratedTime() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  
  if (configSheet) {
    configSheet.getRange('B2').setValue(new Date());
  }
}
