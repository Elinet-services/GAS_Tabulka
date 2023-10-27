// Uloží aktuální stav filtru
const setFilterState = (state) => {
  PropertiesService.getScriptProperties().setProperty('filterState', state);
};

// Získá aktuální stav filtru
const getFilterState = () => {
  return PropertiesService.getScriptProperties().getProperty('filterState');
};

// Get column index by header name
function getColumnIndexByName(sheet, name) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(name) + 1;
}

// Check if it's a new row
const isNewRow = (sheet, row) => {
  const statusCol = getColumnIndexByName(sheet, "Stav");
  const createdCol = getColumnIndexByName(sheet, "Založeno");
  return sheet.getRange(row, statusCol).getValue() === '' && 
         sheet.getRange(row, createdCol).getValue() === '';
};

// Format date
const formatDate = (date) => {
  return Utilities.formatDate(date, "GMT+2", "dd/MM/yyyy HH:mm:ss");
}

// Update timestamp
const updateTimestamp = (sheet, row, col, time) => {
  sheet.getRange(row, col).setValue(formatDate(time));
};

// Log changes
const logChange = (sheet, row, logCol, changedCol, newValue, currentTime) => {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const formattedTime = formatDate(currentTime);
  const prevValue = sheet.getRange(row, changedCol).getValue();
  const logEntry = `${formattedTime}; ${headers[changedCol - 1]}; Original Value:; ${prevValue}; New Value; ${newValue}`;
  const existingLog = sheet.getRange(row, logCol).getValue();
  const newLog = existingLog ? `${logEntry}\n${existingLog}` : logEntry;
  sheet.getRange(row, logCol).setValue(newLog);
};

// Update and log changes
const updateAndLog = (sheet, row, col, currentTime) => {
  const UPDATED_COL = getColumnIndexByName(sheet, "Aktualizováno");
  const LOG_COL = getColumnIndexByName(sheet, "Log");

  if (isNewRow(sheet, row)) {
    setCreationTimeAndStatus(sheet, row, currentTime);
  }

  updateTimestamp(sheet, row, UPDATED_COL, currentTime);
  const newValue = sheet.getRange(row, col).getValue();
  logChange(sheet, row, LOG_COL, col, newValue, currentTime);
};

// Set creation time and status for new rows
const setCreationTimeAndStatus = (sheet, row, currentTime) => {
  const statusCol = getColumnIndexByName(sheet, "Stav");
  const createdCol = getColumnIndexByName(sheet, "Založeno");
  updateTimestamp(sheet, row, createdCol, currentTime);
  sheet.getRange(row, statusCol).setValue('Assigned');
};

// Highlight cells based on conditions
const highlightCells = (sheet, row) => {
  const statusCol = getColumnIndexByName(sheet, "Stav");
  const updatedCol = getColumnIndexByName(sheet, "Aktualizováno");
  const currentTime = new Date();
  
  const status = sheet.getRange(row, statusCol).getValue();
  const updatedTime = new Date(sheet.getRange(row, updatedCol).getValue());
  const timeDifference = (currentTime - updatedTime) / 3600000; // Difference in hours
  
  const cell = sheet.getRange(row, updatedCol);
  const format = SpreadsheetApp.newTextStyle();
  
  if (status !== 'Done') {
    if (timeDifference > 96) {
      // Highlight text in bold and set cell background to light red
      cell.setTextStyle(format.setBold(true).build());
      cell.setBackground("#FFCCCC");
    } else if (timeDifference > 32) {
      // Highlight text in bold
      cell.setTextStyle(format.setBold(true).build());
      cell.setBackground(null);
    } else {
      // Remove bold and background color
      cell.setTextStyle(format.setBold(false).build());
      cell.setBackground(null);
    }
  }
};

// On edit event
const onEdit = (e) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = e.range;
  const col = range.getColumn();
  const row = range.getRow();
  const currentTime = new Date();
  const statusCol = getColumnIndexByName(sheet, "Stav");

  if (row === 1) {
    return;
  }

  if (col === statusCol && sheet.getRange(row, statusCol).getValue() === 'Done') {
    const filterState = getFilterState();
    if (filterState === 'active') {
      filterActive();
    } else if (filterState === 'all') {
      filterAll();
    }
  }

  updateAndLog(sheet, row, col, currentTime);
  highlightCells(sheet, row);

  // Update the list of unique Solvers
  // This function is located in the Options.gs file
  updateUniqueSolvers();
};

// On spreadsheet open
const onOpen = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  createMenu();
  
  for (let row = 2; row <= lastRow; row++) {
    highlightCells(sheet, row);
  }
};

// Create custom menu
const createMenu = () => {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Filter Items')
    .addItem('Only Active', 'filterActive')
    .addItem('All Records', 'filterAll')
    .addToUi();
};

// Remove existing filters
const removeExistingFilters = (sheet) => {
  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }
};

// Apply new filter
const applyFilter = (sheet, colIndex, hiddenValues) => {
  const criteria = SpreadsheetApp.newFilterCriteria()
                                .setHiddenValues(hiddenValues)
                                .build();
  sheet.getFilter().setColumnFilterCriteria(colIndex, criteria);
};

// Filter only active items
const filterActive = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  removeExistingFilters(sheet);
  const range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
  range.createFilter();
  const statusCol = getColumnIndexByName(sheet, "Stav");
  applyFilter(sheet, statusCol, ['Done']);
  setFilterState('active');
};

// Remove filter to show all records
const filterAll = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  removeExistingFilters(sheet);
  const range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
  range.createFilter();
  setFilterState('all');
};
