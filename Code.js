//Code.js
let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
let activeSheet = activeSpreadsheet.getActiveSheet()

// Save the current filter state
const setFilterState = (state) => {
  PropertiesService.getScriptProperties().setProperty("filterState", state)
}

// Retrieve the current filter state
const getFilterState = () => {
  return PropertiesService.getScriptProperties().getProperty("filterState")
}

// On edit event
const onEdit = (e) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const sheetName = sheet.getName()
  const range = e.range
  const col = range.getColumn()
  const row = range.getRow()
  const currentTime = new Date()
  const statusCol = getColumnIndexByName(sheet, "Stav")
  const updatedCol = getColumnIndexByName(sheet, "Aktualizováno") // Get the index of "Aktualizováno"

  if (sheetName !== "AREL 2310" || row === 1 || col === updatedCol) {
    return // Skip if the sheet is not "AREL 2310", or it's the header row, or the "Aktualizováno" column
  }

  // Set the function as running
  PropertiesService.getScriptProperties().setProperty("isRunning", "true")

  // Generate button for new rows
  if (isNewRow(sheet, row)) {
    sheet.getRange(row, 1).setValue("Vypnuto") // Initial state
  }

  // Toggle button state
  if (col === 1 && row > 1) {
    // Assuming "Prio" is the first column
    const cell = sheet.getRange(row, 1)
    cell.setValue(cell.getValue() === "Zapnuto" ? "Vypnuto" : "Zapnuto")
  }

  // Automatically turn off the button if the status is "Done"
  if (
    col === statusCol &&
    sheet.getRange(row, statusCol).getValue() === "Done"
  ) {
    sheet.getRange(row, 1).setValue("Vypnuto")
  }

  // Existing code for filtering based on status
  if (
    col === statusCol &&
    sheet.getRange(row, statusCol).getValue() === "Done"
  ) {
    const filterState = getFilterState()
    if (filterState === "active") {
      filterActive()
    } else if (filterState === "all") {
      filterAll()
    }
  }
}
// Update and log changes
const updateAndLog = (sheet, row, col, currentTime) => {
  const UPDATED_COL = getColumnIndexByName(sheet, "Aktualizováno")
  const LOG_COL = getColumnIndexByName(sheet, "Log")

  if (isNewRow(sheet, row)) {
    setCreationTimeAndStatus(sheet, row, currentTime)
  }

  // Only update the "Aktualizováno" column if it wasn't manually edited
  if (col !== UPDATED_COL) {
    updateTimestamp(sheet, row, UPDATED_COL, currentTime)
  }

  const newValue = sheet.getRange(row, col).getValue()
  logChange(sheet, row, LOG_COL, col, newValue, currentTime)
}

// On spreadsheet open
const onOpen = () => {
  const sheet = activeSheet
  const lastRow = sheet.getLastRow()

  createMenu()

  for (let row = 2; row <= lastRow; row++) {
    highlightCells(sheet, row)
  }
}

// Create custom menu
const createMenu = () => {
  const ui = SpreadsheetApp.getUi()
  ui.createMenu("Filter Items")
    .addItem("Only Active", "filterActive")
    .addItem("All Records", "filterAll")
    .addToUi()
}

// Filter only active items
const filterActive = () => {
  const sheet = activeSheet
  const lastRow = sheet.getLastRow()
  removeExistingFilters(sheet)
  const range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn())
  range.createFilter()
  const statusCol = getColumnIndexByName(sheet, "Stav")
  applyFilter(sheet, statusCol, ["Done"])
  setFilterState("active")
}

// Remove filter to show all records
const filterAll = () => {
  const sheet = activeSheet
  const lastRow = sheet.getLastRow()
  removeExistingFilters(sheet)
  const range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn())
  range.createFilter()
  setFilterState("all")
}
