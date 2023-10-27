//Columns.js
// Get column index by header name
function getColumnIndexByName(sheet, name) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
  return headers.indexOf(name) + 1
}

// Check if it's a new row
const isNewRow = (sheet, row) => {
  const statusCol = getColumnIndexByName(sheet, "Stav")
  const createdCol = getColumnIndexByName(sheet, "Založeno")
  return (
    sheet.getRange(row, statusCol).getValue() === "" &&
    sheet.getRange(row, createdCol).getValue() === ""
  )
}

// Update timestamp
const updateTimestamp = (sheet, row, col, time) => {
  sheet.getRange(row, col).setValue(formatDate(time))
}

// Set creation time and status for new rows
const setCreationTimeAndStatus = (sheet, row, currentTime) => {
  const statusCol = getColumnIndexByName(sheet, "Stav")
  const createdCol = getColumnIndexByName(sheet, "Založeno")
  updateTimestamp(sheet, row, createdCol, currentTime)
  sheet.getRange(row, statusCol).setValue("Assigned")
}

// Remove existing filters
const removeExistingFilters = (sheet) => {
  const filter = sheet.getFilter()
  if (filter) {
    filter.remove()
  }
}

// Apply new filter
const applyFilter = (sheet, colIndex, hiddenValues) => {
  const criteria = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues(hiddenValues)
    .build()
  sheet.getFilter().setColumnFilterCriteria(colIndex, criteria)
}
