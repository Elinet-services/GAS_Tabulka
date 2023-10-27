//Logging.js

// Log changes
const logChange = (
  sheet,
  row,
  logCol,
  changedCol,
  newValue,
  currentTime,
  headers,
  prevValue,
  existingLog
) => {
  const formattedTime = formatDate(currentTime)
  const logEntry = `${formattedTime}; ${
    headers[changedCol - 1]
  }; Original Value:; ${prevValue}; New Value; ${newValue}`
  const newLog = existingLog ? `${logEntry}\n${existingLog}` : logEntry
  sheet.getRange(row, logCol).setValue(newLog)
}

// Update and log changes
const updateAndLog = (sheet, row, col, currentTime) => {
  // Start logging execution time
  const startTime = new Date().getTime()

  const UPDATED_COL = getColumnIndexByName(sheet, "Aktualizováno")
  const LOG_COL = getColumnIndexByName(sheet, "Log")

  // Batch read multiple values
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
  const newValue = values[col - 1]
  const prevValue = sheet.getRange(row, col).getValue()
  const existingLog = sheet.getRange(row, LOG_COL).getValue()

  if (isNewRow(sheet, row)) {
    setCreationTimeAndStatus(sheet, row, currentTime)
  }

  updateTimestamp(sheet, row, UPDATED_COL, currentTime)
  logChange(
    sheet,
    row,
    LOG_COL,
    col,
    newValue,
    currentTime,
    headers,
    prevValue,
    existingLog
  )

  // End logging execution time
  const endTime = new Date().getTime()
  const executionTime = endTime - startTime
  Logger.log(`Execution time: ${executionTime} ms`)
}
