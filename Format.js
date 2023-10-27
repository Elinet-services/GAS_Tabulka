//Format.js
// Highlight cells based on conditions
const highlightCells = (sheet, row) => {
  const statusCol = getColumnIndexByName(sheet, "Stav")
  const updatedCol = getColumnIndexByName(sheet, "Aktualizováno")
  const currentTime = new Date()

  const status = sheet.getRange(row, statusCol).getValue()
  const updatedTime = new Date(sheet.getRange(row, updatedCol).getValue())
  const timeDifference = (currentTime - updatedTime) / 3600000 // Difference in hours

  const cell = sheet.getRange(row, updatedCol)
  const format = SpreadsheetApp.newTextStyle()

  if (status !== "Done") {
    if (timeDifference > 96) {
      // Highlight text in bold and set cell background to light red
      cell.setTextStyle(format.setBold(true).build())
      cell.setBackground("#FFCCCC")
    } else if (timeDifference > 32) {
      // Highlight text in bold
      cell.setTextStyle(format.setBold(true).build())
      cell.setBackground(null)
    } else {
      // Remove bold and background color
      cell.setTextStyle(format.setBold(false).build())
      cell.setBackground(null)
    }
  }
  Logger.log(`Status: ${status}`)
  Logger.log(`Updated Time: ${updatedTime}`)
  Logger.log(`Current Time: ${currentTime}`)
  Logger.log(`Time Difference: ${timeDifference}`)
}

// Format date
const formatDate = (date) => {
  return Utilities.formatDate(date, "GMT+2", "dd/MM/yyyy HH:mm:ss")
}
