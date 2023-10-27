//Options.js
// Updates the list of unique "Solvers" in the "Options" sheet
function updateUniqueSolvers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()

  // Open the main sheet and the "Options" sheet by name
  const mainSheet = ss.getSheetByName("AREL 2310")
  const optionsSheet = ss.getSheetByName("Options")

  Logger.log(`Main Sheet: ${mainSheet ? "Found" : "Not Found"}`)
  Logger.log(`Options Sheet: ${optionsSheet ? "Found" : "Not Found"}`)

  if (!mainSheet || !optionsSheet) {
    Logger.log("Exiting function due to missing sheet.")
    return // Exit the function if either sheet is not found
  }

  // Fetch all values in the "Solver" column (Column I)
  const solvers = mainSheet
    .getRange("I2:I" + mainSheet.getLastRow())
    .getValues()
    .flat()

  Logger.log(`Solvers: ${solvers}`)

  // Fetch existing values in the "Options" sheet under the "Resources" column (Column A)
  const existingOptions = optionsSheet
    .getRange("A2:A" + optionsSheet.getLastRow())
    .getValues()
    .flat()

  Logger.log(`Existing Options: ${existingOptions}`)

  // Combine the two lists and remove empty values
  const combinedSolvers = solvers.concat(existingOptions).filter(Boolean)

  Logger.log(`Combined Solvers: ${combinedSolvers}`)

  // Create a unique list
  const uniqueSolvers = Array.from(new Set(combinedSolvers))

  // Sort the unique solvers alphabetically
  uniqueSolvers.sort()

  Logger.log(`Unique Solvers: ${uniqueSolvers}`)

  // Clear existing values in the "Options" sheet under the "Resources" column (Column A)
  optionsSheet.getRange("A2:A" + optionsSheet.getLastRow()).clear()

  // Write unique, sorted values to the "Options" sheet
  optionsSheet
    .getRange(2, 1, uniqueSolvers.length, 1)
    .setValues(uniqueSolvers.map((solver) => [solver]))

  Logger.log("Updated unique solvers in Options sheet.")
}
