// Updates the list of unique "Solvers" in the "Options" sheet
function updateUniqueSolvers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Open the main sheet and the "Options" sheet by name
  const mainSheet = ss.getSheetByName("AREL 2310"); 
  const optionsSheet = ss.getSheetByName("Options");
  
  if (!mainSheet || !optionsSheet) {
    return; // Exit the function if either sheet is not found
  }
  
  // Fetch all values in the "Solver" column (Column I)
  const solvers = mainSheet.getRange("I2:I" + mainSheet.getLastRow()).getValues().flat();
  
  // Remove empty values and create a unique list
  const uniqueSolvers = Array.from(new Set(solvers.filter(Boolean)));
  
  // Sort the unique solvers alphabetically
  uniqueSolvers.sort();
  
  // Clear existing values in the "Options" sheet under the "Resources" column (Column A)
  optionsSheet.getRange("A2:A" + optionsSheet.getLastRow()).clear();
  
  // Write unique, sorted values to the "Options" sheet
  optionsSheet.getRange(2, 1, uniqueSolvers.length, 1).setValues(uniqueSolvers.map(solver => [solver]));
}
