/**
 * Crucial setup function! To be executed ONCE from editor for initializing
 */
function installTrigger(){
  ScriptApp.newTrigger("onOpen")
    .forSpreadsheet(SS)
    .onOpen()
    .create()
  initializeSpreadsheet()
}

/**
 * Creates the Dropdown with button to execute the function
 */
function onOpen(e) {
  const UI = SpreadsheetApp.getUi()
  UI.createMenu("TabHider")
    .addItem("Execute Hide/Reveal Commands", "main")
    .addToUi();
  refreshReportData()
  setControlReport()
  sortReportData()
}

/**
 * Calls functions in proper order to hide/reveal requested sheets
 */
function main(){
  // execute bulk and then individual
  refreshReportData()
  updateBulkVisibility()
  updateIndividualVisibility()
  refreshReportData()
  setControlReport()
  sortReportData()
}

