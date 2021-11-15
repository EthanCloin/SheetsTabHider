//GLOBAL REFERENCES
const SS = SpreadsheetApp.getActiveSpreadsheet();
const ControlSheet = SS.getSheetByName("SheetHider");
const ctrlMap = {
  sheetName:0, status:1, command:2
}

/**
 * Object containing a list of sheetNames, and booleans representing hiddenStatus
 * Also has methods to hide or reveal sheets
 */
class ControlObject {
  constructor(){
    this.sheetNames = []
    this.hiddenStatus = []
  }
  /**
   * Assigns the param to the sheetNames attribute
   * @param sheetNames Array of strings representing the names of the sheets to be revealed
   */
  setSheetNames(sheetNames){
    this.sheetNames = sheetNames.flat(1);
  }
  /**
   * Assigns the param to the hiddenStatus attribute
   * @param statusArray Array of bools where true is hidden and false is revealed
   */
  setHiddenStatus(statusArray){
    this.hiddenStatus = statusArray;
  }

  /**
   * Uses the Sheets API to hide sheets with names in given array
   * @param sheetNames Array of strings representing the names of the sheets to be hidden
   */
  hideSheets(sheetNames){
    let sheets = sheetNames.map(name => SS.getSheetByName(name))
    sheets.forEach(sheet => sheet.hideSheet())
  }

  /**
   * Uses the Sheets API to reveal sheets with names in given array
   * @param sheetNames Array of strings representing the names of the sheets to be revealed
   */
  revealSheets(sheetNames){
    let sheets = sheetNames.map(name => SS.getSheetByName(name))
    sheets.forEach(sheet => sheet.showSheet())
  }

}
// GLOBAL OBJECT INSTANCE
const Controller = new ControlObject();

/**
 * Collects names of all Sheets in current Spreadsheet and assigns the value to Controller
 */
function getAllSheetNames() {
  let allSheets = SS.getSheets()
  let allSheetNames = allSheets.map(sheet => [sheet.getName()])
  allSheetNames = allSheetNames.filter(sheetName => sheetName != ControlSheet.getName())
  
  Controller.setSheetNames(allSheetNames)
}

/**
 * Collects status of all Sheets visibility in current Spreadsheet and assigns the value to Controller
 */
function getAllHiddenStatus(){
  let consideredSheets = SS.getSheets().filter(x => x.getName() != ControlSheet.getName())
  Controller.setHiddenStatus(consideredSheets.map(sheet => sheet.isSheetHidden()))
}

/**
 * Uses data in Controller to populate the ControlSheet with sheetNames and hiddenStatus
 */
function setControlReport(){
  let formattedSheetNames = Controller.sheetNames.map(x => [x])
  ControlSheet.getRange(2, 1, Controller.sheetNames.length, 1)
              .setValues(formattedSheetNames)

  let formattedStatus = Controller.hiddenStatus.map(function(status){
    return(status ? ["Hidden"] : ["Revealed"])
  })

  ControlSheet.getRange(2, 2, Controller.hiddenStatus.length, 1)
              .setValues(formattedStatus);
}

/**
 * Collects sheetNames and hiddenStatus of all Sheets in current Spreadsheet and updates Controller
 * Implementation is just calling both getAll functions
 */
function refreshReportData(){
  getAllSheetNames();
  getAllHiddenStatus();
}

/**
 * Reads commands from each line of ControlSheet and updates the visibility of the Sheets
 */
function updateIndividualVisibility(){
  // Check 3rd Column for cmds
  let controlSheetData = ControlSheet.getRange(2, 1, Controller.hiddenStatus.length, 3)
              .getValues()
  controlSheetData = controlSheetData.filter(x => x != "")
  
  // Update visibility
  let sheetNamesToReveal = []
  let sheetNamesToHide = []
  controlSheetData.forEach(function(entry){
    // HIDE  
    if (entry[ctrlMap.command] == "HIDE"){
      if (entry[ctrlMap.status] != "Hidden"){
        sheetNamesToHide.push(entry[ctrlMap.sheetName])
      }
    }
    // REVEAL
    if (entry[ctrlMap.command] == "REVEAL"){
      if (entry[ctrlMap.status] != "Revealed"){
        sheetNamesToReveal.push(entry[ctrlMap.sheetName])
      }
    }
  })
  
  // Update Controller
  Controller.hideSheets(sheetNamesToHide)
  Controller.revealSheets(sheetNamesToReveal)
}

/**
 * Reads the BulkCommand from ControlSheet and updates the visibility of all Sheets accordingly
 */
function updateBulkVisibility(){
  // Check 4th Column for Command
  let command = ControlSheet.getRange(2, 4).getValue()
  if (command == "HIDE ALL"){
    Controller.hideSheets(Controller.sheetNames)
  }
  else if (command =="REVEAL ALL"){
    Controller.revealSheets(Controller.sheetNames)
  }
}

/**
 * Sorts the report from third row down sending visible Sheets to top
 */
function sortReportData(){
  ControlSheet.getRange(3, 1, ControlSheet.getLastRow() - 3, 3)
              .sort({column:2, ascending:false})
  ControlSheet.getRange(2, 3, ControlSheet.getLastRow() - 2, 1).clearContent()
}



