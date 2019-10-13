// -----------------------------------------------------------------------------------------------------------------
// Function: onOpen
// Adds custom menu to the current spreadsheet, can be used to trigger manual executions. 
// -----------------------------------------------------------------------------------------------------------------
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Mission Tracker')
      .addItem('Update Roster', 'updateRoster')
      //.addItem('Update Form Names', 'updateFormNames')
      .addToUi();
}


// -------------------------------------------------------------------------------------------------------
// Function:  updateRoster
// This will compare a local roster to the occupancy report, adding any new intakes and marking any that exit 
// 
//
//  
// 
// -------------------------------------------------------------------------------------------------------

function updateRoster() {
  var rosterBlankIndex = [];
  var date = new Date();
  var formattedTime = date.toLocaleTimeString();
  var formattedDate = Utilities.formatDate(date, "GMT-8", "MM/dd/yyyy")
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rosterSheet = ss.getSheetByName("Roster")
  var rosterRange = rosterSheet.getRange("B3:I200") //consider replacing with getDataRange
  var rosterData = rosterRange.getValues()
  var rosterDataFlat = rosterRange.getValues().reduce(function (a, b) { //flatten the 2D array obtained by .getValues()
    return a.concat(b);
  });
  
  var mtSheet = ss.getSheetByName("Mission Tracker Link")
  var mtRange = mtSheet.getRange("A3:N29")  // consider replacing with getDataRange
  var mtData = mtRange.getValues()
  var mtDataFlat = mtRange.getValues().reduce(function (a, b) { //flatten the 2D array obtained by .getValues()
    return a.concat(b);
  });
  
  var historicalSheet = ss.getSheetByName("Historical Visits Link")
  var historicalRange = historicalSheet.getDataRange()  //
  var historicalData = historicalRange.getValues()
  var historicalDataFlat = historicalRange.getValues().reduce(function (a, b) {
    return a.concat(b);
  });
  
  Logger.log(historicalData)
  Logger.log(historicalDataFlat)
   
  //Search for People Moved out-
  for (var i in rosterData)
  {
    var rosterName = rosterData[i][0]
    var nameFoundAtIndex = mtDataFlat.indexOf(rosterName)
    
    // if name not found, and exit date is blank, they have moved out, set exit date to today
    if (nameFoundAtIndex == -1 && rosterData[i][2] == "") {
      // change to matched exit date from Historical Visits Report
      rosterData[i][2] = formattedDate 
    }
    if ( rosterName == "" ){
      rosterBlankIndex.push(i)
    }
    //Logger.log(i + " [" + rosterName + "] Found in MT at index: " + nameFoundAtIndex + " Exit Date of " + rosterData[i][2])
  }
  
  //Search for new Adds-
  for (var i in mtData)
  {
    var mtName = mtData[i][0]
    var mtIntake = mtData[i][4]
    var mtVetStatus = mtData[i][11]
    var mtWorkAssign = mtData[i][7]
    var mtAge = mtData[i][13]
    var mtNameFoundAtIndex = rosterDataFlat.indexOf(mtName)
    
    // Intake Detected
    if (mtNameFoundAtIndex == -1) {
      var blankIndex = rosterBlankIndex[0]
      Logger.log( "Add found. Adding " + mtName + " to RosterData at index: " + blankIndex + " Intake date of : " + mtIntake)
      rosterData[blankIndex][0] = mtName
      rosterData[blankIndex][1] = mtIntake
      if ( mtVetStatus == "Yes" ){
        rosterData[blankIndex][4] = "Veteran"
        if ( mtWorkAssign == "College Student") {
          rosterData[blankIndex][6] = "College"
        } else {
          rosterData[blankIndex][6] = "In Program"
        }
        
      } else if ( mtVetStatus =="No" && mtWorkAssign == "Child" ){
        if (mtAge >= 4){
          rosterData[blankIndex][4] = "Child"
          rosterData[blankIndex][6] = "School K-12"
        }else if (mtAge < 4) {
          rosterData[blankIndex][4] = "Child"
          rosterData[blankIndex][6] = "None"
        }
      } else if ( mtVetStatus == "No" ){
        rosterData[blankIndex][4] = "Spouse"
        if ( mtWorkAssign == "College Student") {
          rosterData[blankIndex][6] = "College"
        } else {
          rosterData[blankIndex][6] = "In Program"
        }
      }
      
      
      rosterBlankIndex.shift()
      Logger.log("New Blank Index is " + rosterBlankIndex[0])
    }
    
    Logger.log(i + " [" + mtName + "] Found in Roster at index: " + mtNameFoundAtIndex + " Vet Status and Age: " + mtVetStatus + mtAge + " Work Assignment: " + mtWorkAssign )
  }
  //Logger.log(rosterData)
  
  //Send all Data To Spreadsheet
  rosterRange.setValues(rosterData) 
  rosterSheet.getRange("C1").setValue(formattedDate + " " + formattedTime)
}

// -------------------------------------------------------------------------------------------------------
// Function:  updateRoster
//Send Current TVO Adults as choices to the Name field of a Schedule Form
//
// replace xxx with formId and ItemID
// 
// -------------------------------------------------------------------------------------------------------

function updateFormNames(){
  // call your form and connect to the drop-down item
  var form = FormApp.openById("xxxxx");
   
  var namesList = form.getItemById("xxxxx").asListItem();

// identify the sheet where the data resides needed to populate the drop-down
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mtSheet = ss.getSheetByName("Mission Tracker Link")
  var mtRange = mtSheet.getRange("A3:N29")  // consider replacing with getDataRange
  var mtData = mtRange.getValues()
  
  
  var namesToSend = []
  for (var i in mtData){
    var mtName = mtData[i][0]
    //var mtIntake = mtData[i][4]
    //var mtVetStatus = mtData[i][11]
    var mtWorkAssign = mtData[i][7]
    //var mtAge = mtData[i][13]
    
    //To filter out college: && mtWorkAssign != "College Student"
    if ( mtName != "" && mtWorkAssign != "Child" && mtWorkAssign != "Alumni" ){
      namesToSend.push(mtName)
    }
    
  }
  namesList.setChoiceValues(namesToSend)
  
}
