// -----------------------------------------------------------------------------------------------------------------
// Function: onOpen
// Adds custom menu to the current spreadsheet, can be used to trigger manual executions. 
// -----------------------------------------------------------------------------------------------------------------
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Mission Tracker Link')
      .addItem('Refresh Master Occupancy', 'updateOccupancyReport')
      .addItem('Check Roster for Updates', 'updateRoster')
      .addToUi();
}


// -----------------------------------------------------------------------------------------------------------------
// Function: logIn
// Logs in to Mission Tracker, returns cookie string used to validate all future UrlFetchApp requests
// Replace xxxxx with domain, and log in credentials
// -----------------------------------------------------------------------------------------------------------------
function logIn(){
  var url = "https://xxxxx/login"; // Replace xxxxx with mission tracker domain
  var payload = {
    "username":"xxxx", //Replace xxxx with Mission Tracker User Name
    "password":"xxxx" //Replace xxxx with Mission Tracker pw
  }; 
  var opt = {
    "payload":payload,
    "method":"post",
    "followRedirects" : false
  };
  var response = UrlFetchApp.fetch(encodeURI(url),opt);
  if ( response.getResponseCode() == 200 ) { //could not log in.
    var result = "Login Failed code: " + response.getResponseCode();
  } 
  else if ( response.getResponseCode() == 302 ) { //303 also possible
    var result = "Login successful code: " + response.getResponseCode();
     var cookie = response.getAllHeaders()['Set-Cookie'];     
     var header = {
       'Cookie':cookie[1] //cookie[1] for redirect, cookie[0] if page does not redirect.
     };
  }
  
  Logger.log(result);
  return cookie[1]
}


// -------------------------------------------------------------------------------------------------------
// Function:  updateOccupancyReport
// Download the occupancy report from mission tracker as an .xls, then convert it to google sheets and replace the specified spreadsheet with it.
// This function updates the occupancy report data located on a separate spreadsheet
// A timestamp sheet is appended for importing into other sheets
//
// Replace xxxxx with mission tracker domain 
// Enter the id of the google sheet you are targeting with this update
// -------------------------------------------------------------------------------------------------------

function updateOccupancyReport(){
  var auth = logIn();
  var url = "https://xxxxx/resTracker/reports/general/masterOccReport";
  var OCC_REPORT_ID = "PUT GOOGLE SHEET ID HERE";

  var date = new Date();
  var formattedTime = date.toLocaleTimeString();
  var formattedDate = Utilities.formatDate(date, "GMT-8", "MM/dd/yyyy")
  Logger.log( formattedDate + " " + formattedTime )
  //Utilities.formatDate(d, "GMT-5", "yyyy-MM-dd")
  
  var downloadPayload = 
      {
        "date": formattedDate,  //send current day
        "excel":"Export to Excel" 
      };// trace from Chrome Inspector
  var downloadXls = UrlFetchApp.fetch( url, 
                                  {"headers" : {"Cookie" : auth},
                                   "method" : "post",
                                   "payload" : downloadPayload,
                                  });
  //convert payload to blob
  var xlsFileBlob = downloadXls.getBlob()
  var file = { 
    "title": "Master Occupancy (updated " + formattedDate + " at " + formattedTime + ")" , 
   // "parents": [{"id": folderId}]
  };
  
  
  var occupancyReportSS = SpreadsheetApp.openById(OCC_REPORT_ID)

  //Replace this spreadsheet with downloaded xls file, converted to google sheets
  //To generate new files use file = Drive.Files.insert(file, xlsFileBlob, {
  file = Drive.Files.update(file, occupancyReportSS.getId(), xlsFileBlob, {
    "convert": true
  });
  
  var numSheets = occupancyReportSS.getNumSheets();
  //thisSpreadsheet.insertSheet("Timestamp", numSheets +1).getRange(1, 1).setValue( formattedDate + " " + formattedTime)
  occupancyReportSS.insertSheet("Timestamp", numSheets +1).getRange(1, 1).setValue( formattedDate + " " + formattedTime)
  
}


// -------------------------------------------------------------------------------------------------------
// Function:  updateRoster
// This will compare a local version of the occupancy report to the current one found on mission tracker and 
// display the differences
//
//  
// 
// -------------------------------------------------------------------------------------------------------

// This version of updateRoster utilizes the local roster, and shows changes from mission tracker in the Green Side as suggested changes
function updateRoster() {
  //var rosterBlankIndex = [];
  var date = new Date();
  var formattedTime = date.toLocaleTimeString();
  var formattedDate = Utilities.formatDate(date, "GMT-8", "MM/dd/yyyy")
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var rosterSheet = ss.getSheetByName("Roster")
  var existingRosterRange = rosterSheet.getDataRange()
  var existingRosterData = existingRosterRange.getValues()
  //var rosterDataFlat = rosterRange.getValues().reduce(function (a, b) { //flatten the 2D array obtained by .getValues()
  //return a.concat(b);
  //});
  //for (var i in existingRosterData){
  //  Logger.log(existingRosterData[i][1] + " " + existingRosterData[i][5])
  //}
  
  var occSheet = ss.getSheetByName("Occupancy Data Link")
  var occRange = occSheet.getDataRange()
  var occData = occRange.getValues()
  //var mtDataFlat = mtRange.getValues().reduce(function (a, b) { //flatten the 2D array obtained by .getValues()
  //return a.concat(b);
  //});
   
  var occDataCleaned = []
  for (var i in occData){
    var occName = occData[i][0]  //Hidden Trimmed name in Column A to prevent Spacing failures in searches
    
    var occCM = occData[i][5]
    if (occCM == ""){
      occCM = "BLANK!";
    }
    var occClass = occData[i][6]
    if (occClass == ""){
      occClass = "BLANK!";
    }
    var occWork = occData[i][7]
    if (occWork == ""){
      occWork = "BLANK!";
    }
    var occStatus = ""
    var occExtras = ""
    var classDiff = ""
    var workDiff = ""
    var cmDiff = ""
    
    //&& occWork != "College Student"
    if( occName != "" && occWork != "Child" ){
      //Logger.log(occName.trim() + " " + occClass + " " + occWork)
      
      //Restore any previous entered Extra Duties
      for (var j in existingRosterData){
        if ( existingRosterData[j][1] == occName ){
          if ( existingRosterData[j][2] != occClass) {
            //classDiff = existingRosterData[j][2]
            classDiff = occClass
            occClass = existingRosterData[j][2]
            Logger.log("Diff Found: Class: Existing: " + existingRosterData[j][2] + " OCC: " + occClass)
            Logger.log(classDiff)
          }
          if ( existingRosterData[j][3] != occWork) {
            //workDiff = existingRosterData[j][3]
            workDiff = occWork
            occWork = existingRosterData[j][3]
            Logger.log("Diff Found: Work: Existing: " + existingRosterData[j][3] + " OCC: " + occWork)
            Logger.log(workDiff)
          }
          if ( existingRosterData[j][4] != occCM) {
            //cmDiff = existingRosterData[j][4]
            cmDiff = occCM
            occCM = existingRosterData[j][4]
            Logger.log("Diff Found: CM: Existing: " + existingRosterData[j][4] + " OCC: " + occCM)
            Logger.log(cmDiff)
          }
          
         // if ( existingRoster
          occStatus = existingRosterData[j][5]
          occExtras = existingRosterData[j][6]
        }
        
      }
      occDataCleaned.push([occName.trim(), occClass.trim(), occWork, occCM.trim(), occStatus, occExtras, classDiff, workDiff, cmDiff ])
    }
    
  }
  var newRosterRange = rosterSheet.getRange(3, 2, occDataCleaned.length, 9)
  
  //remove old data
  //rosterSheet.clear()
  rosterSheet.getRange(3, 2, existingRosterData.length, 9).clearContent()
  
  //Sort Data By Work Assign
  occDataCleaned.sort(function(x,y){
    var xp = x[2]; // 3rd column
    var yp = y[2]; // 3rd column 
    return xp == yp ? 0 : xp < yp ? -1 : 1;
    });
  
  //push new data to Spreadsheet and timestamp
  newRosterRange.setValues(occDataCleaned)
  rosterSheet.getRange("C1").setValue(formattedDate + " " + formattedTime)
  DriveApp.getFileById(ss.getId()).setName("Summary (updated: " + formattedDate + " at " + formattedTime + ")" )
  
}



