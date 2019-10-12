// -----------------------------------------------------------------------------------------------------------------
// SAMPLE HEADER
//entryDate: 07/01/2018
//exitDate: 06/30/2019
//fName: 
//lName: 
//buildingID: 15  //  
//incDeletedProfiles: 0
//page: 1
//search: search
//excel: Export to Excel
// -----------------------------------------------------------------------------------------------------------------


// -----------------------------------------------------------------------------------------------------------------
// Function: onOpen
// Adds custom menu to the current spreadsheet, can be used to trigger manual executions. 
// -----------------------------------------------------------------------------------------------------------------
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Mission Tracker Link')
      .addItem('Update Report', 'updateHistoricalVisits')
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
// Function:  updateHistoricalVisits
// Download the Historical Visits report from mission tracker as an .xls, then convert it to google sheets and replace 
// the current spreadsheet with it.
// A timestamp sheet is appended for importing into other sheets
//
// Replace xxxxx with mission tracker domain 
// Replace header values with what is needed for report
// -------------------------------------------------------------------------------------------------------
function updateHistoricalVisits(){
  var auth = logIn();
  var url = "https://xxxxx/resTracker/reports/general/historicalVisits";
  
  // REPLACE THESE VALUES AS NEEDED FOR REPORT
  var entryDate = "07/01/2019";
  var exitDate = "06/30/2020";
  var fName = "";
  var lName = "";
  var buildingID = 15;

  var date = new Date();
  var formattedTime = date.toLocaleTimeString();
  var formattedDate = Utilities.formatDate(date, "GMT-8", "MM/dd/yyyy")
  Logger.log( formattedDate + " " + formattedTime )
  //Utilities.formatDate(d, "GMT-5", "yyyy-MM-dd")
  
  var downloadPayload = 
      {
        "entryDate": entryDate,
        "exitDate": exitDate,
        "fName": fName,
        "lName": lName,
        "buildingID": buildingID,
        "incDeletedProfiles": 0,
        "page": 1,
        "search": "search",
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
    "title": "TVO Historical Visits (updated " + formattedDate + " at " + formattedTime + ")" , 
   // "parents": [{"id": folderId}]
  };
  
  var thisSpreadsheet = SpreadsheetApp.getActiveSpreadsheet() 
  //Replace this spreadsheet with downloaded xls file, converted to google sheets
  //To generate new files use file = Drive.Files.insert(file, xlsFileBlob, {
  
  file = Drive.Files.update(file, thisSpreadsheet.getId(), xlsFileBlob, {
    "convert": true
  });
  
  var numSheets = thisSpreadsheet.getNumSheets();
  thisSpreadsheet.insertSheet("Timestamp", numSheets +1).getRange(1, 1).setValue( formattedDate + " " + formattedTime)
  
}
