// -----------------------------------------------------------------------------------------------------------------
// Add Menu to Spreadsheet
// -----------------------------------------------------------------------------------------------------------------
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu( "Mission Tracker" )
  //.addItem("Combine Tables", "combineTables")
  .addSubMenu(ui.createMenu("Rebuild Database")
                .addItem('Pre Intake', 'rebuildPreIntakeAuto')
                .addItem('At A Glance', 'rebuildAtAGlanceAuto')
                .addItem('Visits', 'rebuildVisitsAuto') 
               )
  .addItem("Show Sidebar", "sidebarMenuItem")
  .addItem("Clear Log", "logSClear")
  .addToUi();
}

//Confirmation prompt for expensive rebuilds
function showRebuildAlert( ) {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert(
     'Are you sure?',
     'This will erase all existing data.',
      ui.ButtonSet.YES_NO);
  return result
}

// -----------------------------------------------------------------------------------------------------------------
// Display HTML in sidebar
// Replace xxxxx
// -----------------------------------------------------------------------------------------------------------------
function sidebarMenuItem() {
  var baseUrlTest = "https://xxxxx"
  var auth = logIn()
  var sideBar = HtmlService.createHtmlOutputFromFile("profile").setTitle("HTML Test");
  //var sideBar = HtmlService.createHtmlOutput(getProfileData(baseUrlTest, auth)).setTitle("Profile" + 976)

  SpreadsheetApp.getUi().showSidebar(sideBar);
}

function combineTables () {
  var combinedData = []
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var preIntakeSheet = ss.getSheetByName("PreIntake")
  var preIntakeRange = preIntakeSheet.getDataRange()
  var preIntakeData = preIntakeRange.getValues()
  //var preIntakeBlank = ["", "", "", "", "", "", ""]
  
  var atAGlanceSheet = ss.getSheetByName("AtAGlance")
  var atAGlanceRange = atAGlanceSheet.getDataRange()
  var atAGlanceData = atAGlanceRange.getValues()
  //var atAGlanceBlank = ["", "", "", "", "", "", "", "", ""]
  
  var visitsSheet = ss.getSheetByName("Visits")
  var visitsRange = visitsSheet.getDataRange()
  var visitsData = visitsRange.getValues()
  //var visitsBlank = ["", "", "", "", "", ""]
  
  var skipLogSheet = ss.getSheetByName("Log")
  var skipLogRange = skipLogSheet.getDataRange()
  var skipLogData = skipLogRange.getValues()
  
  
  
  for (var i = 0; i < preIntakeData.length; i++){
    var id = preIntakeData[i][0]
    //var newBlank = [i, "", "", "", "", "", "", ""]
    if ( typeof id === "number"){
      combinedData.push(preIntakeData[i])
    }
  }
  
  for (var i = 0; i < combinedData.length; i++){
    var combinedId = combinedData[i][0]
    for (var a = 0; a < atAGlanceData.length; i++) {
      if (combinedId === atAGlanceData[a][0] ){
        //combinedData
      }
    }
    
  }
  
  //for (var i = 0; i< 100; i++){
  //  Logger.log("element i[0] "+ i + " " + typeof preIntakeData[i][0])
  //}


  
  //var preIntakeFlat = preIntakeRange.getValues().reduce(function (a, b) { //flatten the 2D array obtained by .getValues()
  //  return a.concat(b);
  //});
  
  clearDataSheet("Combined")
  writeArrayToSpreadsheet(combinedData, 1, "Combined")
  
}




// -----------------------------------------------------------------------------------------------------------------
// Top Function, Trigger from Menu
// -----------------------------------------------------------------------------------------------------------------
function rebuildPreIntakeAuto() {
  var ui = SpreadsheetApp.getUi();
  var confirm = showRebuildAlert()
  
  if (confirm === ui.Button.YES) {
    ui.alert("Executing rebuild for: PreIntake.  This will usually take 20-30 minutes.")
    // resets the loop counter if it's not 0
    refreshUserProperties("PreIntake");
    // clear all the sheets
    clearDataSheet("PreIntake");
    // create trigger to update DB in chunks
    createTrigger("buildPreIntake");
  } else {
    ui.alert("Operation Cancelled")
    return
  }

}

// -----------------------------------------------------------------------------------------------------------------
//  
// -----------------------------------------------------------------------------------------------------------------
function buildPreIntake(){
  const PRE_INTAKE_COUNTER_KEY = "PreIntakeCounter"
  var userProperties = PropertiesService.getUserProperties();
  var loopCounter = Number(userProperties.getProperty(PRE_INTAKE_COUNTER_KEY));
  //var loopCounter = 4
  const brickSize = 1000
  const loopLimit = 4
  var startAtRecord = (loopCounter * brickSize) - brickSize + 1;
  var endAtRecord = loopCounter * brickSize
  //var dataToWrite = []
  Logger.log("buildPreIntake: Building brick: " + loopCounter + " of " + loopLimit)
  Logger.log("buildPreIntake: Brick size: " + brickSize)
  Logger.log("buildPreIntake: Scraping profiles: " + startAtRecord + " to " + endAtRecord)
  
  if (loopCounter <= loopLimit ){
    loopCounter++;
    Logger.log("buildPreIntake: PreIntakeCounter incremented to: " + loopCounter)
    userProperties.setProperty(PRE_INTAKE_COUNTER_KEY, loopCounter);
    
    var dataToWrite = scrapePreIntake(startAtRecord, endAtRecord)

    //Logger.log("buildPreIntake: scrapeAtAGlance complete for: " + startAtRecord + " to " + endAtRecord)
    if (dataToWrite == null  || dataToWrite.length == 0) {
      Logger.log("buildPreIntake:  scrapePreIntake() returned no data.  Dumping result.  Attempting to terminate execution.")
      userProperties.setProperty(PRE_INTAKE_COUNTER_KEY, 99);
      Logger.log("buildPreIntake: failed result: \n" + dataToWrite)
      logS()
      dumpS(dataToWrite) 
      
    } else{
      Logger.log("buildPreIntake: scrape successful, writing")
      writeArrayToSpreadsheet(dataToWrite, startAtRecord, "PreIntake")
      logS()
    }
  } else {
    // Build limit reached or exceeded, terminate the triggers and sort the results
    Logger.log("buildPreIntake: Execution limit of " + loopLimit + " exceeded. Count: " + loopCounter)
    deleteAllTriggers("buildPreIntake")
    sortDataSheet("PreIntake")
    logS()  
  }

}

// -----------------------------------------------------------------------------------------------------------------
// Top Function, Trigger from Menu
// -----------------------------------------------------------------------------------------------------------------
function rebuildAtAGlanceAuto() {
  var ui = SpreadsheetApp.getUi();
  var confirm = showRebuildAlert()
  
  if (confirm === ui.Button.YES) { 
    ui.alert("Executing rebuild for: At A Glance.  This will usually take 20-30 minutes.")
    // resets the loop counter if it's not 0
    refreshUserProperties("AtAGlance");
    // clear all the sheets
    clearDataSheet("AtAGlance");
    // create trigger to update DB in chunks
    createTrigger("buildAtAGlance");
  } else {
    ui.alert("Operation Cancelled")
    return
  }
  
}

// -----------------------------------------------------------------------------------------------------------------
//  
// -----------------------------------------------------------------------------------------------------------------
function buildAtAGlance(){
  const AAG_COUNTER_KEY = "AtAGlanceCounter"
  var userProperties = PropertiesService.getUserProperties();
  var loopCounter = Number(userProperties.getProperty(AAG_COUNTER_KEY));
  //var loopCounter = 1
  const brickSize = 1000
  const loopLimit = 4
  var startAtRecord = (loopCounter * brickSize) - brickSize + 1;
  var endAtRecord = loopCounter * brickSize
  //var dataToWrite = []
  Logger.log("buildAtAGlance: Building brick: " + loopCounter + " of " + loopLimit)
  Logger.log("buildAtAGlance: Brick size: " + brickSize)
  Logger.log("buildAtAGlance: Scraping profiles: " + startAtRecord + " to " + endAtRecord)
  
  if (loopCounter <= loopLimit ){
    loopCounter++;
    Logger.log("buildAtAGlance: AtAGlanceCounter incremented to: " + loopCounter)
    userProperties.setProperty(AAG_COUNTER_KEY, loopCounter);
   
    var dataToWrite = scrapeAtAGlance(startAtRecord, endAtRecord)

    //Logger.log("buildAtAGlance: scrapeAtAGlance complete for: " + startAtRecord + " to " + endAtRecord)
    if (dataToWrite == null  || dataToWrite.length == 0) {
      Logger.log("buildAtAGlance:  scrapeAtAGlance() returned no data.  Dumping result.  Attempting to terminate execution.")
      userProperties.setProperty(AAG_COUNTER_KEY, 99);
      Logger.log("buildAtAGlance: failed result: \n" + dataToWrite)
      logS()
      dumpS(dataToWrite) 
      
    } else{
      Logger.log("buildAtAGlance: Scrape complete. Writing to spreadsheet.")
      writeArrayToSpreadsheet(dataToWrite, startAtRecord, "AtAGlance")
      logS()
    }
  } else {
    // Build limit reached or exceeded, terminate the triggers and sort the results
    Logger.log("buildAtAGlance: Execution limit of " + loopLimit + " exceeded. Count: " + loopCounter)
    deleteAllTriggers("buildAtAGlance")
    sortDataSheet("AtAGlance")
    logS()  
  }

}

// -----------------------------------------------------------------------------------------------------------------
// Top Function, Trigger from Menu
// -----------------------------------------------------------------------------------------------------------------
function rebuildVisitsAuto() {
  var ui = SpreadsheetApp.getUi();
  var confirm = showRebuildAlert()
  
  if (confirm === ui.Button.YES) {
    ui.alert("Executing rebuild for: Visits.  This will usually take 40-50 minutes.")
    // resets the loop counter if it's not 0
    refreshUserProperties("Visits");
    // clear all the sheets
    clearDataSheet("Visits");
    // create trigger to update DB in chunks
    createTrigger("buildVisits");
  } else {
    ui.alert("Operation Cancelled")
    return
  }

}

function buildVisits(){
  const VISITS_COUNTER_KEY = "VisitsCounter"
  var userProperties = PropertiesService.getUserProperties();
  var loopCounter = Number(userProperties.getProperty(VISITS_COUNTER_KEY));
  //var loopCounter = 1
  const brickSize = 500
  const loopLimit = 8
  var startAtRecord = (loopCounter * brickSize) - brickSize + 1;
  var endAtRecord = loopCounter * brickSize
  
  Logger.log("buildVisits: Building brick: " + loopCounter + " of " + loopLimit)
  Logger.log("buildVisits: Brick size: " + brickSize)
  Logger.log("buildVisits: Scraping profiles: " + startAtRecord + " to " + endAtRecord)
  
  if (loopCounter <= loopLimit ){
    loopCounter++;
    Logger.log("buildVisits: VisitsCounter incremented to: " + loopCounter)
    userProperties.setProperty(VISITS_COUNTER_KEY, loopCounter);

    var dataToWrite = scrapeOpenClose(startAtRecord, endAtRecord)

    //Logger.log("buildVisits: scrapeAtAGlance complete for: " + startAtRecord + " to " + endAtRecord)
    if (dataToWrite == null  || dataToWrite.length == 0) {
      Logger.log("buildVisits:  scrapeOpenClose() returned no data.  Dumping result.  Attempting to terminate execution.")
      userProperties.setProperty(VISITS_COUNTER_KEY, 99);
      Logger.log("buildVisits: failed result: \n" + dataToWrite)
      logS()
      dumpS(dataToWrite) 
      
    } else{
      Logger.log("buildVisits: Scrape complete. Writing to spreadsheet")
      writeArrayToSpreadsheet(dataToWrite, startAtRecord, "Visits")
      logS()
    }
  } else {
    // Build limit reached or exceeded, terminate the triggers and sort the results
    Logger.log("buildVisits: Execution limit of " + loopLimit + " exceeded. Count: " + loopCounter)
    deleteAllTriggers("buildVisits")
    sortDataSheet("Visits")
    logS()  
  }

}

// -----------------------------------------------------------------------------------------------------------------
// Reset Loop Counter, persists across executions
// -----------------------------------------------------------------------------------------------------------------
function refreshUserProperties( typeString ) {
  var userProperties = PropertiesService.getUserProperties();
  var loopCounterType = typeString + "Counter"
  userProperties.setProperty(loopCounterType, 1);
  userProperties.setProperty("loopCounter", 1)
  userProperties.setProperty("urlFetchCount", 0)
}

// -----------------------------------------------------------------------------------------------------------------
// Create Trigger to build a chunk every 5 minutes
// -----------------------------------------------------------------------------------------------------------------
function createTrigger(funcName) {  
  if (funcName == null) {
    Logger.log("createTrigger: Failed.  funcName parameter is null")
    return
  }else {
    ScriptApp.newTrigger(funcName)
      .timeBased()
      .everyMinutes(5)
      .create();
    Logger.log("createTrigger: Successful for " + funcName + " every 5 minutes.")
  }
  
}

// -----------------------------------------------------------------------------------------------------------------
// Erase all triggers from project after loop limit is reached
// -----------------------------------------------------------------------------------------------------------------
function deleteAllTriggers( functionName ) {
  // Loop over all triggers and delete them
  var allTriggers = ScriptApp.getProjectTriggers();
  if (functionName == null) {
    Logger.log("deleteAllTriggers: function name not specified, Removing all triggers")
  }
  
  for (var i = 0; i < allTriggers.length; i++) {
    var currentTriggerName = allTriggers[i].getHandlerFunction()
    if (functionName === currentTriggerName || functionName == null ) {
      ScriptApp.deleteTrigger(allTriggers[i])
      Logger.log("deleteAllTriggers: currentTriggerName: " + currentTriggerName + " Trigger Removed")
    } else{
      Logger.log("deleteAllTriggers: currentTriggerName: not matched, No Trigger Removed")
    }
  }
}

// -----------------------------------------------------------------------------------------------------------------
// Replace xxxxx
//"https://xxxx/resTracker/editresident/profile/preintake/"
// Values retrieved: ID, fName, mName, lName, dob, cusInitialContactDate, screenDate, interviewDate
// -----------------------------------------------------------------------------------------------------------------
function scrapePreIntake (start, end, idArray) {
  // starting and ending profiles for testing this function from editor
  if (start == null) { start = 1 }  
  if (end == null) { end = 10}  
  
  // if an array of id numbers is passed, scrape that list of ids instead of using start and end in order
  var useIdArray = false
  if ( idArray != null && idArray.length > 0) {
    start = 0
    end = idArray.length - 1
    useIdArray = true
  }
  
  var preIntakeUrl = "https://xxxxx/resTracker/editresident/profile/preintake/"
  var authCookie = logIn()
  var scrapedData = []
  var skippedIds = []
  //var fetchCount = 0
  
  const fNameRegex = /id="fName" value="(.*?)"/;
  const mNameRegex = /id="mName" value="(.*?)"/;
  const lNameRegex = /id="lName" value="(.*?)"/;
  const birthRegex = /id="dob" value="(.*?)"/;
  const contactRegex = /id="cusInitialContactDate" value="(.*?)"/
  const screenRegex = /id="screenDate" value="(.*?)"/
  const interviewRegex = /id="interviewDate" value="(.*?)"/
 
  //
  for (var i = start; i<= end; i++){
    var profileId = i
    if (useIdArray) {
      profileId = idArray[i]
    }
    var preIntakeString = getProfileData(preIntakeUrl + profileId, authCookie)
    //fetchCount++
    
    if (!verifySize(preIntakeString)){
      if (preIntakeString.length < 2000) { authCookie = logIn() }
      skippedIds.push( [ profileId, preIntakeString.length, "/resTracker/editresident/profile/preintake/" + profileId ])
      Logger.log("scrapePreIntake: Profile " + profileId + " skipped")
      continue
    }
    // perform regex matches here
    var nameFirst = preIntakeString.match(fNameRegex)[1] || ""
    var nameMiddle = preIntakeString.match(mNameRegex)[1] || ""
    var nameLast = preIntakeString.match(lNameRegex)[1] || ""
    var birthDate = preIntakeString.match(birthRegex)[1] || ""
    var initialContactDate = preIntakeString.match(contactRegex)[1] || ""
    var screenDate = preIntakeString.match(screenRegex)[1] || ""
    var interviewDate = preIntakeString.match(interviewRegex)[1] || ""

    scrapedData.push( [profileId, nameFirst, nameMiddle, nameLast, birthDate, initialContactDate, screenDate, interviewDate ] )

  } //end main loop
  if (skippedIds.length > 0){
    logSkipped(skippedIds)
    Logger.log("scrapedPreIntake Complete. Skipped " + skippedIds.length + " profiles.  See Log tab for details ")
  } else {
    Logger.log("scrapedPreIntake Complete. No profiles skipped.")
  }
  return scrapedData
}

// -----------------------------------------------------------------------------------------------------------------
// 
//https://xxxxx/resTracker/editresident/profile/ataglance/976
//
// -----------------------------------------------------------------------------------------------------------------
function scrapeAtAGlance(start, end, idArray){
  if (start == null) { start = 1 }
  if (end == null) { end = 10}
  
  // if an array of id numbers is passed, scrape that list of ids instead of using start and end in order
  var useIdArray = false
  if ( idArray != null && idArray.length > 0) {
    start = 0
    end = idArray.length - 1
    useIdArray = true
  }
  
  var atAGlanceUrl = "https://xxxxx/resTracker/editresident/profile/ataglance/"
  var authCookie = logIn()
  var scrapedData = []
  var skippedIds = []

  const profileNumberRegex = /<a href="\/resTracker\/residents\/profile\/(.*?)"/;
  const parentRegex = /The Parent, Guardian, or Legal Custodian of <a href="\/resTracker\/residents\/profile\/(.*?)"\/>(.*?)<\/a>/g;
  const childRegex = /The Child or Legal Ward of <a href="\/resTracker\/residents\/profile\/(.*?)"\/>(.*?)<\/a>/g;
  //const spouseRegex = /The Spouse of <a href="\/resTracker\/residents\/profile\/(.*?)"\/>(.*?)<\/a>/g;
  const classFreshRegex = /id="classFresh" value="(.*?)"/;
  const classSoftRegex = /id="classSoft" value="(.*?)"/;
  const classJuniorRegex = /id="classJunior" value="(.*?)"/;
  const classSeniorRegex = /id="classSenior" value="(.*?)"/;
  const gradDateRegex = /id="gradDate" value="(.*?)"/;
  const transDateRegex = /id="transDate" value="(.*?)"/;

  //Main Loop
  for (var i = start; i<= end; i++){
    // if scraping in order, use i, if using the array, read the profile id from the list
    var profileId = i
    if (useIdArray) {
      profileId = idArray[i]
    }
    
    var atAGlanceString = getProfileData(atAGlanceUrl + profileId, authCookie)
    
    //Check for Profile doesn't exist or error
    if (!verifySize(atAGlanceString)){
      if (atAGlanceString.length < 2000) { authCookie = logIn() }
      skippedIds.push( [ profileId, atAGlanceString.length, "/resTracker/editresident/profile/ataglance/" + profileId ])
      Logger.log("scrapeAtAGlance: Profile " + profileId + " skipped")
      continue
    }

    // put regex matches here
    var classFreshman = atAGlanceString.match(classFreshRegex)[1] || ""
    var classSophomore = atAGlanceString.match(classSoftRegex)[1] || ""
    var classJunior = atAGlanceString.match(classJuniorRegex)[1] || ""
    var classSenior = atAGlanceString.match(classSeniorRegex)[1] || ""
    var gradDate = atAGlanceString.match(gradDateRegex)[1] || ""
    var transDate = atAGlanceString.match(transDateRegex)[1] || ""
    
    var parentOfIds = []
    var isParentOf = atAGlanceString.match(parentRegex) || ""
    for (var match in isParentOf) {
      var rawMatch = isParentOf[match]
      var idNum = rawMatch.match(profileNumberRegex)[1] || ""
      parentOfIds.push(idNum)
    }
    
    var childOfIds = []
    var isChildOf = atAGlanceString.match(childRegex) || ""
    for (var match in isChildOf) {
      var rawMatch = isChildOf[match]
      var idNum = rawMatch.match(profileNumberRegex)[1] || ""
      childOfIds.push(idNum)
    }

    scrapedData.push( [profileId, classFreshman, classSophomore, classJunior, classSenior, gradDate, transDate, parentOfIds.toString(), childOfIds.toString() ] )

  } // end loop
  if (skippedIds.length > 0){
    logSkipped(skippedIds)
    Logger.log("scrapeAtAGlance: Complete. Skipped " + skippedIds.length + " profiles.  See Log tab for details ")
  } else {
    Logger.log("scrapeAtAGlance: Complete. No profiles skipped.")
  }
  return scrapedData
}

// -----------------------------------------------------------------------------------------------------------------
// 
//"https://xxxxx/resTracker/visits/index/"
//regex match the entire visits table and feed it to parseVisitsTable for processing
//Uses 1 urlfetch per profile, daily limit is 20k
// 
// -----------------------------------------------------------------------------------------------------------------
function scrapeOpenClose (start, end, idArray) {
  // starting and ending profiles for testing this function from editor
  if (start == null) { start = 539 }
  if (end == null) { end = 540} 
  
  // if an array of id numbers is passed, scrape that list of ids instead of using start and end in order
  var useIdArray = false
  if ( idArray != null && idArray.length > 0) {
    start = 0
    end = idArray.length - 1
    useIdArray = true
  }
  
  var openCloseUrl = "https://xxxxx/resTracker/visits/index/"
                             
  var authCookie = logIn()
  var scrapedData = []
  var skippedIds = []
  
  const visitsTableRegex = /Action<\/th>\n\t\t<\/tr>\n\t<\/thead>\n\t<tbody>([\s\S]*?)<\/tbody>/g;

  //
  for (var i = start; i<= end; i++){
    var profileId = i
    if (useIdArray) {
      profileId = idArray[i]
    }
    var openCloseString = getProfileData(openCloseUrl + profileId, authCookie)
    
    if (!verifySize(openCloseString)){
      if (openCloseString.length < 2000) { authCookie = logIn() }
      skippedIds.push( [ profileId, openCloseString.length, "/resTracker/visits/index/" + profileId ])
      Logger.log("scrapeOpenClose: Profile " + profileId + " skipped")
      //continue
    }
    var visitsTableMatch = openCloseString.match(visitsTableRegex) || ""
    var parsedVisits = parseVisitsTable(visitsTableMatch[0], authCookie, profileId)
    
    for (var visit in parsedVisits){
      scrapedData.push(parsedVisits[visit])
    }

  } //end main loop

  if (skippedIds.length > 0){
    logSkipped(skippedIds)
    Logger.log("scrapeOpenClose: Complete. Skipped " + skippedIds.length + " profiles.  See Log tab for details ")
  } else {
    Logger.log("scrapeOpenClose: Complete. No profiles skipped.")
  }
  return scrapedData
}

// -----------------------------------------------------------------------------------------------------------------
//  Convert Visits Table match from string to table, uses <tr> tag to identify visit and extract the dates
// Return: Array of [visitStatus, startDate, endDate] arrays, if not visits found returns an empty visit
// -----------------------------------------------------------------------------------------------------------------
function parseVisitsTable( tableString, authCookie, profileId ) {
  var parsedArray = []
  const linkRegex = /\[&nbsp;<a href="(.*?)"/;
  
  // Check the received match string to make sure it is valid, and if it has visits or not.
  if (tableString == null) {
    Logger.log("parseVisitsTable: tableString is null " + tableString)
    return
  } else if ( tableString.indexOf("No visits found.") != -1 ) {
    parsedArray.push([ profileId, "None", "N/A", "N/A", "N/A", "N/A" ])
    //Logger.log("parseVisitsTable: " + " 0 visits found: " + parsedArray)
    return parsedArray
  // Table has dates, Convert it to an array and extract them
  }else {
    // Strip out  needless HTML and tabs
    var cleanString = tableString.replace("\n", "").replace(/\t/g, "").replace(/<td>/g, "").replace(/<\/td>/g , "")
     
    // Convert each \n of the string to an array to process line by line looking for tr
    var rawArray = cleanString.split("\n")

    for ( var row = 0; row < rawArray.length; row++){
      // Ignore blank elements
      if (rawArray[row] == "" || rawArray[row] == null) { continue }
      
      // Find a <tr> tag row which means there is a visit
      if ( rawArray[row].indexOf("<tr") != -1 ){
        
        // determine if visit is open, if so, label it as open and read the next 2 rows to get start and end date
        if (rawArray[row].indexOf("active") != -1) {
          var linkString = rawArray[row + 8]
          var visitLink = linkString.match(linkRegex)[1] || ""
          parsedArray.push([ profileId, "Open", rawArray[row + 1], rawArray[row + 2], "N/A", visitLink ])
          continue
        // if visit is closed, label it, read the dates, and scrape the visit endpoint to get the Stable status  
        } else {
          var linkString = rawArray[row + 8]
          var visitLink = linkString.match(linkRegex)[1] || ""
         // Logger.log("parseVisitsTable: visitLink sending to scrapeVisit:  " + visitLink)
          var stableStatus = scrapeVisit(visitLink, authCookie)
          parsedArray.push([ profileId, "Closed", rawArray[row + 1], rawArray[row + 2], stableStatus, visitLink ])
          continue
        }
        
      }// end if its a tr
    
    }//end looping raw Array
    //Logger.log("parseVisitArray: "+ parsedArray.length + " visits found: " + parsedArray)
    return parsedArray
  }// end if table has dates
   
}

// -----------------------------------------------------------------------------------------------------------------
// perm_housing_exit_0 is no, perm_housing_exit_1 is yes
// Returns status: Stable, Unstable, or Unknown
// -----------------------------------------------------------------------------------------------------------------
function scrapeVisit ( visitPath, auth ) {
  var baseUrl = "https://xxxxx"
  var visitUrl = baseUrl + visitPath
  var returnStatus = ""
  //var authCookie = logIn()
  
  const permHousingNoRegex = /id="perm_housing_exit_0" value="0"(.*?)\/>/;
  const permHousingYesRegex = /id="perm_housing_exit_1" value="1"(.*?)\/>/;
  
  var visitString = getProfileData(visitUrl, auth)
    
    if (!verifySize(visitString)){
      Logger.log("scrapeVisit: visit at: " + visitPath + " skipped")
      return
    }
  
  var noIsCheckedString = visitString.match(permHousingNoRegex)[1] || ""
  var yesIsCheckedString = visitString.match(permHousingYesRegex)[1] || ""
  
  //determine if yes or no (or neither) is selected and set the status
  if (noIsCheckedString.indexOf("checked") != -1 ){
    returnStatus = "Not Stable"
  } else if (yesIsCheckedString.indexOf("checked") != -1 ) {
    returnStatus = "Stable"
  } else {
    returnStatus = "Unknown"
  }
  //Logger.log("scrapeVisit: returnStatus: " + returnStatus + " for visit: " + visitPath)

  return returnStatus
}

// -----------------------------------------------------------------------------------------------------------------
// Strip out null rows and write all data to a spreadsheet
// -----------------------------------------------------------------------------------------------------------------
function writeArrayToSpreadsheet(arr, startingRecord, sheetName) {
  if (arr == null) {
    Logger.log("writeArrayToSpreadsheet: Cannot write to spreadsheet, Array is null")
    return
  }
  
  Logger.log("writeArrayToSpreadsheet:  Executing for ids beginning with: " + startingRecord)
  var cleanArray = arr.filter(stripNulls)
  function stripNulls (value){
    return value;
  }
  
  var rows = cleanArray.length
  var cols = cleanArray[0].length
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (sheetName == null ){
    sheetName = "DataTest"
  }
  
  var dataSheet = ss.getSheetByName(sheetName)
  var dataRange = dataSheet.getRange(dataSheet.getLastRow() + 1 , 1, rows, cols)
  //dataRange.clearContent()
  Logger.log("writeArrayToSpreadsheet: Writing array Starting Id: " + cleanArray[0][0] + " Ending Id: " + cleanArray[cleanArray.length - 1][0] )
  dataRange.setValues(cleanArray)

}

//-----------------------------------------------------------------------------------------------------------------------------------
//                           UTILITIES
//-----------------------------------------------------------------------------------------------------------------------------------

// -----------------------------------------------------------------------------------------------------------------
// Logs in to Mission Tracker, returns cookie string used to validate all future UrlFetchApp requests
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

// -----------------------------------------------------------------------------------------------------------------
// Returns HTML of website as a string, auth cookie required from successful login
// -----------------------------------------------------------------------------------------------------------------
function getProfileData(url, auth) { 
   var doc = UrlFetchApp.fetch(url,
                               {"headers" : {"Cookie" : auth} }
                              ).getContentText();
  return doc
}

// -----------------------------------------------------------------------------------------------------------------
// Return True if http String is over a certain size
// -----------------------------------------------------------------------------------------------------------------
function verifySize ( httpString ){
  if ( httpString == null ) {
    Logger.log("verifySize: Null Http Response")
    return false
  }else if ( httpString.length < 1500 ){
    //Login Screen has a length of 1277, if we get this it means auth failed
    Logger.log("verifySize: Login screen (length 1277) detected. TODO: Re-scrape this id.  Length returned: " + httpString.length)
    //dumpS(httpString)
    ///Logger.log(httpString)
     return false
  }else if ( httpString.length < 5000 ){
    // No profile found page has length of 4712, TODO: exclude profile numbers that return this from future scrape calls
    Logger.log("verifySize: No profile found for this id. (length 4712). Length returned: " + httpString.length)
     return false
  }else{
    return true
  }
}

// Clear the Data Tab leaving the headers
function clearDataSheet( sheetName ) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (sheetName == null){
    sheetName = "DataTest"
  }
  var sheet = ss.getSheetByName(sheetName);
  
  // clear out the matches and output sheets
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(3, 1, lastRow-1, sheet.getLastColumn() ).clearContent();
  }
}

// Sort the Data Tab by the ID column
function sortDataSheet(sheetName){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (sheetName == null){
    sheetName = "DataTest"
  }
  var dataSheet = ss.getSheetByName(sheetName)
  var dataRange = dataSheet.getRange(3 , 1, dataSheet.getLastRow(), dataSheet.getLastColumn())
  var dataValues = dataRange.getValues()
  dataValues.sort(sortFunction)
  dataRange.clearContent()
  dataRange.setValues(dataValues)
}


//sort used by Array.sort(sortFunction), sort by first column, column can be changed by changing x at a[x] and b[x]
function sortFunction(a, b) {
    if (a[0] === b[0]) {
        return 0;
    }
    else {
        return (a[0] < b[0]) ? -1 : 1;
    }
}

//write Logger to Log google document
//replace xxxxx with url of google doc
function logS () {
  var doc = DocumentApp.openByUrl("https://xxxxx")
  var logString = Logger.getLog()
  doc.getBody().appendParagraph(logString)
}

//replace xxxxx with url of google doc 2
function logD () {
  var doc = DocumentApp.openByUrl("https://xxxxx")
  var logString = Logger.getLog()
  doc.getBody().appendParagraph(logString)
}

//write list of skipped ids to log tab
function logSkipped (arr, startCol) {
  if (arr == null) {
    Logger.log("logS: null array recieved, unable to log")
    return
  }
  if (startCol == null){
    startCol = 1
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Log")
  var logRange = logSheet.getRange(logSheet.getLastRow() + 1 , startCol, arr.length, arr[0].length)
  logRange.setValues(arr)
}

//write content parameter to Dump google document
//replace xxxxx with url of google doc
function dumpS ( content ) {
  var doc = DocumentApp.openByUrl("https://xxxxx")
  doc.getBody().appendParagraph(content)
}

// Clear the Log Tab
function logSClear() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Log")
  logSheet.getRange(3, 1, logSheet.getLastRow() -1 , logSheet.getLastColumn()).clearContent()
}

//Generate a random number between 1 and max param
function randomInt ( max ) {
  return Math.ceil(Math.random() * max)
}