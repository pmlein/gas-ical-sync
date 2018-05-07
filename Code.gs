/* --------------- HOW TO INSTALL ---------------
*
* 1) Click in the menu "File" > "Make a copy..." and make a copy to your Google Drive
* 2) Changes lines 11-19 to be the settings that you want to use
* 3) Click in the menu "Run" > "Run function" > "Main" and authorize the program
*
*/

// --------------- SETTINGS ---------------
//var sourceCalendarURL = "https://drive.google.com/open?id=1crjIhg9AaEQHpdD17vbJ6h3l6jcs6R4r"; //The ics/ical url that you want to get events from
var sourceCalendarName = 'events.ics';
var targetCalendarName = "Luistelu"; //The name of the Google Calendar you want to add events to
//var howFrequent = 15; //What interval (minutes) to run this script on to check for new events
var addEventsToCalendar = true; //If you turn this to "false", you can check the log (View > Logs) to make sure your events are being read correctly before turning this on
var addAlerts = true; //Whether to add the ics/ical alerts as notifications on the Google Calendar events
var descriptionAsTitles = false; //Whether to use the ics/ical descriptions as titles (true) or to use the normal titles as titles (false)
var defaultDuration = 60; //Default duration (in minutes) in case the event is missing an end specification in the ICS/ICAL file
var emailWhenAdded = false; //Will email you when an event is added to your calendar
var email = ""; //OPTIONAL: If "emailWhenAdded" is set to true, you will need to provide your email
// ----------------------------------------

/* --------------- MISCELLANEOUS ----------
*
* This program gets ICalendar file from Google Drive as an input and syncs events to given Google Calendar
* The program for syncing was created by Derek Antrican 
* https://github.com/derekantrican/Google-Apps-Script-Library/tree/master/ICS-ICAL%20Sync
*/


function main() {  
  deleteLuisteluEvents() // Delete old events from Calendar "Luistelu"
  //Get events from events.ics in Google Drive
  var response = getEventsFile();
  response = response.getDataAsString().split("\n");
  
  //Get target calendar information
  var targetCalendar = CalendarApp.getCalendarsByName(targetCalendarName)[0];
  
  //------------------------ Error checking ------------------------
  if(response[0] == "That calendar does not exist.")
    throw "[ERROR] Incorrect ics/ical URL";
  
  if(targetCalendar == null)
    throw "[ERROR] Could not find calendar of name \"" + targetCalendarName + "\"";
  
  if (emailWhenAdded && email == "")
    throw "[ERROR] \"emailWhenAdded\" is set to true, but no email is defined";
  //----------------------------------------------------------------
  
  //------------------------ Parse events --------------------------
  //https://en.wikipedia.org/wiki/ICalendar#Technical_specifications
  //https://tools.ietf.org/html/rfc5545
  //https://www.kanzaki.com/docs/ical
  
  var parsingEvent = false;
  var parsingNotification = false;
  var currentEvent;
  var events = [];
  Logger.log('LENGTH: '+ response.length);
  for (var i = 0; i < response.length; i++){  
    if (response[i] == "BEGIN:VEVENT"){
      parsingEvent = true;
      currentEvent = new Event();
    }
    else if (response[i] == "END:VEVENT"){
      if (currentEvent.endTime == null)
        currentEvent.endTime = new Date(currentEvent.startTime.getTime() + defaultDuration * 60 * 1000);
      
      parsingEvent = false;
      events[events.length] = currentEvent;
    }
    else if (response[i] == "BEGIN:VALARM")
      parsingNotification = true;
    else if (response[i] == "END:VALARM")
      parsingNotification = false;
    else if (parsingNotification){
      if (addAlerts){
        if (response[i].includes("TRIGGER"))
          currentEvent.reminderTimes[currentEvent.reminderTimes.length++] = ParseNotificationTime(response[i].split("TRIGGER:")[1]);
      }
    }
    else if (parsingEvent){
      if (response[i].includes("SUMMARY") && !descriptionAsTitles)
        currentEvent.title = response[i].split("SUMMARY:")[1];
        
      if (response[i].includes("DESCRIPTION") && descriptionAsTitles)
        currentEvent.title = response[i].split("DESCRIPTION:")[1];
      else if (response[i].includes("DESCRIPTION"))
        currentEvent.description = response[i].split("DESCRIPTION:")[1];   
    
      if (response[i].includes("DTSTART"))
        currentEvent.startTime = Moment.moment(GetUTCTime(response[i].split("DTSTART")[1]), "YYYYMMDDTHHmmss").toDate();
        
      if (response[i].includes("DTEND"))
        currentEvent.endTime = Moment.moment(GetUTCTime(response[i].split("DTEND")[1]), "YYYYMMDDTHHmmss").toDate();
        
      if (response[i].includes("LOCATION"))
        currentEvent.location = response[i].split("LOCATION:")[1];
        
      if (response[i].includes("UID"))
        currentEvent.id = response[i].split("UID:")[1];
    }
  }
  //----------------------------------------------------------------
  
  //------------------------ Check results -------------------------
  Logger.log("# of events: " + events.length);
  for (var i = 0; i < events.length; i++){
    Logger.log("Title: " + events[i].title);
    Logger.log("Id: " + events[i].id);
    Logger.log("Description: " + events[i].description);
    Logger.log("Start: " + events[i].startTime);
    Logger.log("End: " + events[i].endTime);
    
    for (var j = 0; j < events[i].reminderTimes.length; j++)
      Logger.log("Reminder: " + events[i].reminderTimes[j] + " seconds before");
    
    Logger.log("");
  }
  //----------------------------------------------------------------
  
  //------------------------ Add events to calendar ----------------
  if (addEventsToCalendar){
    for (var i = 0; i < events.length; i++){
      if (!EventExists(targetCalendar, events[i])){
        var resultEvent = targetCalendar.createEvent(events[i].title, events[i].startTime, events[i].endTime, {location : events[i].location, description : events[i].description + "\n\n" + events[i].id});
        
        for (var j = 0; j < events[i].reminderTimes.length; j++)
          resultEvent.addPopupReminder(events[i].reminderTimes[j] / 60);
          
        if (emailWhenAdded)
          GmailApp.sendEmail(email, "New Event Added", "New event added to calendar \"" + targetCalendarName + "\" at " + events[i].startTime);
      }
    }
  }
  //----------------------------------------------------------------
}

function ParseNotificationTime(notificationString){
  //https://www.kanzaki.com/docs/ical/duration-t.html
  var reminderTime = 0;
  
  //We will assume all notifications are BEFORE the event
  if (notificationString[0] == "+" || notificationString[0] == "-")
    notificationString = notificationString.substr(1);
    
  notificationString = notificationString.substr(1); //Remove "P" character
  
  var secondMatch = RegExp("\\d+S", "g").exec(notificationString);
  var minuteMatch = RegExp("\\d+M", "g").exec(notificationString);
  var hourMatch = RegExp("\\d+H", "g").exec(notificationString);
  var dayMatch = RegExp("\\d+D", "g").exec(notificationString);
  var weekMatch = RegExp("\\d+W", "g").exec(notificationString);
  
  if (weekMatch != null){
    reminderTime += parseInt(weekMatch[0].slice(0, -1)) & 7 * 24 * 60 * 60; //Remove the "W" off the end
    
    return reminderTime; //Return the notification time in seconds
  }
  else{
    if (secondMatch != null)
      reminderTime += parseInt(secondMatch[0].slice(0, -1)); //Remove the "S" off the end
      
    if (minuteMatch != null)
      reminderTime += parseInt(minuteMatch[0].slice(0, -1)) * 60; //Remove the "M" off the end
      
    if (hourMatch != null)
      reminderTime += parseInt(hourMatch[0].slice(0, -1)) * 60 * 60; //Remove the "H" off the end
      
    if (dayMatch != null)
      reminderTime += parseInt(dayMatch[0].slice(0, -1)) * 24 * 60 * 60; //Remove the "D" off the end
      
    return reminderTime; //Return the notification time in seconds
  }
}

/* 
 * File will be read from Google Drive by name (fileNameToGet)
 */
function getEventsFile() {
  var allFilesInFolder,cntFiles,docContent,fileNameToGet,fldr,
    thisFile,whatFldrIdToUse;
 
  whatFldrIdToUse = 'root';
  fileNameToGet = sourceCalendarName;
  fldr = DriveApp.getFolderById(whatFldrIdToUse);
  allFilesInFolder = fldr.getFilesByName(fileNameToGet);
  Logger.log('allFilesInFolder: ' + allFilesInFolder);

  if (allFilesInFolder.hasNext() === false) {
    //If no file is found, the user gave a non-existent file name
    return false;
  }
  cntFiles = 0;
//Even if it's only one file, must iterate a while loop in order to access the file.
//Google drive will allow multiple files of the same name.
  while (allFilesInFolder.hasNext()) {
    thisFile = allFilesInFolder.next();
    cntFiles = cntFiles + 1;
    Logger.log('File Count: ' + cntFiles);
    
    docContent = thisFile.getAs('text/plain');
    // 
    Logger.log('docContent : ' + docContent );
  }
  if (cntFiles === 0)
     docContent = "That calendar does not exist."; 
  return docContent;
}


function EventExists(calendar, event){
  var events = calendar.getEvents(event.startTime, event.endTime, {search : event.id});
  
  return events.length > 0;
}

function GetUTCTime(parameter){
  parameter = parameter.substr(1); //Remove leading ; or : character
  if (parameter.includes("TZID")){
    var tzid = parameter.split("TZID=")[1].split(":")[0];
    var time = parameter.split(":")[1];
    return Moment.moment.tz(time,tzid).tz("Etc/UTC").format("YYYYMMDDTHHmmss") + "Z";    
  }
  else
    return parameter;
}

/*
 * Deletes old events from target calendar before new events will be added
 */ 
function deleteLuisteluEvents() {

    var fromDate = new Date(2015,0,1,0,0,0);
    var toDate = new Date(2022,0,4,0,0,0);
    var calendarName = targetCalendarName;

    // delete from Jan 1, 2015 to end of Jan 4, 2022 (for month 0 = Jan, 1 = Feb...)
  
  var calendar = CalendarApp.getCalendarsByName(calendarName)[0];
  var events = calendar.getEvents(fromDate, toDate);  
  if (events.length > 0) {
    for (i = 0; i < events.length; i++) {
      var event = events[i];
      Logger.log( event.getTitle());
      event.deleteEvent();
    }
  } else {
    Logger.log('No events found.');
  }
}