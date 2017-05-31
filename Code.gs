/**
 * A special function that inserts a custom menu when the spreadsheet opens.
 */
function onOpen() {
  var menu = [{name: 'Set up conference', functionName: 'setUpConference_'}];
  SpreadsheetApp.getActive().addMenu('Conference', menu);
}

/**
 * A set-up function that uses the conference data in the spreadsheet to create
 * Google Calendar events.
 */
function setUpConference_() {
  if (ScriptProperties.getProperty('calId')) {
    Browser.msgBox('Your conference is already set up. Look in Google Calendar!');
  }
  
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheets()[0];
  var range = sheet.getDataRange();
  var values = range.getValues();
  setUpCalendar_(values, range);
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit()
      .create();
  ss.removeMenu('Conference');
}

/**
 * Creates a Google Calendar with events for each conference session in the
 * spreadsheet, then writes the event IDs to the spreadsheet for future use.
 *
 * @param {String[][]} values Cell values for the spreadsheet range.
 * @param {Range} range A spreadsheet range that contains conference data.
 */
function setUpCalendar_(values, range) {
  var cal = CalendarApp.getCalendarsByName('Notifications');
  if(cal == null)
    cal = CalendarApp.createCalendar('Notifications');
  
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var title = session[2][3];
    //var titlelist = title.join(' ');
    var time = session[4];
    
    //Separate time from single cell
    for(var j=0; j < time.length; j++) {
      
      //get first time
      var time1 = "";
      time1 += time[j];
      
      //get second time
      if (time[j] == '-'){
        var time2 = "";
        time[j]=time[j+1];
        time2 += time[j];
      }
    
    }
    
    var start = joinDateAndTime_(session[6], time1);
    var end = joinDateAndTime_(session[7], time2);
    var options = {location: session[2], sendInvites: false};
    var event = cal.createEvent(titlelist, start, end, options)
        .setGuestsCanSeeGuests(false);
  }
  range.setValues(values);

  // Store the ID for the Calendar, which is needed to retrieve events by ID.
  ScriptProperties.setProperty('calId', cal.getId());
}

/**
 * Creates a single Date object from separate date and time cells.
 *
 * @param {Date} date A Date object from which to extract the date.
 * @param {Date} time A Date object from which to extract the time.
 * @return {Date} A Date object representing the combined date and time.
 */
function joinDateAndTime_(date, time) {
  date = new Date(date);
  date.setHours(time.getHours());
  date.setMinutes(time.getMinutes());
  return date;
}
