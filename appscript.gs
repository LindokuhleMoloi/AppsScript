const FREE_SLOTS_CALENDAR_ID = "***********************"
const BOOKED_SLOTS_CALENDAR_ID = "*************@group.calendar.google.com"

function moveEventToBookedCalendar(e) {
  Logger.log(e.namedValues["Boardroom (Responses)"])
  var startingTime = new Date(e.namedValues["Boardroom (Responses)"])
  const availableCalendar = CalendarApp.getCalendarById(FREE_SLOTS_CALENDAR_ID)
  const bookedCalendar = CalendarApp.getCalendarById("BOOKED_SLOTS_CALENDAR_ID")
  var endingTime = new Date(startingTime.getTime() + 1000 * 60 * 60)
  var events = availableCalendar.getEvents(startingTime, endingTime)
  events.forEach(event => {
    event.deleteEvent()
    bookedCalendar.createEvent("Booked Slot", startingTime, endingTime)


  })
  
}

function deleteAllAvailableSlots() {
  const calendar = CalendarApp.getCalendarById(FREE_SLOTS_CALENDAR_ID)
  const events = calendar.getEvents(new Date(2022, 8, 1), new Date(2023, 8, 30))
  events.forEach(event => {
    event.deleteEvent();
  })
  bringEvents();
}


function deleteAllBookedSlots() {
  const calendar = CalendarApp.getCalendarById(BOOKED_SLOTS_CALENDAR_ID)
  const events = calendar.getEvents(new Date(2022, 8, 1), new Date(2023, 8, 30))
  events.forEach(event => {
    event.deleteEvent();
  })
  bringEvents();
}

function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu("Booking Functions")

  /*var ui = SpreadsheetApp.getUi()
  var menu = ui.createMenu("Booking Functions")*/
  menu.addItem("Delete All Available Slots", "deleteAllAvailableSlots")
    .addItem("Delete All Booked Slots", "deleteAllBookedSlots")
    .addItem("Update Available Slots", "bringEvents")
    .addItem("Update Bookings", "fillOutdates")
    .addToUi();

}

function bringEvents() {
  const ws = SpreadsheetApp.getActiveSpreadsheet();
  const optionsSheet = ws.getSheetByName("Options")
  const availablecalendar = CalendarApp.getCalendarById(FREE_SLOTS_CALENDAR_ID)
  var events = availablecalendar.getEvents(new Date(2022, 8, 13), new Date(2024, 8, 30))
  //Logger.log(events)
  //optionsSheet.getRange(2, 1, lastRow(1), 2).clearContent()
  var row =0;
  events.forEach((event ,index) => {
    optionsSheet.getRange(index+2 + 1, 1, 1, 2).setValues([[event.getStartTime(), event.getId()]])
    row = index+2
  })
  //fillOutdates()
 const bookedcalendar = CalendarApp.getCalendarById(BOOKED_SLOTS_CALENDAR_ID)
 events = bookedcalendar.getEvents(new Date(2022, 8, 13), new Date(2022, 8, 30))
  //Logger.log(events)
  //optionsSheet.getRange(2, 1, lastRow(1), 2).clearContent()
  events.forEach((event ,index) => {
    if(row ==0) row=index+2
    optionsSheet.getRange(row , 1, 1, 2).setValues([[event.getStartTime(), event.getId()]])
    row++
  })
  fillOutdates()
}

function fillOutdates() {
  const ws = SpreadsheetApp.getActiveSpreadsheet();
  const optionsSheet = ws.getSheetByName("Options")
  const dates = optionsSheet.getRange(2, 1, optionsSheet.getLastRow() - 1, 4).getDisplayValues();
  const dateList = dates.filter(row => row[3] == "FREE").map(row => row[0])
  //Logger.log(dates)

  const form = FormApp.openById("*******")
  const dateQuestion = form.getItemById("1250237619")
  if (dateList.length == 0) {
    dateQuestion.asListItem().setChoiceValues(["No Dates Available"])
  } else {
    dateQuestion.asListItem().setChoiceValues(dateList)
  }







}
function findOutIdofQuestions() {
  const form = FormApp.openById("***************")
  const questions = form.getItems();
  questions.forEach(question => {
    const id = question.getId();
    const name = question.getTitle();
    Logger.log("name: " + name + " id:" + id)
  })

}

function lastRow(col) {
  const ws = SpreadsheetApp.getActiveSpreadsheet();
  const ss = ws.getActiveSheet();
  const lastRow = ss.getMaxRows()
  const range = ss.getRange(1, col, lastRow).getValues();
  for (i = lastRow - 1; i >= 0; i--) {
    //Logger.log(i)
    if (range[i][0]) {
      return i + 1
    }
  }

}


/*
Calendar ID
****************
*/

//[[Tue Sep 06 00:00:00 GMT+02:00 2022], [Wed Sep 07 00:00:00 GMT+02:00 2022], [Thu Sep 08 00:00:00 GMT+02:00 2022], [Fri Sep 09 00:00:00 GMT+02:00 2022], [Sat Sep 10 00:00:00 GMT+02:00 2022], [Sun Sep 11 00:00:00 GMT+02:00 2022], [Mon Sep 12 00:00:00 GMT+02:00 2022], [Tue Sep 13 00:00:00 GMT+02:00 2022]]


/*
1:06:20 PM	Info	name: Name id:1219493470
1:06:20 PM	Info	name: Which Boardroom? id:1336875801
1:06:20 PM	Info	name: Date id:1250237619
*/
