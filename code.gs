/**
 * When there is a change to the calendar, searches for events that include "#meetingagenda"
 * in the decrisption.
 *
 */
function onCalendarChange() {
  // Gets recent events with the #meetingagenda tag
  const now = new Date();
  const dayOfWeek = now.getDay();
  
  // Check if it is Monday (dayOfWeek = 1) and if it is not a holiday (e.g., Christmas Day)
  if (dayOfWeek === 1) {
        // Calculate the start of the current week (Sunday)
        const startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dayOfWeek);
  
        // Calculate the end of the current week (Saturday)
        const endOfWeek = new Date(startOfWeek.getTime() + 6 * 24 * 60 * 60 * 1000);
        
    const events = CalendarApp.getEvents(
        startOfWeek,
        endOfWeek,
        {search: '#meetingagenda'},
    );


    // Loops through any events found
    for (i = 0; i < events.length; i++) {
      const event = events[i];

      // Confirms whether the event has the #meetingagenda tag
      let description = event.getDescription();
      if (description.search('#meetingagenda') == -1) continue;
    
      const eventDate = Utilities.formatDate(event.getStartTime(), "GMT", "MM-dd-yyyy");
        
        // Get the year from the event date
        const year = eventDate.split("-")[2];
        
        // Get the month number from the event date
        const monthNum = eventDate.split("-")[0];
        
        // Get the month name from the month number
        const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
        const month = monthNames[parseInt(monthNum, 10) - 1];


  /**
  * Checks to see if there is a folder for the year the event takes place. If not, it creates one.
  * Then checks to see if there is a folder for the month that the event takes place. If not, it creates one.
  *
  */
        // Get the root folder
        var rootFolder = DriveApp.getFolderById("FOLDER_ID_WHERE_FILE_WILL_BE_STORED_GOES_HERE");
        
        // Check if the year folder exists
        var yearFolder = rootFolder.getFoldersByName(year);
        if (yearFolder.hasNext()) {
          yearFolder = yearFolder.next();
        } else {
          // Create the year folder
          yearFolder = rootFolder.createFolder(year);
        }
        
        // Check if the month folder exists
        var monthFolder = yearFolder.getFoldersByName(month);
        if (monthFolder.hasNext()) {
          monthFolder = monthFolder.next();
        } else {
          // Create the month folder
          monthFolder = yearFolder.createFolder(month);
        }

// Only works with events created by the owner of this calendar
if (event.isOwnedByMe()) {
  // Creates a new sheet from the template file and renames it for this event's date
  const templateSheetId = "TEMPLATE_FILE_ID_GOES_HERE";
  const newSheet = DriveApp.getFileById(templateSheetId).makeCopy("Meeting Agenda " + eventDate, monthFolder);
  const newSheetId = newSheet.getId();
  console.log(newSheetId);

  // Replaces the hashtag in the event description with a link to the agenda sheet
  const agendaUrl = 'https://docs.google.com/spreadsheets/d/' + newSheetId;
  let description = event.getDescription();
  description = description.replace('#meetingagenda', '<a href=' + agendaUrl + '>Meeting Agenda</a>');
  event.setDescription(description);


  // Adds attendees as editors of the new sheet
  const editors = event.getGuestList().filter(guest => guest.getEmail() !== "").map(guest => guest.getEmail());
  Drive.Permissions.insert({
     role: 'writer',
     type: 'user',
     value: editors 
     }, 
     
     newSheetId, 
     { 
       'sendNotificationEmails': 'true',
       'emailMessage': "Hi everyone, Attached here is the agenda for this week's meeting. Please take a few minutes between now and then to note some projects and tasks which you worked on last week and some that you are tackling this week." });
  }
}
    return;
  }
}
