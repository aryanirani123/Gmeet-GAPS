// This code is wriiten in Google Apps Script(JavaScript)
// This code creates Calendar Events with Google Meet
// All the details regarding the events is stored in a Google Sheet
// The code takes the details from the sheet and creates events in Google Calendar

function createNewEventWithMeet() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Calendar_Events");
  var last_row = sheet.getLastRow();
  var data = sheet.getRange("A2:E" + last_row).getValues();
  var cal = CalendarApp.getCalendarById("aryanirani123@gmail.com");

  for(var i = 0;i< data.length;i++){

    var event_name = data[i][0];
    var start_time = data[i][1];
    var end_time = data[i][2];
    var event_description = data[i][3];
    var attendees_event = data[i][4];

  const gmt = "+05:30";
  const calendarId = "aryanirani123@gmail.com"; // Consider making this dynamic or a script property

  // 1. Create Google Meet link using Meet API
  var googleMeet_Link = "";
  try {
    const meetApiUrl = "https://meet.googleapis.com/v2/spaces";
    const token = ScriptApp.getOAuthToken();
    const options = {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + token,
      },
      payload: JSON.stringify({}), // Empty body for default space settings
      muteHttpExceptions: true,
    };
    const meetResponse = UrlFetchApp.fetch(meetApiUrl, options);
    const responseCode = meetResponse.getResponseCode();
    const responseBody = meetResponse.getContentText();

    if (responseCode === 200) {
      const meetData = JSON.parse(responseBody);
      googleMeet_Link = meetData.meetingUri;
      if (!googleMeet_Link) {
         console.error("meetingUri not found in Meet API response: " + responseBody);
         // Potentially fall back to old method or skip event creation if Meet link is crucial
         googleMeet_Link = "Error creating Meet link - see logs"; // Placeholder
      }
    } else {
      console.error("Error creating Google Meet link. Response Code: " + responseCode + ". Response Body: " + responseBody);
      // Handle error, e.g., skip adding meet or log error
      googleMeet_Link = "Error creating Meet link - see logs"; // Placeholder
    }
  } catch (e) {
    console.error("Exception creating Google Meet link: " + e);
    googleMeet_Link = "Exception creating Meet link - see logs"; // Placeholder
  }

  // 2. Create Calendar event resource with the obtained Meet link
  const resource = {
    start: { dateTime: start_time + gmt }, // Ensure start_time is in correct ISO format portion
    end: { dateTime: end_time + gmt },     // Ensure end_time is in correct ISO format portion
    attendees: [{ email: attendees_event }],
    conferenceData: {
      entryPoints: [{
        entryPointType: "video",
        uri: googleMeet_Link,
        label: "Google Meet"
      }]
    },
    summary: event_name,
    description: event_description + "\n\nJoin Meeting: " + googleMeet_Link,
  };

  try {
    const res = Calendar.Events.insert(resource, calendarId, {
      conferenceDataVersion: 1, // May still be relevant for entryPoints
    });
    console.log("Event created: " + res.htmlLink + " with Meet link: " + googleMeet_Link);
  } catch (e) {
    console.error("Error creating calendar event: " + e + ". Event details: " + JSON.stringify(resource));
  }
  }
}
