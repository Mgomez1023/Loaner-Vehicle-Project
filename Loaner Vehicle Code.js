// Change to true to enable email notifications
var emailNotification = false;
var emailAddress = "Change_to_your_Email";

// DO NOT EDIT THESE NEXT PARAMS
var isNewSheet = false;
var receivedData = [];

/**
 * this is a function that fires when the webapp receives a GET request
 * Not used but required.
 */
function doGet(e) {
  return ContentService.createTextOutput(
        JSON.stringify({
          success: true,
          data: { message: "Submission successful!" }
        })
      ).setMimeType(ContentService.MimeType.JSON);
}

// Webhook Receiver - triggered with form webhook to published App URL.

function doPost(e) {
  try {
    var params = JSON.stringify(e.parameter);
    params = JSON.parse(params);

    console.log("Received parameters: " + params); // Logging received parameters

    var formId = params["form_id"];
    var specificFormId = "63efc8f"; 

    console.log("Form ID: " + formId);

    if (formId === specificFormId) {
      var email = params["form_fields[email]"];
      console.log('Email: ' + email);
      var date = params["date"] || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var time = params["time"] || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');

      var emailFoundRow = searchForEmail(email);
      console.log('Email found at row: ' + emailFoundRow);

      if (emailFoundRow == -1) {
        console.log('Error: Email not found');
        return ContentService.createTextOutput("Submission failed.").setMimeType(ContentService.MimeType.TEXT);

      }
      else {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        var timeColumn = getColumnIndexByHeader("Time Out");
        var timeCellValue = sheet.getRange(emailFoundRow, timeColumn).getValue();
        console.log('Time Out column value: ' + timeCellValue);

        if (timeCellValue != "") {
          console.log('Error: Time Out column is not empty');
          return ContentService.createTextOutput(
            JSON.stringify({
              "success": false,
              "data": {
                "message": "Time Column not empty.",
                "errors": [],
                "data": []
              }
            }))
            .setMimeType(ContentService.MimeType.JSON);
        } else {
          var dateColumn = getColumnIndexByHeader("Date Out");
          sheet.getRange(emailFoundRow, dateColumn).setValue(date);
          sheet.getRange(emailFoundRow, timeColumn).setValue(time);

          console.log('Success: Data updated successfully fr');
          // Verification after attempting to update the cells
          var updatedDateValue = sheet.getRange(emailFoundRow, dateColumn).getValue();
          var updatedTimeValue = sheet.getRange(emailFoundRow, timeColumn).getValue();

          // If the cells were updated successfully
          const response = {
            success: true,
            data: {
              message: "Submission Successful"
            }
          };
          
          return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);

        }
      }
    } else {
      insertToSheet(params);
    } 
  } catch (error) {
    console.log("Error occurred: " + error.message); // Log the error message
    return ContentService.createTextOutput(
      JSON.stringify({
        "success": false,
        "data": {
          "message": "Error: An unexpected error occurred - " + error.message,
          "errors": [],
          "data": []
        }
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function searchForEmail(email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();


  var emailColumnIndex = getColumnIndexByHeader("Enter Email");  // Replace with your actual email column name
  var timeOutColumnIndex = getColumnIndexByHeader("Time Out");   // Replace with your actual "Time Out" column name

  // Loop through the rows and search for the email and empty "Time Out" column
  for (var i = 1; i < values.length; i++) {
    if (values[i][emailColumnIndex - 1] === email && !values[i][timeOutColumnIndex - 1]) {
      return i + 1;  // Return the row number (1-based index) if the email matches and "Time Out" is empty
    }
  }

  return -1;  // Email not found with empty "Time Out" column
}

function searchForEmailAndEmptyTimeOut(email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var emailColumnIndex = getColumnIndexByHeader("Enter Email"); // Replace with your email column name
  var timeOutColumnIndex = getColumnIndexByHeader("Time Out"); // Replace with your time out column name

  // Loop through the rows and search for the email and check if "Time Out" is empty
  for (var i = 1; i < values.length; i++) {
    if (values[i][emailColumnIndex - 1] == email && values[i][timeOutColumnIndex - 1]) {
      return i + 1; // Return the row number (1-based index)
    }
  }

  return -1; // Email not found or "Time Out" is not empty
}

function getColumnIndexByHeader(headerName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] == headerName) {
      return i + 1; // Return the column index (1-based index)
    }
  }
  return -1; // Header not found
}

// Flattens a nested object for easier use with a spreadsheet
function flattenObject(ob) {
  var toReturn = {};
  for (var i in ob) {
    if (!ob.hasOwnProperty(i)) continue;
    if (typeof ob[i] == "object") {
      var flatObject = flattenObject(ob[i]);
      for (var x in flatObject) {
        if (!flatObject.hasOwnProperty(x)) continue;
        toReturn[i + "." + x] = flatObject[x];
      }
    } else {
      toReturn[i] = ob[i];
    }
  }
  return toReturn;
}

// normalize headers
function getHeaders(formSheet, keys) {
  var headers = [];

  // retrieve existing headers
  if (!isNewSheet) {
    headers = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
  }

  // add any additional headers
  var newHeaders = [];
  newHeaders = keys.filter(function (k) {
    return headers.indexOf(k) > -1 ? false : k;
  });

  newHeaders.forEach(function (h) {
    headers.push(h);
  });
  return headers;
}

// normalize values
function getValues(headers, flat) {
  var values = [];
  // push values based on headers
  headers.forEach(function (h) {
    values.push(flat[h]);
  });
  return values;
}

// Insert headers
function setHeaders(sheet, values) {
  var headerRow = sheet.getRange(1, 1, 1, values.length);
  headerRow.setValues([values]);
  headerRow.setFontWeight("bold").setHorizontalAlignment("center");
}

// Insert Data into Sheet
function setValues(sheet, values) {
  var lastRow = Math.max(sheet.getLastRow(), 1);
  sheet.insertRowAfter(lastRow);
  sheet.getRange(lastRow + 1, 1, 1, values.length).setValues([values]).setFontWeight("normal").setHorizontalAlignment("center");
}

// Find or create sheet for form
function getFormSheet(formName) {
  var formSheet;
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet();

  // create sheet if needed
  if (activeSheet.getSheetByName(formName) == null) {
    formSheet = activeSheet.insertSheet();
    formSheet.setName(formName);
    isNewSheet = true;
  }
  return activeSheet.getSheetByName(formName);
}

// magic function where it all happens
function insertToSheet(data) {
  var flat = flattenObject(data);
  var keys = Object.keys(flat);
  var formName = data["form_name"];
  var formSheet = getFormSheet(formName);
  var headers = getHeaders(formSheet, keys);
  var values = getValues(headers, flat);
  setHeaders(formSheet, headers);
  setValues(formSheet, values);

  if (emailNotification) {
    sendNotification(data, getSheetURL());
  }
}

function getSheetURL() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  return spreadsheet.getUrl();
}

function sendNotification(data, url) {
  var subject = "A new Elementor Pro Forms submission has been inserted to your sheet";
  var message = "A new submission has been received via " + data["form_name"] + " form and inserted into your Google sheet at: " + url;
  MailApp.sendEmail(emailAddress, subject, message, {
    name: "Automatic Emailer Script",
  });
}