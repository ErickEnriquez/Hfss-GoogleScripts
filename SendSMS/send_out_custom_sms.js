///////////////////////////////////////////////////////////////////////////////////////
const SID = "AC9f2f8a075203b765ee7724fbf1c8bb17";
const token = "d53462cfa67afe893f96c6a7bad52cc7";
const companyPhone = "6026712206";
const turtlesImage =
  "https://hfss-website.s3-us-west-2.amazonaws.com/turtles.png";
///////////////////////////////////////////////////////////////////////////////////////

//======================================================================================================================================================
//Creates the Custom Menu that will display show sidebar
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .createMenu("Custom Menu")
    .addItem("Show sidebar", "showSidebar")
    .addToUi();
}
//======================================================================================================================================================
//calls Page.html to render the html as a sidebar
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("Page")
    .setTitle("Helper")
    .setWidth(800);
  SpreadsheetApp.getUi().showSidebar(html);
}
//======================================================================================================================================================

// Sets the value of A1 cell to value entered in the input field in the side bar!
function enterName(name) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  sheet.getRange("A1").setValue(name);
}
//======================================================================================================================================================

function sendSms(to, body) {
  var messages_url =
    "https://api.twilio.com/2010-04-01/Accounts/" + SID + "/Messages.json";

  var payload = {
    To: to,
    Body: body,
    From: companyPhone,
  };

  var options = {
    method: "post",
    payload: payload,
  };

  options.headers = {
    Authorization: "Basic " + Utilities.base64Encode(SID + ":" + token),
  };

  UrlFetchApp.fetch(messages_url, options);
}

//======================================================================================================================================================

//takes a the formatted data and passes it to send SMS function to send text out
function sendOutTexts() {
  let spreadsheet = SpreadsheetApp.getActive();
  let data = spreadsheet.getDataRange().getValues();
  let status = "";
  for (i = 1; i < data.length; i++) {
    let row = data[i];

    try {
      response_data = sendSms(String(row[2]), messageText);
      status = "sent";
    } catch (err) {
      Logger.log(err);
      status = "error";
    }
    spreadsheet.getRange("C" + (Number(i) + 1)).setValue(status);
  }
}
//======================================================================================================================================================
//Simple utility function to print out an an output
function printObject(output){
    SpreadsheetApp.getUi().alert('Hello, world' + output + ' IS WHAT IT SHOULD SAY');
}