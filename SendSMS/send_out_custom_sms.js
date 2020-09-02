///////////////////////////////////////////////////////////////////////////////////////
const SID = "AC9f2f8a075203b765ee7724fbf1c8bb17";
const token = "d53462cfa67afe893f96c6a7bad52cc7";
const companyPhone = "6026712206";
const turtlesImage =
  "https://hfss-website.s3-us-west-2.amazonaws.com/turtles.png";
const MASTER_SHEET = "PASTE NAMES HERE";
const OUTPUT_SHEET = "RESULTS";
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
function sendOutTexts(res) {
  let spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(MASTER_SHEET), true);
  let data = spreadsheet.getDataRange().getValues();
  const phoneColumnIndex = getPhoneColumnIndex(res.phoneNumberColumn);
  let status = "";
  let output = [];
  let temp = {};
  let flag = true;
  /*SpreadsheetApp.getUi().alert(
    "HEADERS CHECKED " +
      res.headerRows +
      "\nMESSAGE TEXT " +
      res.message +
      "\nPHONE NUM COL " +
      res.phoneNumberColumn
  );*/
  for (let i = 0; i < data.length; i++) {
    if (res.headerRows == true && flag == true) {
      i++;
      flag = false;
    }
    let text = replaceSmartTags(res.message, data[i]);
    try {
      //responseData = sendSms(res.phoneNumberColumn, text)
      status = "Sent";
    } catch (err) {
      Logger.log(err);
      status = "Error";
    }
    temp.textStatus = status;
    temp.phoneNumber = data[phoneColumnIndex];
    output.push(temp);
  }
  SpreadsheetApp.getUi().alert(output.length + " " + phoneColumnIndex);

  writeResults(output);
}
//======================================================================================================================================================
//Simple utility function to print out an an output
function printObject(output) {
  SpreadsheetApp.getUi().alert(output);
}

//======================================================================================================================================================
//function to replace all of the smart tags with their string literals
function replaceSmartTags(response, data) {
  let text = response.replace(/\{a\}/gi, data[0]);
  text = text.replace(/\{b\}/gi, data[1]);
  text = text.replace(/\{c\}/gi, data[2]);
  text = text.replace(/\{d\}/gi, data[3]);
  text = text.replace(/\{e\}/gi, data[4]);
  return text;
}

//======================================================================================================================================================
//This function switches to results sheet and gives the results of the text
function writeResult(output) {
  let spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(OUTPUT_SHEET), true);
  spreadsheet.getDataRange().clearContent(); // clear out old content from spreadsheet
  const headerRow = "A1:B1";
  spreadsheet.getRange(headerRow).setValues([["Phone Number", "Text Status"]]);
  rowIndex = 2;
  for (let i in output) {
    let row = "A" + rowIndex + ":B" + rowIndex;
    spreadsheet
      .getRange(row)
      .setValues([[output[i].phoneNumber, output[i].textStatus]]);
    rowIndex++;
  }
}

//======================================================================================================================================================
//gets the index of the phone number column where input is the letters a-e
function getPhoneColumnIndex(input) {
  switch (input) {
    case "A":
      return 0;
    case "B":
      return 1;
    case "C":
      return 2;
    case "D":
      return 3;
    case "E":
      return 4;
  }
}
