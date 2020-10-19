///////////////////////////////////////////////////////////////////////////////////////
const SID = process.env.SID;
const token = process.env.token;
const companyPhone = process.env.companyPhone;
const MASTER_SHEET = "PASTE NAMES HERE";
const RESPONSES_SHEET = "RESPONSES";
const numberToRetrieve = 500;
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
  let flag = true;
  spreadsheet.getRange("H1").setValue("TEXT STATUS");
  for (let i = 0; i < data.length; i++) {
    if (res.headerRows == true && flag == true) {
      i++;
      flag = false;
    }
    let text = replaceSmartTags(res.message, data[i]);
    try {
      responseData = sendSms(data[i][phoneColumnIndex], text);
      status = "Sent";
    } catch (err) {
      Logger.log(err);
      status = "Error";
    }
    spreadsheet.getRange("H" + (Number(i) + 1)).setValue(status);
  }
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
//======================================================================================================================================================
//call this to retrieve all 200 messages
const receiveResponses = () => {
  let options = {
    method: "get",
  };
  options.headers = {
    Authorization: "Basic " + Utilities.base64Encode(SID + ":" + token),
  };
  let url ="https://api.twilio.com/2010-04-01/Accounts/" + SID + "/Messages.json?To=" + companyPhone + "&PageSize=" + numberToRetrieve;
  let response = UrlFetchApp.fetch(url, options);
  // -------------------------------------------
  // Parse the JSON data and put it into the spreadsheet's active page.
  // Documentation: https://www.twilio.com/docs/api/rest/response
  let spreadsheet = SpreadsheetApp.getActiveSheet();
  let rowIndex = 2;
  var dataAll = JSON.parse(response.getContentText());
  for (i = 0; i < dataAll.messages.length; i++) {
    let rowDate = dataAll.messages[i].date_sent;
    let dateObj = new Date(rowDate);
    if (isNaN(dateObj.valueOf())) {
      dateObj = "NOT A VALID DATE";
    } else {
      //dateObj.setHours(dateObj.getHours() + hoursOffset);
    }
    let row = "A" + rowIndex + ":E" + rowIndex;
    spreadsheet
      .getRange(row)
      .setValues([
        [
          dateObj,
          dateObj.toLocaleTimeString(),
          dataAll.messages[i].to,
          dataAll.messages[i].from,
          dataAll.messages[i].body,
        ],
      ]);
    rowIndex++;
  }
}