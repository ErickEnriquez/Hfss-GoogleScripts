///////////////////////////////////////////////////////
const SID = "AC9f2f8a075203b765ee7724fbf1c8bb17";
const token = "d53462cfa67afe893f96c6a7bad52cc7";
const companyPhone = "6026712206";
const formatSheet = "Format for text all";
const workingSheet = "WORKING sheet GMs";
const turtlesImage =
  "https://hfss-website.s3-us-west-2.amazonaws.com/turtles.png";
///////////////////////////////////////////////////////

// CONSTANT VALUES FOR COLUMNS TO MAKE CHANGES TO SCRIPT EASIER
const FIRST_NAME_COL = 0;
const LAST_NAME_COL = 1;
const WORKING_COLUMN = 2;
const PHONE_COL = 3;
const MON = 10; // testing value 9, correct = 10
const TUE = 12; // testing value 11, correct = 12
const WED = 14; // testing value 13, correct = 14
const THU = 16; // testing value 15, correct = 16
const FRI = 18; // testing value 17, correct = 18
const SAT = 20; // testing value 19, correct = 20
const SUN = 22; // testing value 21, correct = 22
const PHX = 4;
const PEORIA = 5;
const RV = 6;
const GY = 7;

const TOTAL_SCHEDULE_COL = 10; // column of total schedule for a person
///////////////////////////////////////////////////////

//==================================================================================================================================

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
//==================================================================================================================================

function copyToFormatted() {
  let spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(workingSheet), true);

  let data = spreadsheet.getDataRange().getValues();

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(formatSheet));
  spreadsheet.getDataRange().clearContent(); // clear out old content from spreadsheet

  const headerRow = "A1:P1";
  spreadsheet
    .getRange(headerRow)
    .setValues([
      [
        "First Name",
        "Last Name",
        "Phone",
        " Mon Sched",
        "Tue Sched",
        "Wed Sched",
        "Thu Sched",
        "Fri Sched",
        "Sat Sched",
        "Sun Sched",
        "Total",
        "PHX",
        "PEORIA",
        "RV",
        "GY",
        "Text Status",
      ],
    ]);
  spreadsheet.getActiveSheet().setColumnWidths(1, 3, 120);
  spreadsheet.getActiveSheet().setColumnWidths(4, 13, 70);
  spreadsheet.getActiveSheet().setColumnWidths(11, 1, 130);
  spreadsheet
    .getRange(headerRow)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setBackground("#dddddd").setTextRotation(0);
  spreadsheet.getRange("D:J").setBackground("#dddddd");
  spreadsheet.getRange("K:K").setBackground("#024278").setFontColor("#FFFFFF");
  let row = 1;
  for (i = 1; i < data.length; i++) {
    if (checkIfWorking(data[i]) == true) {
      Logger.log(data[i][PHX]);
      row++;
      let range = "A" + row + ":O" + row;
      spreadsheet
        .getRange(range)
        .setValues([
          [
            data[i][FIRST_NAME_COL],
            data[i][LAST_NAME_COL],
            data[i][PHONE_COL],
            data[i][MON],
            data[i][TUE],
            data[i][WED],
            data[i][THU],
            data[i][FRI],
            data[i][SAT],
            data[i][SUN],
            concatSchedule(data[i]),
            data[i][PHX],
            data[i][PEORIA],
            data[i][RV],
            data[i][GY],
          ],
        ]);
    }
  }
  spreadsheet
    .getRange("A:P")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");
}
//==================================================================================================================================

function checkIfWorking(input) {
  if (input[WORKING_COLUMN] == "Yes, Excited to be back!") {
    return true;
  }
  return false;
}
//==================================================================================================================================

//returns the concatenated schedule for given person
function concatSchedule(input) {
  return (
    isEmpty(input[MON], "MON") +
    isEmpty(input[TUE], "TUE") +
    isEmpty(input[WED], "WED") +
    isEmpty(input[THU], "THU") +
    isEmpty(input[FRI], "FRI") +
    isEmpty(input[SAT], "SAT") +
    isEmpty(input[SUN], "SUN")
  );
}
//==================================================================================================================================

//check if string is empty null or undefined
function isEmpty(str, dayOfWeek) {
  if (!str || 0 === str.length) {
    return "";
  }
  return dayOfWeek + ":" + str + " ";
}
//==================================================================================================================================

function sendOutTexts() {
  let sheet = SpreadsheetApp.getActiveSheet();

  let data = sheet.getDataRange().getValues();
  let status = "";
  for (i = 1; i < data.length; i++) {
    let row = data[i];
    let messageText =
      "Hi " +
      row[0] +
      ", here is your schedule for the upcoming season " +
      row[10] +
      " if you have any questions please contact your general manager";
      
    try {
      response_data = sendSms(row[2], messageText);
      status = "sent";
    } catch (err) {
      Logger.log(err);
      status = "error";
    }
    sheet.getRange('P' + (Number(i) + 1)).setValue(status);
  }
  sheet.getRange('P:P').setTextRotation(0)
}
