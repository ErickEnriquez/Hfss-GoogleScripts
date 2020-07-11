/////////////////////////////////////////////////////// 
const SID = 'AC9f2f8a075203b765ee7724fbf1c8bb17'
const token = 'd53462cfa67afe893f96c6a7bad52cc7'
const companyPhone = "6026712206"
const formatSheet = 'Format for text all'
const workingSheet = 'WORKING sheet GMs'
/////////////////////////////////////////////////////// 
let FIRST_NAME_COLUMN = 'A2:A'
let LAST_NAME_COLUMN = 'B2:B'
let PHONE_NUMBER_COLUMN = 'D2:D'
let SCHEDULE_DATA = 'J2:W'
///////////////////////////////////////////////////////

function sendSms(to, body) {
  var messages_url = "https://api.twilio.com/2010-04-01/Accounts/" + SID + "/Messages.json";

  var payload = {
    "To": to,
    "Body" : body,
    "From" : companyPhone
  };

  var options = {
    "method" : "post",
    "payload" : payload
  };

  options.headers = { 
    "Authorization" : "Basic " + Utilities.base64Encode(SID+':'+token)
  };

  UrlFetchApp.fetch(messages_url, options);
}


function copyToFormatted(){
   

    let spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(workingSheet), true);

    let numbRows = spreadsheet.getLastRow();

    FIRST_NAME_COLUMN+= numbRows;
    LAST_NAME_COLUMN+= numbRows;
    PHONE_NUMBER_COLUMN+= numbRows;
    SCHEDULE_DATA+= numbRows
    
    let firstName = spreadsheet.getRange(FIRST_NAME_COLUMN).getValues();
    let lastName = spreadsheet.getRange(LAST_NAME_COLUMN).getValues();
    let phoneNumber =spreadsheet.getRange(PHONE_NUMBER_COLUMN).getValues();
    let schedule = spreadsheet.getRange(SCHEDULE_DATA).getValues();

    try {
        spreadsheet.setActiveSheet(spreadsheet.getSheetByName(formatSheet))
        spreadsheet.getDataRange().clearContent();// clear out old content from spreadsheet
      
    } catch (error) {
        Logger.log(error)
    }   
     const headerRow = 'A1:P1'
        spreadsheet.getRange(headerRow).setValues([['First Name','Last Name','Phone',' Mon Sched','Tue Sched','Wed Sched','Thu Sched','Fri Sched','Sat Sched','Sun Sched','Total','PHX','PEORIA','RV', 'GY', 'Text Status']])
        spreadsheet.getActiveSheet().setColumnWidths(1, 3, 120)
        spreadsheet.getActiveSheet().setColumnWidths(4, 13, 60)

}




function sendOutTexts() {
  
}