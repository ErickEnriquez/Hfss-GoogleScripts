///////////////////////////////////////////////////////////////////////////////////////
const SID = "AC9f2f8a075203b765ee7724fbf1c8bb17";
const token = "d53462cfa67afe893f96c6a7bad52cc7";
const companyPhone = "6026712206";
const turtlesImage =
    "https://hfss-website.s3-us-west-2.amazonaws.com/turtles.png";
///////////////////////////////////////////////////////////////////////////////////////

//==================================================================================================================================
//main driver for the entire script
function main() {
  let message = showInputPrompt()


}

//==================================================================================================================================
function showInputPrompt() {
    let ui = SpreadsheetApp.getUi();
    let response = ui.prompt('Enter your message here', 'To make personalized message include ${name} where you want to add persons name ie \"Hello ${name}\"', ui.ButtonSet.YES_NO)
    return response
}

//==================================================================================================================================

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