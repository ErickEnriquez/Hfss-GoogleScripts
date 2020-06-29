var monAM = [];
var monPM = [];
var tueAM = [];
var tuePM = [];
var wedAM = [];
var wedPM = [];
var thuAM = [];
var thuPM = [];
var friAM = [];
var friPM = [];
var sat = [];
var sunAM = [];
var sunPM = [];

var shiftNames = [
  "Mon-AM", //0
  "Mon-PM", //1
  "Tue-AM", //2
  "Tue-PM", //3
  "Wed-AM", //4
  "Wed-PM", //5
  "Thu-AM", //6
  "Thu-PM", //7
  "Fri-AM", //8
  "Fri-PM", //9
  "Sat-AM", //10
  "Sat-PM", //11
  "Sun-AM", //12
  "Sun-PM", //13
  "Sat", //temp 14
];
var currents = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // hold the number of currents for all of the shifts throughout a week
var maxes = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // hold the maximum values for all of the shifts throughout a week

//var x = 2;
// ==================================================================================================================

function main() {
  //createNewSheets() //create the new sheets
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Data"), true);
  var data = spreadsheet.getDataRange().getValues();

  let output = "";

  for (var i = 1; i < data.length; i++) {
    output = getShift(data, i, 1); // get the shift that this will belong to
    addToList(output, data[i]); //enter the whole row into the correct list
    addToMax(output, data[i][3]); //get the values of maxes for all of the shifts
    addToCurrents(output, data[i][4]); //get the values for the currents of all of the shifts
  
  }
  writeToDashboard(maxes, currents, shiftNames); //function to write the aggragate data to dashboard sheet

  let shiftList = [
    monAM,
    monPM,
    tueAM,
    tuePM,
    wedAM,
    wedPM,
    thuAM,
    thuPM,
    friAM,
    friPM,
    [],//add empty list to line up indexed on writeToShiftSheet
    sat,
    sunAM,
    sunPM,
  ];


  for (let z = 0; z <= 13; z++) {
      writeToShiftSheet(z, shiftNames, shiftList, maxes, currents);
  }
}

// ==================================================================================================================
//returns the shift data given an array and indexes
// ==================================================================================================================

function getShift(data, i, j) {
  var outputShift = "";
  var day = data[i][j].slice(0, 3);
  var time = data[i][j].slice(4, 6);
  var flag = false;
  if (day == "Sun") {
    if((time[1] == ":" && time[0] == 8) ||
    time[0] == 9 ||
    time[0] == 1 ){
      return day + "-AM"
    }

  } else if (
  (time[1] == ":" && time[0] == 8) ||
    time[0] == 9 ||
    time[0] == 1 ||
    time[0] == 2
  ) {
    flag = true;
  } else if (time[1] == 0 || time[1] == 1) {
    flag = true;
  }
  outputShift = day;
  if (flag == true) {
    outputShift = outputShift + "-AM";
  } else {
    outputShift = outputShift + "-PM";
  }
  return outputShift;
}
// ==================================================================================================================
//creates all of the sheets that we need
// ==================================================================================================================

function createNewSheets() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName(shiftNames[13]); //sun PM
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName(shiftNames[12]); // sun am
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName(shiftNames[14]); // sat
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName(shiftNames[9]); //fri PM
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName(shiftNames[8]); //fri am
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName(shiftNames[7]); //thu pm
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName(shiftNames[6]); //thu am
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName(shiftNames[5]); //wed pm
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName(shiftNames[4]); //wed am
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName(shiftNames[3]); //tue pm
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName(shiftNames[2]); //tue am
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName(shiftNames[1]); //mon pm
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName(shiftNames[0]); // mon am
  spreadsheet.insertSheet(1);
  spreadsheet.getActiveSheet().setName("Dashboard");
}
// ==================================================================================================================
//checks the date given and adds the data to correct list
// ==================================================================================================================

function addToList(date, data) {
  switch (date) {
    case shiftNames[0]:
      monAM.push(data);
      break;
    case shiftNames[1]:
      monPM.push(data);
      break;
    case shiftNames[2]:
      tueAM.push(data);
      break;
    case shiftNames[3]:
      tuePM.push(data);
      break;
    case shiftNames[4]:
      wedAM.push(data);
      break;
    case shiftNames[5]:
      wedPM.push(data);
      break;
    case shiftNames[6]:
      thuAM.push(data);
      break;
    case shiftNames[7]:
      thuPM.push(data);
      break;
    case shiftNames[8]:
      friAM.push(data);
      break;
    case shiftNames[9]:
      friPM.push(data);
      break;
    case shiftNames[10]:
      sat.push(data);
      break;
    case shiftNames[11]:
      sat.push(data);
      break;
    case shiftNames[12]:
      sunAM.push(data);
      break;
    case shiftNames[13]:
      sunPM.push(data);
      break;
  }
}
// ==================================================================================================================

// ==================================================================================================================

function addToMax(date, data) {
  switch (date) {
    case shiftNames[0]: //mon am
      maxes[0] = maxes[0] + data;
      break;
    case shiftNames[1]: //mon pm
      maxes[1] = maxes[1] + data;
      break;
    case shiftNames[2]: //tue am
      maxes[2] = maxes[2] + data;
      break;
    case shiftNames[3]: //tue pm
      maxes[3] = maxes[3] + data;
      break;
    case shiftNames[4]: //wed am
      maxes[4] = maxes[4] + data;
      break;
    case shiftNames[5]: //wed pm
      maxes[5] = maxes[5] + data;
      break;
    case shiftNames[6]: //thur am
      maxes[6] = maxes[6] + data;
      break;
    case shiftNames[7]: //thur pm
      maxes[7] = maxes[7] + data;
      break;
    case shiftNames[8]: //fri am
      maxes[8] = maxes[8] + data;
      break;
    case shiftNames[9]: //fri pm
      maxes[9] = maxes[9] + data;
      break;
    case shiftNames[10]: //sat am
      maxes[10] = maxes[10] + data;
      break;
    case shiftNames[11]: //sat pm
      maxes[10] = maxes[10] + data;
      break;
    case shiftNames[12]: //sun am
      maxes[11] = maxes[11] + data;
      break;
    case shiftNames[13]: //sun pm
      maxes[12] = maxes[12] + data;
      break;
  }
}
// ==================================================================================================================

// ==================================================================================================================

function addToCurrents(date, data) {
  switch (date) {
    case shiftNames[0]: //mon am
      currents[0] = currents[0] + data;
      break;
    case shiftNames[1]: //mon pm
      currents[1] = currents[1] + data;
      break;
    case shiftNames[2]: //tue am
      currents[2] = currents[2] + data;
      break;
    case shiftNames[3]: //tue pm
      currents[3] = currents[3] + data;
      break;
    case shiftNames[4]: //wed am
      currents[4] = currents[4] + data;
      break;
    case shiftNames[5]: //wed pm
      currents[5] = currents[5] + data;
      break;
    case shiftNames[6]: //thur am
      currents[6] = currents[6] + data;
      break;
    case shiftNames[7]: //thur pm
      currents[7] = currents[7] + data;
      break;
    case shiftNames[8]: //fri am
      currents[8] = currents[8] + data;
      break;
    case shiftNames[9]: //fri pm
      currents[9] = currents[9] + data;
      break;
    case shiftNames[10]: //sat am
      currents[10] = currents[10] + data;
      break;
    case shiftNames[11]: //sat pm
      currents[10] = currents[10] + data;
      break;
    case shiftNames[12]: //sun am
      currents[11] = currents[11] + data;
      break;
    case shiftNames[13]: //sun pm
      currents[12] = currents[12] + data;
      break;
  }
}
// ==================================================================================================================
// =======================    Fills in the data for the dashboard sheet part of the spreadsheet =====================
function writeToDashboard(max, currents, shiftName) {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Dashboard"), true); //switch to dashboard sheet

  spreadsheet.getRange("A2:A14").setValues([
    [shiftName[0]],
    [shiftName[1]],
    [shiftName[2]],
    [shiftName[3]],
    [shiftName[4]],
    [shiftName[5]],
    [shiftName[6]],
    [shiftName[7]],
    [shiftName[8]],
    [shiftName[9]],
    [shiftName[14]], //skip sat am and pm instead write sat
    [shiftName[12]],
    [shiftName[13]],
  ]);

  spreadsheet
    .getRange("A1:E1")
    .setValues([["Data Range", "Max", "Current", "Openings", "Percent Full"]]);

  spreadsheet.getRange("A1:E1").setBackground("#999999");

  let maxSum = 0;
  let currentSum = 0;
  let openingSum = 0;

  let temp = 2;
  for (let y = 0; y < 13; y++) {
    let maxCol = "B" + temp;
    let currentCol = "C" + temp;
    let openCol = "D" + temp;
    let percentCol = "E" + temp;

    let openings = max[y] - currents[y];
    let percentFull = ((currents[y] / max[y]) * 100.0).toFixed(2);
    percentFull = percentFull + "%";

    spreadsheet.getRangeByName(maxCol).activate();
    spreadsheet.getCurrentCell().setValue(max[y]);
    spreadsheet.getRangeByName(currentCol).activate();
    spreadsheet.getCurrentCell().setValue(currents[y]);
    spreadsheet.getRangeByName(openCol).activate();
    spreadsheet.getCurrentCell().setValue(openings);
    spreadsheet.getRangeByName(percentCol).activate();
    spreadsheet.getCurrentCell().setValue(percentFull);
    
    let r = y+1
    
    if(y %2 == 0 && y!= 0){spreadsheet.getRange('A'+r+':E'+r).setBackground('#bbbbbb')}

    maxSum = maxSum + max[y];
    currentSum = currentSum + currents[y];
    openingSum = openingSum + openings;
    temp++;
  }

  let tempList = [
    [
      "Total",
      maxSum,
      currentSum,
      openingSum,
      ((currentSum / maxSum) * 100).toFixed(2),
    ],
  ];
  spreadsheet.getRange("A15:E15").setValues(tempList);
  spreadsheet.getRange("A15:E15").setBackground("#000000");
  spreadsheet.getRange("A15:E15").setFontColor("#ffffff");
}

// ==================================================================================================================
// =======================    Fills in the data for the shift sheet part of the spreadsheet =====================
function writeToShiftSheet(index, shiftName, shiftList, max, current) {
  if (index == 10) {
    return;
  }
  var spreadsheet = SpreadsheetApp.getActive();
  if (index == 11) {// switch to the shift sheet
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(shiftName[14]), true); // open the special Sat spreadsheet
  } 
  else {
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(shiftName[index]),true);
  } 
  spreadsheet
    .getRange("A1:G1")
    .setValues([
      [
        "Class Name",
        "Schedule",
        "Instructor",
        "Max",
        "Current",
        "Openings",
        "Percent Full",
      ],
    ]); //write the title row
  spreadsheet.getRange("A1:1").setBackground("#000000");
  spreadsheet.getRange("A1:G1").setFontColor("#ffffff");
  
  let numRows = shiftList[index].length;
  let temp = 3;
    for (let y = 0; y < numRows; y++) {
      let range = "A" + temp + ":F" + temp; //grab the current row
      spreadsheet
        .getRange(range)
        .setValues([
          [
            shiftList[index][y][0],
            shiftList[index][y][1],
            shiftList[index][y][2],
            shiftList[index][y][3],
            shiftList[index][y][4],
            shiftList[index][y][5],
          ],
        ]);
        if(temp%2 == 0){//darken alternating rows
        spreadsheet.getRange(range).setBackground('#bbbbbb')
        }
      temp++;
    }
    let x = index
    if(index == 11 || index == 12 || index == 13){
    x--
    }
        spreadsheet.getRange('D2:G2').setValues([[max[x],current[x],(max[x]-current[x]),(((current[x]/max[x])*100).toFixed(2))]])
        spreadsheet.getRange('D2:G2').setBackground('#00ff00')
}
