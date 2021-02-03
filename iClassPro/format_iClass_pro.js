// ==================================================================================================================
// This script takes in raw data and formats it to a usable format for hfss

// author Erick Enriquez
// ==================================================================================================================

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
// ==================================================================================================================

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

var ltsLevels = {
  BS: { abrv: "BS", name: "Baby Splash", index: 0 },
  LS1: { abrv: 'LS1', name: 'Little Snapper 1', index: 1 },
  LS2: { abrv: 'LS2', name: 'Little Snapper 2', index: 2 },
  LSA: { abrv: 'LSA', name: 'Little Snapper Advanced', index: 3 },
  CF: { abrv: 'CF', name: 'Clownfish', index: 4 },
  GF: { abrv: 'GF', name: 'Goldfish', index: 5 },
  JF: { abrv: 'JF', name: 'Jellyfish', index: 6 },
  OCT: { abrv: 'OCT', name: 'Octopus', index: 7 },
  LOB: { abrv: 'LOB', name: 'Lobster', index: 8 },
  HHJr: { abrv: 'HHJr', name: 'Hammerhead Junior', index: 9 },
  HHSr: { abrv: 'HHSr', name: 'Hammerhead Senior', index: 10 },
  Private: { abrv: 'Private', name: 'Private - Teacher and Me', index: 11 },
  PrivateSN: { abrv: 'Private SN', name: "Private - Special Needs", index: 18 },
  Semi: { abrv: 'Semi', name: 'Semi', index: 12 },
  SN: { abrv: 'SN', name: 'SN', index: 13 },//should be unused in calculations leave in until we remove all code that references SN 
  Unassigned: { abrv: 'Unassigned', name: '.Unassigned (Teacher and Me) Level', index: 14 },
  WaterWatcher: { abrv: 'Water Watcher', name: 'Water Watcher', index: 15 },
  Break: { abrv: 'Break', name: 'Break', index: 16 },
  Squads: { abrv: 'Squads', name: 'Squad', index: 17 },
  Other: { abrv: 'Other', name: 'Other', index: 19 },
}
// ==================================================================================================================

var maxList =     [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // holds max for whole week by level
var currentList = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // holds current for whole week by level
var countList =   [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // holds count for whole week by level

// ==================================================================================================================
// COLUMNS WHERE SPECIFIC DATA IS LOCATED ON DATA SHEET, CHANGE IF NEEDED
var CLASS_COLUMN = 11;
var TIME_COLUMN = 2;
var INSTRUCTOR_COLUMN = 3;
var MAX_COLUMN = 7;
var CURRENT_COLUMN = 8;
var OPENINGS_COLUMNS = 9;
var MASTER_SHEET_NAME = "RAW";
var DASHBOARD_SHEET_NAME = "Dashboard";
// ==================================================================================================================

var currents = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // hold the number of currents for all of the shifts throughout a week
var maxes = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // hold the maximum values for all of the shifts throughout a week

// ==================================================================================================================

function main() {
  createNewSheets(); //create the new sheets comment out if already created sheets
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(
    spreadsheet.getSheetByName(MASTER_SHEET_NAME),
    true
  );
  var data = spreadsheet.getDataRange().getValues();

  let output = "";

  for (var i = 1; i < data.length; i++) {
    output = getShift(data, i, TIME_COLUMN); // get the shift that this will belong to
    addToData(output, data[i]); //enter the whole row into the correct list and update max and current lists
    calcTotalLevelStats(data, i); // add the class to its level stats
  }
  writeToDashboard(maxes, currents, shiftNames); //function to write the aggregate data to dashboard sheet

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
    [], //add empty list to line up indexed on writeToShiftSheet
    sat,
    sunAM,
    sunPM,
  ];

  for (let z = 0; z <= 13; z++) {
    writeToShiftSheet(z, shiftNames, shiftList, maxes, currents);
  }
  spreadsheet.setActiveSheet(
    spreadsheet.getSheetByName(MASTER_SHEET_NAME),
    true
  ); // go back to master sheet
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
    if ((time[1] == ":" && time[0] == 8) || time[0] == 9 || time[0] == 1) {
      return day + "-AM";
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

  try {
    spreadsheet.setActiveSheet(
      spreadsheet.getSheetByName(DASHBOARD_SHEET_NAME),
      true
    ); // try and open sheet to see if sheets have already been created
    return;
  } catch (err) {
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
    spreadsheet.getActiveSheet().setName(DASHBOARD_SHEET_NAME);
  }
}
// ==================================================================================================================
//checks the date given and adds the data to correct list and adds to max and current lists
// ==================================================================================================================

function addToData(date, data) {
  switch (date) {
    case shiftNames[0]:
      pushData(monAM,0,data)
      break;
    case shiftNames[1]:
      pushData(monPM,1,data)
      break;
    case shiftNames[2]:
      pushData(tueAM,2,data)
      break;
    case shiftNames[3]:
      pushData(tuePM,3,data)
      break;
    case shiftNames[4]:
      pushData(wedAM,4,data)
      break;
    case shiftNames[5]:
      pushData(wedPM,5,data)
      break;
    case shiftNames[6]:
      pushData(thuAM,6,data)
      break;
    case shiftNames[7]:
      pushData(thuPM,7,data)
      break;
    case shiftNames[8]:
      pushData(friAM,8,data)
      break;
    case shiftNames[9]:
      pushData(friPM,9,data)
      break;
    case shiftNames[10]:
      pushData(sat,10,data)
      break;
    case shiftNames[11]:
      pushData(sat,10,data)
      break;
    case shiftNames[12]:
      pushData(sunAM,11,data)
      break;
    case shiftNames[13]:
      pushData(sunPM,12,data)
      break;
  }
  //actual function that pushes data to avoid repetition
   function pushData (dateArray, index, data) {
    dateArray.push(data);
    maxes[index] = maxes[index] + data[MAX_COLUMN];
    currents[index] = currents[index] + data[CURRENT_COLUMN];
  }
}

// ==================================================================================================================
// =======================    Fills in the data for the dashboard sheet part of the spreadsheet =====================
function writeToDashboard(max, currents, shiftName) {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(
    spreadsheet.getSheetByName(DASHBOARD_SHEET_NAME),
    true
  ); //switch to dashboard sheet
  spreadsheet.getDataRange().clearContent(); // clear out old content from spreadsheet


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

  spreadsheet.getRange("A1:E1").setBackground("#0000ff");
  spreadsheet.getRange("A1:E1").setFontColor("#ffffff");
  spreadsheet.getRange("A1:E1").setFontWeight("bold");

  let maxSum = 0;
  let currentSum = 0;
  let openingSum = 0;

  let temp = 2;
  for (let y = 0; y < 13; y++) {
    let openings = max[y] - currents[y];
    let percentFull = ((currents[y] / max[y]) * 100.0).toFixed(2);
    percentFull = isNotANumber(percentFull);

    let range = "B" + temp + ":E" + temp;
    spreadsheet
      .getRange(range)
      .setValues([[max[y], currents[y], openings, percentFull]]);

    let r = y + 1;

    if (y % 2 == 0 && y != 0) {
      spreadsheet.getRange("A" + r + ":E" + r).setBackground("#bbbbbb");
    }

    maxSum = maxSum + max[y];
    currentSum = currentSum + currents[y];
    openingSum = openingSum + openings;
    temp++;
  }

  spreadsheet
    .getRange("A15:E15")
    .setValues([
      [
        "Total",
        maxSum,
        currentSum,
        openingSum,
        isNotANumber(((currentSum / maxSum) * 100).toFixed(2)),
      ],
    ]);
  spreadsheet.getRange("A15:E15").setBackground("#000000");
  spreadsheet.getRange("A15:E15").setFontColor("#ffffff");

  const levelStatRange = "G1:I20";
  const titleRange = "G1:I1";
  spreadsheet.getRange(levelStatRange).setValues([
    ["Level", "# of Classes", " PercentFull"],
    [ ltsLevels.BS.abrv, countList[0], isNotANumber(((currentList[0] / maxList[0]) * 100).toFixed(2)), ],
    [ ltsLevels.LS1.abrv, countList[1], isNotANumber(((currentList[1] / maxList[1]) * 100).toFixed(2)), ],
    [ ltsLevels.LS2.abrv, countList[2], isNotANumber(((currentList[2] / maxList[2]) * 100).toFixed(2)), ],
    [ ltsLevels.LSA.abrv, countList[3], isNotANumber(((currentList[3] / maxList[3]) * 100).toFixed(2)), ],
    [ ltsLevels.CF.abrv, countList[4], isNotANumber(((currentList[4] / maxList[4]) * 100).toFixed(2)), ],
    [ ltsLevels.GF.abrv, countList[5], isNotANumber(((currentList[5] / maxList[5]) * 100).toFixed(2)), ],
    [ ltsLevels.JF.abrv, countList[6], isNotANumber(((currentList[6] / maxList[6]) * 100).toFixed(2)), ],
    [ ltsLevels.OCT.abrv, countList[7], isNotANumber(((currentList[7] / maxList[7]) * 100).toFixed(2)), ],
    [ ltsLevels.LOB.abrv, countList[8], isNotANumber(((currentList[8] / maxList[8]) * 100).toFixed(2)), ],
    [ ltsLevels.HHJr.abrv, countList[9], isNotANumber(((currentList[9] / maxList[9]) * 100).toFixed(2)), ],
    [ ltsLevels.HHSr.abrv, countList[10], isNotANumber(((currentList[10] / maxList[10]) * 100).toFixed(2)), ],
    [ ltsLevels.Private.abrv, countList[11], isNotANumber(((currentList[11] / maxList[11]) * 100).toFixed(2)), ],
    [ltsLevels.PrivateSN.abrv, countList[18], isNotANumber(((currentList[18] / maxList[18]) * 100).toFixed(2)) ],
    [ ltsLevels.Semi.abrv, countList[12], isNotANumber(((currentList[12] / maxList[12]) * 100).toFixed(2)), ],
    //[ ltsLevels.SN.abrv, countList[13], isNotANumber(((currentList[13] / maxList[13]) * 100).toFixed(2)), ],
    [ ltsLevels.Unassigned.abrv, countList[14], isNotANumber(((currentList[14] / maxList[14]) * 100).toFixed(2)), ],
    [ ltsLevels.WaterWatcher.abrv, countList[15], isNotANumber(((currentList[15] / maxList[15]) * 100).toFixed(2)), ],
    [ ltsLevels.Break.abrv, countList[16], isNotANumber(((currentList[16] / maxList[16]) * 100).toFixed(2)), ],
    [ ltsLevels.Squads.abrv, countList[17], isNotANumber(((currentList[17] / maxList[17]) * 100).toFixed(2)), ],
    [ ltsLevels.Other.abrv, countList[19], isNotANumber(((currentList[19] / maxList[19]) * 100).toFixed(2)), ],
  ]);
  spreadsheet.getRange(titleRange).setBackground("#0000ff");
  spreadsheet.getRange(titleRange).setFontColor("#ffffff");
  formatSheet(spreadsheet, "A:K");
}

// ==================================================================================================================
// =======================    Fills in the data for the shift sheet part of the spreadsheet =====================
function writeToShiftSheet(index, shiftName, shiftList, max, current) {
  if (index == 10) {
    return;
  }
  var spreadsheet = SpreadsheetApp.getActive();
  if (index == 11) {
    // switch to the shift sheet
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName(shiftName[14]), true); // open the special Sat spreadsheet
  } else {
    spreadsheet.setActiveSheet(
      spreadsheet.getSheetByName(shiftName[index]),
      true
    );
  }
  spreadsheet.getDataRange().clearContent(); // clear out old content from spreadsheet
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

  let numRows = shiftList[index].length;
  let temp = 3;
  for (let y = 0; y < numRows; y++) {
    let range = "A" + temp + ":G" + temp; //grab the current row
    spreadsheet
      .getRange(range)
      .setValues([
        [
          shiftList[index][y][CLASS_COLUMN],
          shiftList[index][y][TIME_COLUMN],
          shiftList[index][y][INSTRUCTOR_COLUMN],
          shiftList[index][y][MAX_COLUMN],
          shiftList[index][y][CURRENT_COLUMN],
          shiftList[index][y][OPENINGS_COLUMNS],
          isNotANumber(
            (shiftList[index][y][CURRENT_COLUMN] /
              shiftList[index][y][MAX_COLUMN]) *
              (100.0).toFixed(2)
          ),
        ],
      ]);
    if (temp % 2 == 0) {
      //darken alternating rows
      spreadsheet.getRange(range).setBackground("#bbbbbb");
    }
    temp++;
  }
  let x = index;
  if (index == 11 || index == 12 || index == 13) {
    x--;
  }
  spreadsheet
    .getRange("D2:G2")
    .setValues([
      [
        max[x],
        current[x],
        max[x] - current[x],
        isNotANumber(((current[x] / max[x]) * 100).toFixed(2)),
      ],
    ]);
  spreadsheet.getRange("D2:G2").setBackground("#00ff00");
  spreadsheet.getActiveSheet().setColumnWidths(1, 3, 180);

  spreadsheet.getRange("A1:G1").setBackground("#000000");
  spreadsheet.getRange("A1:G1").setFontColor("#ffffff");
  calcLevelStats(shiftList, x);
  formatSheet(spreadsheet, "A:J");
}
// ==================================================================================================================
// CALCULATES THE STATS per LEVEL for each shift
// ==================================================================================================================
function calcLevelStats(shiftList, index) {
  var spreadsheet = SpreadsheetApp.getActive();

  const levelRange = "I3:K22";
  const titleRange = "I3:K3";

  let numClasses = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; //store the count of classes
  let maxSum = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0]; //store the total number of students per level
  let currentSum = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0]; //store the current sum of students per level

  if (index == 10 || index == 11 || index == 12) {
    index++;
  }

  for (i = 0; i < shiftList[index].length; i++) {
    switch (shiftList[index][i][CLASS_COLUMN].trim()) {
      case "Baby Splash":
        calculateStats(ltsLevels.BS.index, index, i)
        break;
      case "Little Snapper 1":
        calculateStats(ltsLevels.LS1.index, index, i)
        break;
      case "Little Snapper 2":
        calculateStats(ltsLevels.LS2.index, index, i)
        break;
      case "Little Snapper Advanced":
        calculateStats(ltsLevels.LSA.index, index, i)
        break;
      case "Clownfish":
        calculateStats(ltsLevels.CF.index, index, i)
        break;
      case "Goldfish":
        calculateStats(ltsLevels.GF.index, index, i)
        break;
      case "Jellyfish":
        calculateStats(ltsLevels.JF.index, index, i)
        break;
      case "Octopus":
        calculateStats(ltsLevels.OCT.index, index, i)
        break;
      case "Lobster":
        calculateStats(ltsLevels.LOB.index, index, i)
        break;
      case "Hammerhead Junior":
        calculateStats(ltsLevels.HHJr.index, index, i)
        break;
      case "Hammerhead Senior":
        calculateStats(ltsLevels.HHSr.index, index, i);
        break;
      case "Private - Teacher and Me":
        calculateStats(ltsLevels.Private.index, index, i)
        break;
      case "Private - Special Needs":
        calculateStats(ltsLevels.PrivateSN.index, index, i)
        break;
      case "Semi":
        calculateStats(ltsLevels.Semi.index, index, i)
        break;
      case "SN":
        calculateStats(ltsLevels.SN.index, index, i)
        break;
      case ".Unassigned (Teacher and Me) Level":
        calculateStats(ltsLevels.Unassigned.index, index, i)
        break;
      case "Water Watcher":
        calculateStats(ltsLevels.WaterWatcher.index, index, i)
        break;
      case "Break":
        calculateStats(ltsLevels.Break.index, index, i)
        break;
      case "Squad":
        calculateStats(ltsLevels.Squads.index, index, i)
        break;
      default:
        calculateStats(ltsLevels.Other.index, index, i);
        break;
    }//inner func to calculate stats cleaner
    function calculateStats(classIndex, dayIndex, index3){
      numClasses[classIndex]++;
      maxSum[classIndex] = maxSum[classIndex] + shiftList[dayIndex][index3][MAX_COLUMN];
      currentSum[classIndex] = currentSum[classIndex] + shiftList[dayIndex][index3][CURRENT_COLUMN];
    }
  }
  spreadsheet.getRange(levelRange).setValues([
    ["Level", "# of classes", "Percent Full"],
    [ ltsLevels.BS.abrv, numClasses[0], isNotANumber(((currentSum[0] / maxSum[0]) * 100).toFixed(2)), ],
    [ ltsLevels.LS1.abrv, numClasses[1], isNotANumber(((currentSum[1] / maxSum[1]) * 100).toFixed(2)), ],
    [ ltsLevels.LS2.abrv, numClasses[2], isNotANumber(((currentSum[2] / maxSum[2]) * 100).toFixed(2)), ],
    [ ltsLevels.LSA.abrv, numClasses[3], isNotANumber(((currentSum[3] / maxSum[3]) * 100).toFixed(2)), ],
    [ ltsLevels.CF.abrv, numClasses[4], isNotANumber(((currentSum[4] / maxSum[4]) * 100).toFixed(2)), ],
    [ ltsLevels.GF.abrv, numClasses[5], isNotANumber(((currentSum[5] / maxSum[5]) * 100).toFixed(2)), ],
    [ ltsLevels.JF.abrv, numClasses[6], isNotANumber(((currentSum[6] / maxSum[6]) * 100).toFixed(2)), ],
    [ ltsLevels.OCT.abrv, numClasses[7], isNotANumber(((currentSum[7] / maxSum[7]) * 100).toFixed(2)), ],
    [ ltsLevels.LOB.abrv, numClasses[8], isNotANumber(((currentSum[8] / maxSum[8]) * 100).toFixed(2)), ],
    [ ltsLevels.HHJr.abrv, numClasses[9], isNotANumber(((currentSum[9] / maxSum[9]) * 100).toFixed(2)), ],
    [ ltsLevels.HHSr.abrv, numClasses[10], isNotANumber(((currentSum[10] / maxSum[10]) * 100).toFixed(2)), ],
    [ ltsLevels.Private.abrv, numClasses[11], isNotANumber(((currentSum[11] / maxSum[11]) * 100).toFixed(2)), ],
    [ ltsLevels.PrivateSN.abrv, numClasses[18], isNotANumber(((currentSum[18] / maxSum[18]) * 100).toFixed(2)), ],
    [ ltsLevels.Semi.abrv, numClasses[12], isNotANumber(((currentSum[12] / maxSum[12]) * 100).toFixed(2)), ],
    //[ ltsLevels.SN.abrv, numClasses[13], isNotANumber(((currentSum[13] / maxSum[13]) * 100).toFixed(2)), ],
    [ ltsLevels.Unassigned.abrv, numClasses[14], isNotANumber(((currentSum[14] / maxSum[14]) * 100).toFixed(2)), ],
    [ ltsLevels.WaterWatcher.abrv, numClasses[15], isNotANumber(((currentSum[15] / maxSum[15]) * 100).toFixed(2)), ],
    [ ltsLevels.Break.abrv, numClasses[16], isNotANumber(((currentSum[16] / maxSum[16]) * 100).toFixed(2)), ],
    [ ltsLevels.Squads.abrv, numClasses[17], isNotANumber(((currentSum[17] / maxSum[17]) * 100).toFixed(2)), ],
    [ ltsLevels.Other.abrv, numClasses[19], isNotANumber(((currentSum[19] / maxSum[19]) * 100).toFixed(2)), ],
  ]);
  formatSheet(spreadsheet, levelRange);
  spreadsheet.getRange(titleRange).setBackground("#cccccc");
}
// ==================================================================================================================
//    Checks if the given input is NaN string and leaves plan else it adds percent onto it
// ==================================================================================================================

function isNotANumber(input) {
  if (input === "NaN") {
    return "N/A";
  } else if (isNaN(input)) {
    return "N/A";
  } else if (input == "Infinity") {
    return "N/A";
  } else {
    return Math.round(Number(input)) + "%";// remove the decimal places and round
  }
}
// ==================================================================================================================
// CHECKS LEVEL AND ADDS TO THE TOTAL LIST FOR DASHBOARD USE
// ==================================================================================================================
function calcTotalLevelStats(data, index) {
  switch (data[index][CLASS_COLUMN].trim()) {
    case "Baby Splash":
      calculateData(index,0)
      break;
    case "Little Snapper 1":
      calculateData(index,1)
      break;
    case "Little Snapper 2":
      calculateData(index,2)
      break;
    case "Little Snapper Advanced":
      calculateData(index,3)
      break;
    case "Clownfish":
      calculateData(index,4)
      break;
    case "Goldfish":
      calculateData(index,5)
      break;
    case "Jellyfish":
      calculateData(index,6)
      break;
    case "Octopus":
      calculateData(index,7)
      break;
    case "Lobster":
      calculateData(index,8)
      break;
    case "Hammerhead Junior":
      calculateData(index,9)
      break;
    case "Hammerhead Senior":
      calculateData(index,10)
      break;
    case "Private - Teacher and Me":
      calculateData(index,11)
      break;
    case "Private - Special Needs":
      calculateData(index,18)
      break;
    case "Semi":
      calculateData(index,12)
      break;
    case "SN":
      calculateData(index,13)
      break;
    case ".Unassigned (Teacher and Me) Level":
      calculateData(index,14)
      break;
    case "Water Watcher":
      calculateData(index,15)
      break;
    case "Break":
      calculateData(index,16)
      break;
    case "Squad":
      calculateData(index,17)
      break;
    default:
      calculateData(index,19)
      Logger.log(data[index][CLASS_COLUMN].trim());
      break;
  }
  function calculateData(dataIndex, countIndex) {
    countList[countIndex]++;
    maxList[countIndex] = maxList[countIndex] + data[dataIndex][MAX_COLUMN];
    currentList[countIndex] = currentList[countIndex] + data[dataIndex][CURRENT_COLUMN];
  }
}

// ==================================================================================================================
// FORMATS CURRENT SPREADSHEET
// ==================================================================================================================

function formatSheet(ss, range) {
  ss.getRange(range).activate();
  ss.getActiveRangeList()
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");
}
