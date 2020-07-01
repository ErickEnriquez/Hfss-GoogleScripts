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

var maxList = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // holds max for whole week by level
var currentList = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // holds current for whole week by level
var countList = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // holds count for whole week by level

// ==================================================================================================================
// COLUMNS WHERE SPECIFIC DATA IS LOCATED ON DATA SHEET, CHANGE IF NEEDED
var CLASS_COLUMN = 11;
var TIME_COLUMN = 2;
var INSTRUCTOR_COLUMN = 4;
var MAX_COLUMN = 6;
var CURRENT_COLUMN = 7;
var OPENINGS_COLUMNS = 8;
var MASTER_SHEET_NAME = "Table 01";
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
    addToList(output, data[i]); //enter the whole row into the correct list
    addToMax(output, data[i][MAX_COLUMN]); //get the values of maxes for all of the shifts
    calcTotalLevelStats(data, i); // add the class to its level stats
    addToCurrents(output, data[i][CURRENT_COLUMN]); //get the values for the currents of all of the shifts
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
    //

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
  spreadsheet.setActiveSheet(
    spreadsheet.getSheetByName(DASHBOARD_SHEET_NAME),
    true
  ); //switch to dashboard sheet

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
    percentFull = percentFull + "%";

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
        isNaN(((currentSum / maxSum) * 100).toFixed(2)),
      ],
    ]);
  spreadsheet.getRange("A15:E15").setBackground("#000000");
  spreadsheet.getRange("A15:E15").setFontColor("#ffffff");

  const levelStatRange = "G1:I20";
  spreadsheet.getRange(levelStatRange).setValues([
    ["Level", "# of Classes", " PercentFull"],
    [
      "Baby Splash",
      countList[0],
      isNaN(((currentList[0] / maxList[0]) * 100).toFixed(2)),
    ],
    [
      "LS1",
      countList[1],
      isNaN(((currentList[1] / maxList[1]) * 100).toFixed(2)),
    ],
    [
      "LS2",
      countList[2],
      isNaN(((currentList[2] / maxList[2]) * 100).toFixed(2)),
    ],
    [
      "LSA",
      countList[3],
      isNaN(((currentList[3] / maxList[3]) * 100).toFixed(2)),
    ],
    [
      "CF",
      countList[4],
      isNaN(((currentList[4] / maxList[4]) * 100).toFixed(2)),
    ],
    [
      "GF",
      countList[5],
      isNaN(((currentList[5] / maxList[5]) * 100).toFixed(2)),
    ],
    [
      "JF",
      countList[6],
      isNaN(((currentList[6] / maxList[6]) * 100).toFixed(2)),
    ],
    [
      "OCT",
      countList[7],
      isNaN(((currentList[7] / maxList[7]) * 100).toFixed(2)),
    ],
    [
      "LOB",
      countList[8],
      isNaN(((currentList[8] / maxList[8]) * 100).toFixed(2)),
    ],
    [
      "HHJr",
      countList[9],
      isNaN(((currentList[9] / maxList[9]) * 100).toFixed(2)),
    ],
    [
      "HHSr",
      countList[10],
      isNaN(((currentList[10] / maxList[10]) * 100).toFixed(2)),
    ],
    [
      "Private",
      countList[11],
      isNaN(((currentList[11] / maxList[11]) * 100).toFixed(2)),
    ],
    [
      "Semi",
      countList[12],
      isNaN(((currentList[12] / maxList[12]) * 100).toFixed(2)),
    ],
    [
      "SN",
      countList[13],
      isNaN(((currentList[13] / maxList[13]) * 100).toFixed(2)),
    ],
    [
      "Open",
      countList[14],
      isNaN(((currentList[14] / maxList[14]) * 100).toFixed(2)),
    ],
    [
      "Water Watcher",
      countList[15],
      isNaN(((currentList[15] / maxList[15]) * 100).toFixed(2)),
    ],
    [
      "Break",
      countList[16],
      isNaN(((currentList[16] / maxList[16]) * 100).toFixed(2)),
    ],
    [
      "Squads",
      countList[17],
      isNaN(((currentList[17] / maxList[17]) * 100).toFixed(2)),
    ],
    [
      "Other",
      countList[18],
      isNaN(((currentList[18] / maxList[18]) * 100).toFixed(2)),
    ],
  ]);
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
  spreadsheet.getRange("A1:G1").setBackground("#000000");
  spreadsheet.getRange("A1:G1").setFontColor("#ffffff");

  let numRows = shiftList[index].length;
  let temp = 3;
  for (let y = 0; y < numRows; y++) {
    let range = "A" + temp + ":F" + temp; //grab the current row
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
        isNaN(((current[x] / max[x]) * 100).toFixed(2)),
      ],
    ]);
  spreadsheet.getRange("D2:G2").setBackground("#00ff00");
  spreadsheet.getActiveSheet().setColumnWidths(1, 3, 180);
  spreadsheet
    .getRange("A:F")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("left");
  calcLevelStats(shiftList, x);
}
// ==================================================================================================================
// CALCULATES THE STATS per LEVEL for each shift
// ==================================================================================================================
function calcLevelStats(shiftList, index) {
  var spreadsheet = SpreadsheetApp.getActive();

  const levelRange = "H3:J22";
  const titleRange = "H3:J3";

  let numClasses = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; //store the count of classes
  let maxSum = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; //store the total number of students per level
  let currentSum = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; //store the current sum of students per level

  if (index == 10 || index == 11 || index == 12) {
    index++;
  }

  for (i = 0; i < shiftList[index].length; i++) {
    switch (shiftList[index][i][CLASS_COLUMN].trim()) {
      case "Baby Splash":
        numClasses[0]++;
        maxSum[0] = maxSum[0] + shiftList[index][i][MAX_COLUMN];
        currentSum[0] = currentSum[0] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Little Snapper 1":
        numClasses[1]++;
        maxSum[1] = maxSum[1] + shiftList[index][i][MAX_COLUMN];
        currentSum[1] = currentSum[1] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Little Snapper 2":
        numClasses[2]++;
        maxSum[2] = maxSum[2] + shiftList[index][i][MAX_COLUMN];
        currentSum[2] = currentSum[2] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Little Snapper Advanced":
        numClasses[3]++;
        maxSum[3] = maxSum[3] + shiftList[index][i][MAX_COLUMN];
        currentSum[3] = currentSum[3] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Clownfish":
        numClasses[4]++;
        maxSum[4] = maxSum[4] + shiftList[index][i][MAX_COLUMN];
        currentSum[4] = currentSum[4] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Goldfish":
        numClasses[5]++;
        maxSum[5] = maxSum[5] + shiftList[index][i][MAX_COLUMN];
        currentSum[5] = currentSum[5] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Jellyfish":
        numClasses[6]++;
        maxSum[6] = maxSum[6] + shiftList[index][i][MAX_COLUMN];
        currentSum[6] = currentSum[6] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Octopus":
        numClasses[7]++;
        maxSum[7] = maxSum[7] + shiftList[index][i][MAX_COLUMN];
        currentSum[7] = currentSum[7] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Lobster":
        numClasses[8]++;
        maxSum[8] = maxSum[8] + shiftList[index][i][MAX_COLUMN];
        currentSum[8] = currentSum[8] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Hammerhead Junior":
        numClasses[9]++;
        maxSum[9] = maxSum[9] + shiftList[index][i][MAX_COLUMN];
        currentSum[9] = currentSum[9] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Hammerhead Senior":
        numClasses[10]++;
        maxSum[10] = maxSum[10] + shiftList[index][i][MAX_COLUMN];
        currentSum[10] = currentSum[10] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Private":
        numClasses[11]++;
        maxSum[11] = maxSum[11] + shiftList[index][i][MAX_COLUMN];
        currentSum[11] = currentSum[11] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Semi":
        numClasses[12]++;
        maxSum[12] = maxSum[12] + shiftList[index][i][MAX_COLUMN];
        currentSum[12] = currentSum[12] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "SN":
        numClasses[13]++;
        maxSum[13] = maxSum[13] + shiftList[index][i][MAX_COLUMN];
        currentSum[13] = currentSum[13] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Open":
        numClasses[14]++;
        maxSum[14] = maxSum[14] + shiftList[index][i][MAX_COLUMN];
        currentSum[14] = currentSum[14] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Water Watcher":
        numClasses[15]++;
        maxSum[15] = maxSum[15] + shiftList[index][i][MAX_COLUMN];
        currentSum[15] = currentSum[15] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "x Break":
        numClasses[16]++;
        maxSum[16] = maxSum[16] + shiftList[index][i][MAX_COLUMN];
        currentSum[16] = currentSum[16] + shiftList[index][i][CURRENT_COLUMN];
        break;
      case "Squad":
        numClasses[17]++;
        maxSum[17] = maxSum[17] + shiftList[index][i][MAX_COLUMN];
        currentSum[17] = currentSum[17] + shiftList[index][i][CURRENT_COLUMN];
        break;
      default:
        numClasses[18]++;
        maxSum[18] = maxSum[18] + shiftList[index][i][MAX_COLUMN];
        currentSum[18] = currentSum[18] + shiftList[index][i][CURRENT_COLUMN];
        break;
    }
  }
  spreadsheet.getRange(levelRange).setValues([
    ["Level", "# of classes", "Percent Full"],
    [
      "Baby Splash",
      numClasses[0],
      isNaN(((currentSum[0] / maxSum[0]) * 100).toFixed(2)),
    ],
    [
      "LS1",
      numClasses[1],
      isNaN(((currentSum[1] / maxSum[1]) * 100).toFixed(2)),
    ],
    [
      "LS2",
      numClasses[2],
      isNaN(((currentSum[2] / maxSum[2]) * 100).toFixed(2)),
    ],
    [
      "LSA",
      numClasses[3],
      isNaN(((currentSum[3] / maxSum[3]) * 100).toFixed(2)),
    ],
    [
      "CF",
      numClasses[4],
      isNaN(((currentSum[4] / maxSum[4]) * 100).toFixed(2)),
    ],
    [
      "GF",
      numClasses[5],
      isNaN(((currentSum[5] / maxSum[5]) * 100).toFixed(2)),
    ],
    [
      "JF",
      numClasses[6],
      isNaN(((currentSum[6] / maxSum[6]) * 100).toFixed(2)),
    ],
    [
      "OCT",
      numClasses[7],
      isNaN(((currentSum[7] / maxSum[7]) * 100).toFixed(2)),
    ],
    [
      "LOB",
      numClasses[8],
      isNaN(((currentSum[8] / maxSum[8]) * 100).toFixed(2)),
    ],
    [
      "HHJr",
      numClasses[9],
      isNaN(((currentSum[9] / maxSum[9]) * 100).toFixed(2)),
    ],
    [
      "HHSr",
      numClasses[10],
      isNaN(((currentSum[10] / maxSum[10]) * 100).toFixed(2)),
    ],
    [
      "Private",
      numClasses[11],
      isNaN(((currentSum[11] / maxSum[11]) * 100).toFixed(2)),
    ],
    [
      "Semi",
      numClasses[12],
      isNaN(((currentSum[12] / maxSum[12]) * 100).toFixed(2)),
    ],
    [
      "SN",
      numClasses[13],
      isNaN(((currentSum[13] / maxSum[13]) * 100).toFixed(2)),
    ],
    [
      "Open",
      numClasses[14],
      isNaN(((currentSum[14] / maxSum[14]) * 100).toFixed(2)),
    ],
    [
      "Water Watcher",
      numClasses[15],
      isNaN(((currentSum[15] / maxSum[15]) * 100).toFixed(2)),
    ],
    [
      "Break",
      numClasses[16],
      isNaN(((currentSum[16] / maxSum[16]) * 100).toFixed(2)),
    ],
    [
      "Squad",
      numClasses[17],
      isNaN(((currentSum[17] / maxSum[17]) * 100).toFixed(2)),
    ],
    [
      "Other",
      numClasses[18],
      isNaN(((currentSum[18] / maxSum[18]) * 100).toFixed(2)),
    ],
  ]);
  spreadsheet
    .getRange(levelRange)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("left");
  spreadsheet.getRange(titleRange).setBackground("#cccccc");
}
// ==================================================================================================================
//    Checks if the given input is NaN string and leaves plan else it adds percent onto it
// ==================================================================================================================

function isNaN(input) {
  if (input === "NaN") {
    return "";
  } else {
    return input + "%";
  }
}
// ==================================================================================================================
// CHECKS LEVEL AND ADDS TO THE TOTAL LIST FOR DASHBOARD USE
// ==================================================================================================================
function calcTotalLevelStats(data, index) {
  switch (data[index][CLASS_COLUMN].trim()) {
    case "Baby Splash":
      countList[0]++;
      maxList[0] = maxList[0] + data[index][MAX_COLUMN];
      currentList[0] = currentList[0] + data[index][CURRENT_COLUMN];
      break;
    case "Little Snapper 1":
      countList[1]++;
      maxList[1] = maxList[1] + data[index][MAX_COLUMN];
      currentList[1] = currentList[1] + data[index][CURRENT_COLUMN];
      break;
    case "Little Snapper 2":
      countList[2]++;
      maxList[2] = maxList[2] + data[index][MAX_COLUMN];
      currentList[2] = currentList[2] + data[index][CURRENT_COLUMN];
      break;
    case "Little Snapper Advanced":
      countList[3]++;
      maxList[3] = maxList[3] + data[index][MAX_COLUMN];
      currentList[3] = currentList[3] + data[index][CURRENT_COLUMN];
      break;
    case "Clownfish":
      countList[4]++;
      maxList[4] = maxList[4] + data[index][MAX_COLUMN];
      currentList[4] = currentList[4] + data[index][CURRENT_COLUMN];
      break;
    case "Goldfish":
      countList[5]++;
      maxList[5] = maxList[5] + data[index][MAX_COLUMN];
      currentList[5] = currentList[5] + data[index][CURRENT_COLUMN];
      break;
    case "Jellyfish":
      countList[6]++;
      maxList[6] = maxList[6] + data[index][MAX_COLUMN];
      currentList[6] = currentList[6] + data[index][CURRENT_COLUMN];
      break;
    case "Octopus":
      countList[7]++;
      maxList[7] = maxList[7] + data[index][MAX_COLUMN];
      currentList[7] = currentList[7] + data[index][CURRENT_COLUMN];
      break;
    case "Lobster":
      countList[8]++;
      maxList[8] = maxList[8] + data[index][MAX_COLUMN];
      currentList[8] = currentList[8] + data[index][CURRENT_COLUMN];
      break;
    case "Hammerhead Junior":
      countList[9]++;
      maxList[9] = maxList[9] + data[index][MAX_COLUMN];
      currentList[9] = currentList[9] + data[index][CURRENT_COLUMN];
      break;
    case "Hammerhead Senior":
      countList[10]++;
      maxList[10] = maxList[10] + data[index][MAX_COLUMN];
      currentList[10] = currentList[10] + data[index][CURRENT_COLUMN];
      break;
    case "Private":
      countList[11]++;
      maxList[11] = maxList[11] + data[index][MAX_COLUMN];
      currentList[11] = currentList[11] + data[index][CURRENT_COLUMN];
      break;
    case "Semi":
      countList[12]++;
      maxList[12] = maxList[12] + data[index][MAX_COLUMN];
      currentList[12] = currentList[12] + data[index][CURRENT_COLUMN];
      break;
    case "SN":
      countList[13]++;
      maxList[13] = maxList[13] + data[index][MAX_COLUMN];
      currentList[13] = currentList[13] + data[index][CURRENT_COLUMN];
      break;
    case "Open":
      countList[14]++;
      maxList[14] = maxList[14] + data[index][MAX_COLUMN];
      currentList[14] = currentList[14] + data[index][CURRENT_COLUMN];
      break;
    case "Water Watcher":
      countList[15]++;
      maxList[15] = maxList[15] + data[index][MAX_COLUMN];
      currentList[15] = currentList[15] + data[index][CURRENT_COLUMN];
      break;
    case "x Break":
      countList[16]++;
      maxList[16] = maxList[16] + data[index][MAX_COLUMN];
      currentList[16] = currentList[16] + data[index][CURRENT_COLUMN];
      break;
    case "Squad":
      countList[17]++;
      maxList[17] = maxList[17] + data[index][MAX_COLUMN];
      currentList[17] = currentList[17] + data[index][CURRENT_COLUMN];
      break;
    default:
      countList[18]++;
      maxList[18] = maxList[18] + data[index][MAX_COLUMN];
      currentList[18] = currentList[18] + data[index][CURRENT_COLUMN];
      Logger.log(data[index][CLASS_COLUMN].trim());
      break;
  }
}
