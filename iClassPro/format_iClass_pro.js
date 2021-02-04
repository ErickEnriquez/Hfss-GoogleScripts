// ==================================================================================================================
// This script takes in raw data and formats it to a usable format for hfss

// author Erick Enriquez
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
//==================================================================================================================

var shifts = {
  monAM: { shiftName: 'Mon-AM', index: 0 , classArray:[], max:0 , current:0 },
  monPM: { shiftName: 'Mon-PM', index: 1 , classArray:[], max:0 , current:0},
  tueAM: { shiftName: 'Tue-AM', index: 2 , classArray:[], max:0 , current:0},
  tuePM: { shiftName: 'Tue-PM', index: 3 , classArray:[], max:0 , current:0},
  wedAM: { shiftName: 'Wed-AM', index: 4 , classArray:[], max:0 , current:0},
  wedPM: { shiftName: 'Wed-PM', index: 5 , classArray:[], max:0 , current:0},
  thuAM: { shiftName: 'Thu-AM', index: 6 , classArray:[], max:0 , current:0},
  thuPM: { shiftName: 'Thu-PM', index: 7 , classArray:[], max:0 , current:0},
  friAM: { shiftName: 'Fri-AM', index: 8 , classArray:[], max:0 , current:0},
  friPM: { shiftName: 'Fri-PM', index: 9 , classArray:[], max:0 , current:0},
  sat: { shiftName: 'Sat', index: 14 , classArray:[], max:0 , current:0},
  sunAM: { shiftName: 'Sun-AM', index: 12 , classArray:[], max:0 , current:0},
  sunPM: { shiftName: 'Sun-PM', index: 13 , classArray:[], max:0 , current:0},
  satAM: { shiftName: 'Sat-AM', index: 10 , classArray:[], max:0 , current:0},
  satPM: { shiftName: 'Sat-PM', index: 11 , classArray:[], max:0 , current:0},

}

var shiftMap = new Map()

shiftMap.set('monAM', { shiftName: 'Mon-AM', index: 0, classArray: [], max: 0, current: 0 })
shiftMap.set('monPM', { shiftName: 'Mon-PM', index: 1, classArray: [], max: 0, current: 0 })
shiftMap.set('tueAM', { shiftName: 'Tue-AM', index: 2, classArray: [], max: 0, current: 0 })
shiftMap.set('tuePM', { shiftName: 'Tue-PM', index: 3 , classArray:[], max:0 , current:0})
shiftMap.set('wedAM', { shiftName: 'Wed-AM', index: 4 , classArray:[], max:0 , current:0})
shiftMap.set('wedPM', { shiftName: 'Wed-PM', index: 5 , classArray:[], max:0 , current:0})
shiftMap.set('thuAM', { shiftName: 'Thu-AM', index: 6 , classArray:[], max:0 , current:0})
shiftMap.set('thuPM', { shiftName: 'Thu-PM', index: 7 , classArray:[], max:0 , current:0})
shiftMap.set('friAM', { shiftName: 'Fri-AM', index: 8 , classArray:[], max:0 , current:0})
shiftMap.set('friPM', { shiftName: 'Fri-PM', index: 9 , classArray:[], max:0 , current:0})
shiftMap.set('sat', { shiftName: 'Sat', index: 14 , classArray:[], max:0 , current:0})
shiftMap.set('sunAM', { shiftName: 'Sun-AM', index: 12 , classArray:[], max:0 , current:0})
shiftMap.set('sunPM', { shiftName: 'Sun-PM', index: 13 , classArray:[], max:0 , current:0})


//==================================================================================================================

var classLevels = new Map()

classLevels.set('Baby Splash', { abrv: "BS", name: 'Baby Splash', index: 0, max: 0, current: 0, count: 0 })
classLevels.set('Little Snapper 1', { abrv: 'LS1', name: 'Little Snapper 1', index: 1, max: 0, current: 0, count: 0 })
classLevels.set('Little Snapper 2', { abrv: 'LS2', name: 'Little Snapper 2', index: 2, max: 0, current: 0, count: 0 })
classLevels.set('Little Snapper Advanced', { abrv: 'LSA', name: 'Little Snapper Advanced', index: 3, max: 0, current: 0, count: 0 })
classLevels.set('Clownfish', { abrv: 'CF', name: 'Clownfish', index: 4, max: 0, current: 0, count: 0 })
classLevels.set('Goldfish', { abrv: 'GF', name: 'Goldfish', index: 5, max: 0, current: 0, count: 0 })
classLevels.set('Jellyfish', { abrv: 'JF', name: 'Jellyfish', index: 6, max: 0, current: 0, count: 0 })
classLevels.set('Octopus', { abrv: 'OCT', name: 'Octopus', index: 7, max: 0, current: 0, count: 0 })
classLevels.set('Lobster', { abrv: 'LOB', name: 'Lobster', index: 8, max: 0, current: 0, count: 0 })
classLevels.set('Hammerhead Junior', { abrv: 'HHJr', name: 'Hammerhead Junior', index: 9, max: 0, current: 0, count: 0 })
classLevels.set('Hammerhead Senior', { abrv: 'HHSr', name: 'Hammerhead Senior', index: 10, max: 0, current: 0, count: 0 })
classLevels.set('Private - Teacher and Me', { abrv: 'Private', name: 'Private - Teacher and Me', index: 11, max: 0, current: 0, count: 0 })
classLevels.set('Private - Special Needs', { abrv: 'Private SN', name: 'Private - Special Needs', index: 18, max: 0, current: 0, count: 0 })
classLevels.set('Semi', { abrv: 'Semi', name: 'Semi', index: 12, max: 0, current: 0, count: 0 })
classLevels.set('SN', { abrv: 'SN', name: 'SN', index: 13, max: 0, current: 0, count: 0 })//should be unused in calculations leave in until we remove all code that references SN
classLevels.set('.Unassigned (Teacher and Me) Level', { abrv: 'Unassigned', name: '.Unassigned (Teacher and Me) Level', index: 14, max: 0, current: 0, count: 0 })
classLevels.set('Water Watcher', { abrv: 'Water Watcher', name: 'Water Watcher', index: 15, max: 0, current: 0, count: 0 })
classLevels.set('Break', { abrv: 'Break', name: 'Break', index: 16, max: 0, current: 0, count: 0 })
classLevels.set('Squad', { abrv: 'Squad', name: 'Squad', index: 17, max: 0, current: 0, count: 0 })
classLevels.set('Other', { abrv: 'Other', name: 'Other', index: 19, max: 0, current: 0, count: 0 })



// ==================================================================================================================

//var maxList =     [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // holds max for whole week by level
//var currentList = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // holds current for whole week by level
//var countList =   [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // holds count for whole week by level

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

function main() {
  createNewSheets(); //create the new sheets comment out if already created sheets
  var spreadsheet = SpreadsheetApp.getActive(); spreadsheet.setActiveSheet( spreadsheet.getSheetByName(MASTER_SHEET_NAME));
  var data = spreadsheet.getDataRange().getValues();

  let output = "";
  for (var i = 1; i < data.length; i++) {
    output = getShift(data, i, TIME_COLUMN); // get the shift that this will belong to
    addToData(output, data[i]); //enter the whole row into the correct list and update max and current lists
    calcTotalLevelStats(data, i); // add the class to its level stats
  }

  var currents = [
    shifts.monAM.current,
    shifts.monPM.current,
    shifts.tueAM.current,
    shifts.tuePM.current,
    shifts.wedAM.current,
    shifts.wedPM.current,
    shifts.thuAM.current,
    shifts.thuPM.current,
    shifts.friAM.current,
    shifts.friPM.current,
    shifts.sat.current,
    shifts.sunAM.current,
    shifts.sunPM.current,
  ]; // hold the number of currents for all of the shifts throughout a week for dashboard use
  var maxes = [
    shifts.monAM.max,
    shifts.monPM.max,
    shifts.tueAM.max,
    shifts.tuePM.max,
    shifts.wedAM.max,
    shifts.wedPM.max,
    shifts.thuAM.max,
    shifts.thuPM.max,
    shifts.friAM.max,
    shifts.friPM.max,
    shifts.sat.max,
    shifts.sunAM.max,
    shifts.sunPM.max,
  ]; // hold the maximum values for all of the shifts throughout a week for dashboard use

  writeToDashboard(maxes, currents); //function to write the aggregate data to dashboard sheet

  /*
  let shiftList = [
    shifts.monAM.classArray,
    shifts.monPM.classArray,
    shifts.tueAM.classArray,
    shifts.tuePM.classArray,
    shifts.wedAM.classArray,
    shifts.wedPM.classArray,
    shifts.thuAM.classArray,
    shifts.thuPM.classArray,
    shifts.friAM.classArray,
    shifts.friPM.classArray,
    [],//add empty list to line up indexes on writeToShiftSheet
    shifts.sat.classArray,
    shifts.sunAM.classArray,
    shifts.sunPM.classArray,
  ]

  for (let z = 0; z <= 13; z++) {
    writeToShiftSheet(z, shiftNames, shiftList, maxes, currents);
  }*/
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
  if (day == 'Sat') {// no need to check for AM or PM on Saturday
    return 'Sat'
  }
  else if (day == "Sun") {
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
    //create the needed sheets if not already there
    for (days of shiftMap.values()) {
      spreadsheet.insertSheet(0);
      spreadsheet.getActiveSheet().setName(days.shiftName);
    }
  }
}
// ==================================================================================================================
//checks the date given and adds the data to correct list and adds to max and current lists
// ==================================================================================================================

function addToData(date, data) {
  for (days of shiftMap.values()) {
    if (days.shiftName == date) {
     days.classArray.push(data)
     days.max = days.max + data[MAX_COLUMN]
     days.current = days.current + data[CURRENT_COLUMN]
    }
  }
}

// ==================================================================================================================
// =======================    Fills in the data for the dashboard sheet part of the spreadsheet =====================
function writeToDashboard(max, currents) {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(
    spreadsheet.getSheetByName(DASHBOARD_SHEET_NAME),
    true
  ); //switch to dashboard sheet
  spreadsheet.getDataRange().clearContent(); // clear out old content from spreadsheet
 
  //insert the days of the week in A row
  let index = 2
  for (days in shiftMap.values()) {
    let row = 'A' + index + ':A' +index;
    spreadsheet.getRange(row).setValues([[days.shiftName]])
    index++
  }

  spreadsheet
    .getRange("A1:I1")
    .setValues([["Data Range", "Max", "Current", "Openings", "Percent Full", "" , "# of Classes", "Level", "PercentFull"]]);

  spreadsheet.getRange("A1:I1").setBackground("#0000ff");
  spreadsheet.getRange("A1:I1").setFontColor("#ffffff");
  spreadsheet.getRange("A1:I1").setFontWeight("bold");

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
    //alternate colors ever other row
    if (y % 2 == 0 && y != 0) {
      spreadsheet.getRange("A" + r + ":E" + r).setBackground("#bbbbbb");
    }

    maxSum = maxSum + max[y];
    currentSum = currentSum + currents[y];
    openingSum = openingSum + openings;
    temp++;
  }
 const totalRow = "A15:E15"
  spreadsheet .getRange(totalRow) .setValues([ [ "Total", maxSum, currentSum, openingSum, isNotANumber(((currentSum / maxSum) * 100).toFixed(2)), ], ]);
  spreadsheet.getRange(totalRow).setBackground("#000000");
  spreadsheet.getRange(totalRow).setFontColor("#ffffff");

  formatSheet(spreadsheet, "A:K");

  //write class specific details 
  ( () => {
    let index = 0
    let row = 2
    for (let value of classLevels.values()) {
      let range = 'G' + row + ':I' + row
      spreadsheet.getRange(range).setValues([[value.abrv, value.count, isNotANumber(((value.current / value.max) * 100).toFixed(2))]])
      index++
      row++
    } 
  }) ()
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
  //retrieve the name of the class and calculate the statistics O(n^2) Worse case
  for (i = 0; i < shiftList[index].length; i++) {
    let level = shiftList[index][i][CLASS_COLUMN].trim()
    for (value of classLevels.values()) {
      if (level == value.name) {
        numClasses[value.index]++;
        maxSum[value.index] = maxSum[value.index] + shiftList[index][i][MAX_COLUMN];
        currentSum[value.index] = currentSum[value.index] + shiftList[index][i][CURRENT_COLUMN]
        break;
      }
    }
  }
  // WRITE INDIVIDUAL ROWS ONTO SHIFT SHEET
  ( () => { 
    let index = 0
    let row = 2
    for (let value of classLevels.values()) {
      let range = 'I' + row + ':K' + row
      spreadsheet.getRange(range).setValues([[value.abrv, value.count, isNotANumber(((value.current / value.max) * 100).toFixed(2))]])
      index++
      row++
    } 
  }) ();
  
  formatSheet(spreadsheet, levelRange);
  formatSheet(spreadsheet, titleRange);
  spreadsheet.getRange(titleRange).setBackground("#cccccc");
}
// ==================================================================================================================
//    Checks if the given input is NaN string and leaves plan else it adds percent onto it
// ==================================================================================================================

function isNotANumber(input) {
  if (input === "NaN" || isNaN(input) || input == "Infinity") {
    return "N/A";
  } else {
    return Math.round(Number(input)) + "%";// remove the decimal places and round
  }
}
// ==================================================================================================================
// CHECKS LEVEL AND ADDS TO THE TOTAL LIST FOR DASHBOARD USE
// ==================================================================================================================
function calcTotalLevelStats(data, index) {
  let level = data[index][CLASS_COLUMN].trim()
  for (let value of classLevels.values()) {
    if (level == value.name) {
      Logger.log(level + " = " + value.name)
      value.count++;
      value.max = value.max + data[index][MAX_COLUMN];
      value.current = value.current + data[index][CURRENT_COLUMN];
      break;
    }
  }
}

// ==================================================================================================================
// FORMATS CURRENT SPREADSHEET
// ==================================================================================================================

function formatSheet(spreadsheet,range) {
  spreadsheet.getRange(range).activate();
  spreadsheet.getActiveRangeList()
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");
}
