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

// ==================================================================================================================
// COLUMNS WHERE SPECIFIC DATA IS LOCATED ON DATA SHEET, CHANGE IF NEEDED
var CLASS_COLUMN = 1
var TIME_COLUMN = 2
var INSTRUCTOR_COLUMN = 4
var MAX_COLUMN= 6 
var CURRENT_COLUMN=7
var OPENINGS_COLUMNS=8
var MASTER_SHEET_NAME = 'DATA'
var DASHBOARD_SHEET_NAME = 'Dashboard'
// ==================================================================================================================


var currents = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // hold the number of currents for all of the shifts throughout a week
var maxes = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]; // hold the maximum values for all of the shifts throughout a week

// ==================================================================================================================

function main() {
  createNewSheets() //create the new sheets comment out if already created sheets
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(MASTER_SHEET_NAME), true);
  var data = spreadsheet.getDataRange().getValues();

  let output = "";

  for (var i = 1; i < data.length; i++) {
    output = getShift(data, i, TIME_COLUMN); // get the shift that this will belong to
    addToList(output, data[i]); //enter the whole row into the correct list
    addToMax(output, data[i][MAX_COLUMN]); //get the values of maxes for all of the shifts
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
  
  try{
   spreadsheet.setActiveSheet(spreadsheet.getSheetByName(DASHBOARD_SHEET_NAME), true);// try and open sheet to see if sheets have already been created
   return
  }
  catch(err){// 
  
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
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(DASHBOARD_SHEET_NAME), true); //switch to dashboard sheet

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

  spreadsheet.getRange("A1:E1").setBackground("#ff00ff");
  spreadsheet.getRange('A1:E1').setFontWeight('bold')

  let maxSum = 0;
  let currentSum = 0;
  let openingSum = 0;

  let temp = 2;
  for (let y = 0; y < 13; y++) {
   
    let openings = max[y] - currents[y];
    let percentFull = ((currents[y] / max[y]) * 100.0).toFixed(2);
    percentFull = percentFull + "%";
    
    let range = 'B'+temp+':E'+temp
    spreadsheet.getRange(range).setValues([[max[y],currents[y],openings,percentFull]])
    
    let r = y+1
    
    if(y %2 == 0 && y!= 0){spreadsheet.getRange('A'+r+':E'+r).setBackground('#bbbbbb')}

    maxSum = maxSum + max[y];
    currentSum = currentSum + currents[y];
    openingSum = openingSum + openings;
    temp++;
  }

  spreadsheet.getRange("A15:E15").setValues([
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
            shiftList[index][y][OPENINGS_COLUMNS]
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
    spreadsheet.getRange('D2:G2').setValues([[max[x],current[x],(max[x]-current[x]),isNaN(((current[x]/max[x])*100).toFixed(2))]])
    spreadsheet.getRange('D2:G2').setBackground('#00ff00')
    spreadsheet.getActiveSheet().setColumnWidths(1, 3, 180)
    calcLevelStats(shiftList, x)
      
}
// ==================================================================================================================
// CALCULATES THE STATS per LEVEL for each shift
// ==================================================================================================================
function calcLevelStats(shiftList, index){
    var spreadsheet = SpreadsheetApp.getActive();
     let title = 'H3:J3'
     let levels = 'H4:H18'
    spreadsheet.getRange(title).setValues([['Level','# of Classes' , 'Percent Full']])
    spreadsheet.getRange(title).setBackground('#cccccc')
    spreadsheet.getRange(levels).setValues([['Baby splash'],['LS1'],['LS2'],['LSA'],['CF'],['GF'],['JF'],['OCT'],['LOB'],['HHJr'],['HHSr'],['Private'],['Semi'],['SN'],['Open']])
    let numClasses = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]//store the count of classes
    let maxSum =     [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]//store the total number of students per level
    let currentSum = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]//store the current sum of students per level
    
    if (index == 10 || index == 11 || index == 12 ){index++}
    
    for(i = 0 ; i <shiftList[index].length; i++){
      switch(shiftList[index][i][CLASS_COLUMN].trim()){
      case 'Baby Splash':
        numClasses[0]++
        maxSum[0] = maxSum[0] + shiftList[index][i][MAX_COLUMN]
        currentSum[0] = currentSum[0] + shiftList[index][i][CURRENT_COLUMN]
        break
      case 'Little Snapper 1':
        numClasses[1]++
        maxSum[1] = maxSum[1] + shiftList[index][i][MAX_COLUMN]
        currentSum[1] = currentSum[1] + shiftList[index][i][CURRENT_COLUMN]
        break
      case 'Little Snapper 2':
        numClasses[2]++
        maxSum[2] = maxSum[2] + shiftList[index][i][MAX_COLUMN]
        currentSum[2] = currentSum[2] + shiftList[index][i][CURRENT_COLUMN]
        break
      case 'Little Snapper Advanced':
        numClasses[3]++
        maxSum[3] = maxSum[3] + shiftList[index][i][MAX_COLUMN]
        currentSum[3] = currentSum[3] + shiftList[index][i][CURRENT_COLUMN]
        break
      case 'Clownfish':
        numClasses[4]++
        maxSum[4] = maxSum[4] + shiftList[index][i][MAX_COLUMN]
        currentSum[4] = currentSum[4] + shiftList[index][i][CURRENT_COLUMN]
        break 
      case 'Goldfish':
        numClasses[5]++
        maxSum[5] = maxSum[5] + shiftList[index][i][MAX_COLUMN]
        currentSum[5] = currentSum[5] + shiftList[index][i][CURRENT_COLUMN]
        break 
      case 'Jellyfish':
        numClasses[6]++
        maxSum[6] = maxSum[6] + shiftList[index][i][MAX_COLUMN]
        currentSum[6] = currentSum[6] + shiftList[index][i][CURRENT_COLUMN]
        break
      case 'Octopus':
        numClasses[7]++
        maxSum[7] = maxSum[7] + shiftList[index][i][MAX_COLUMN]
        currentSum[7] = currentSum[7] + shiftList[index][i][CURRENT_COLUMN]
        break
      case 'Lobster':
        numClasses[8]++
        maxSum[8] = maxSum[8] + shiftList[index][i][MAX_COLUMN]
        currentSum[8] = currentSum[8] + shiftList[index][i][CURRENT_COLUMN]
        break
      case 'Hammerhead Junior':
        numClasses[9]++
        maxSum[9] = maxSum[9] + shiftList[index][i][MAX_COLUMN]
        currentSum[9] = currentSum[9] + shiftList[index][i][CURRENT_COLUMN]
        break
      case 'Hammerhead Senior':
        numClasses[10]++
        maxSum[10] = maxSum[10] + shiftList[index][i][MAX_COLUMN]
        currentSum[10] = currentSum[10] + shiftList[index][i][CURRENT_COLUMN]
        break 
      case 'Private':
        numClasses[11]++
        maxSum[11] = maxSum[11] + shiftList[index][i][MAX_COLUMN]
        currentSum[11] = currentSum[11] + shiftList[index][i][CURRENT_COLUMN]
        break  
      case 'Semi':
        numClasses[12]++
        maxSum[12] = maxSum[12] + shiftList[index][i][MAX_COLUMN]
        currentSum[12] = currentSum[12] + shiftList[index][i][CURRENT_COLUMN]
        break 
      }
    }
    let numClassesRange = 'I4:I18'
    let percentFullRange = 'J4:J18'
    spreadsheet.getRange(numClassesRange).setValues([[numClasses[0]],[numClasses[1]],[numClasses[2]],[numClasses[3]],[numClasses[4]],[numClasses[5]],[numClasses[6]],[numClasses[7]],[numClasses[8]],[numClasses[9]],[numClasses[10]],[numClasses[11]],[numClasses[12]],[numClasses[13]],[numClasses[14]]])
    spreadsheet.getRange(percentFullRange).setValues([
      [isNaN(((currentSum[0]/maxSum[0])*100).toFixed(2))],
      [isNaN(((currentSum[1]/maxSum[1])*100).toFixed(2))],
      [isNaN(((currentSum[2]/maxSum[2])*100).toFixed(2))],
      [isNaN(((currentSum[3]/maxSum[3])*100).toFixed(2))],
      [isNaN(((currentSum[4]/maxSum[4])*100).toFixed(2))],
      [isNaN(((currentSum[5]/maxSum[5])*100).toFixed(2))],
      [isNaN(((currentSum[6]/maxSum[6])*100).toFixed(2))],
      [isNaN(((currentSum[7]/maxSum[7])*100).toFixed(2))],
      [isNaN(((currentSum[8]/maxSum[8])*100).toFixed(2))],
      [isNaN(((currentSum[9]/maxSum[9])*100).toFixed(2))],
      [isNaN(((currentSum[10]/maxSum[10])*100).toFixed(2))],
      [isNaN(((currentSum[11]/maxSum[11])*100).toFixed(2))],
      [isNaN(((currentSum[12]/maxSum[12])*100).toFixed(2))],
      [isNaN(((currentSum[13]/maxSum[13])*100).toFixed(2))],
      [isNaN(((currentSum[14]/maxSum[14])*100).toFixed(2))],
      
    ])
}
// ==================================================================================================================
//    Checks if the given input is NaN string and leaves plan else it adds percent onto it
// ==================================================================================================================

function isNaN(input){
  if(input === 'NaN'){
    return ""
  }
  else{
    return input+'%'
  }
}