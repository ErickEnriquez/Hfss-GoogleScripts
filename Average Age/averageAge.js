var CLASS_LIST = {} //dictionary holding all of the classes
var STUDENT_LIST = {} //dictionary holding all of the students
const CLASSES_SHEET = 'CLA-1'
const STUDENTS_SHEET = 'STU-1'
const STATS_SHEET = 'STATISTICS'
var count = { found: 0, notFound: 0 }
//==========================================================================================================================
function main() {
    createNewSheet()
    getRawData()//grab the raw user data and link students with their correct classes
    //Logger.log('\n\nLINKING DONE\n FOUND = ' + count.found + '\nNOT FOUND = ' + count.notFound)
    calculateAverageClassAge() // calculate average age of students per class
    createNewSheet() // try and create new sheet if needed
    fillInSheet()//fill in the sheet with data
}
//==========================================================================================================================
function getRawData() {
    // GRAB DATA FROM THE CLASSES
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(
        spreadsheet.getSheetByName(CLASSES_SHEET),
        true
    );
    var classes = spreadsheet.getDataRange().getValues();
    let x //temp variable  
    let key //key value for the dictionary
    for (i = 1; i < classes.length; i++) {
        if (classes[i][3] != '') {
            x = {
                className: classes[i][0],
                schedule: classes[i][1],
                instructor: classes[i][2],
                ageRange: classes[i][3],
                minAge: Number(classes[i][3].slice(0, 1)),
                maxAge: Number(classes[i][3].slice(4, 5)),
                students: []//will store the students who belong to this class
            }
            key = x.schedule.trim() + x.instructor.trim()
            CLASS_LIST[key] = x
        }
    }

    //GRAB DATA FOR THE STUDENTS
    spreadsheet.setActiveSheet(
        spreadsheet.getSheetByName(STUDENTS_SHEET),
        true
    );
    let studentInfo = spreadsheet.getDataRange().getValues();
    var counter = { value: 0, notFound: 0 }
    for (i = 1; i < studentInfo.length; i++) {
        x = {
            name: studentInfo[i][0],
            age: Number(studentInfo[i][1]),
            studentKeywords: studentInfo[i][2],
            level: studentInfo[i][3],
            time: studentInfo[i][4],
            instructor: studentInfo[i][5],
            swimTimeFormatted: formatTime(studentInfo[i][5], studentInfo[i][4])
        }
        x.classReference = getClass(x.swimTimeFormatted, counter)
        key = x.name
        STUDENT_LIST[key] = x
        linkStudents(count, key)
    }
    // Logger.log('DONE\n' + "COUNTER " + counter.value + '\nNOT Found ' + counter.notFound)
}
//==========================================================================================================================
//CHECKS IF KEY is VALID IN CLASS LIST, returns reference to class if true
function getClass(key, counter) {
    if (CLASS_LIST[key] !== undefined) {
        counter.value++
        return CLASS_LIST[key]
    } else {
        counter.notFound++
        return undefined
    }
}
//==========================================================================================================================
//FORMAT THE STUDENTS TIME TO BE SAME AS INSTRUCTOR TIME
function formatTime(instructor, time) {
    let end = time.indexOf('y', 5) + 3 //find the first y which is end of the day of week, ie monda(y), tuesda(y) and add 3 to get start index of time
    let schedule = time.slice(0, 3) + '-' + time.slice(end)//format the time to be like class schedule
    return schedule + instructor.trim()
}
//==========================================================================================================================
function linkStudents(count, key) {
    if (CLASS_LIST[STUDENT_LIST[key].swimTimeFormatted] !== undefined) { // if key if valid
        CLASS_LIST[[STUDENT_LIST[key].swimTimeFormatted]].students.push(STUDENT_LIST[key])// put students to their respective classes
        count.found++
    }
    else {
        count.notFound++
        //Logger.log(STUDENT_LIST[key].name)
        //Logger.log(STUDENT_LIST[key].swimTimeFormatted)
    }
}
//==========================================================================================================================

function calculateAverageClassAge() {
    for (let key in CLASS_LIST) {
        average = 0
        for (let i in CLASS_LIST[key].students) {
            average += CLASS_LIST[key].students[i].age
        }
        CLASS_LIST[key].averageAge = (average / CLASS_LIST[key].students.length).toFixed(2)
    }
}
//==========================================================================================================================
function createNewSheet() {
    var spreadsheet = SpreadsheetApp.getActive();
    try {
        spreadsheet.setActiveSheet(
            spreadsheet.getSheetByName(STATS_SHEET),
            true
        );
        return
    } catch (e) {
        //create sheet there
        spreadsheet.insertSheet(0);
        spreadsheet.getActiveSheet().setName(STATS_SHEET);
    }
}
//==========================================================================================================================
function fillInSheet() {
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(
        spreadsheet.getSheetByName(STATS_SHEET),
        true
    );
    spreadsheet.getDataRange().clearContent(); // clear out old content from spreadsheet
    const headerRow = "A1:J1";
    spreadsheet.getRange(headerRow).setValues([
        [
            'Student',
            'Age',
            'Student keywords',
            'Schedule',
            'Instructor',
            'Class Time',
            'Age Within Range',
            'Level Name',
            'Age Range',
            'Average Class Age'
        ]
    ])
    spreadsheet.getActiveSheet().setColumnWidths(1, 10, 140);
    spreadsheet.getRange(headerRow).setVerticalAlignment("middle").setHorizontalAlignment("center").setBackground("#000b8c").setFontColor("#FFFFFF")
    let rowIndex = 2
    let inRangeObject
    let flag = false
    for (let key in CLASS_LIST) {
        for (let i in CLASS_LIST[key].students) {
            let row = 'A' + rowIndex + ':J' + rowIndex
            inRangeObject = checkIfWithinRange(i, key),//check if the students age is within the age range
                spreadsheet.getRange(row).setValues([[
                    CLASS_LIST[key].students[i].name,
                    CLASS_LIST[key].students[i].age,
                    CLASS_LIST[key].students[i].studentKeywords,
                    CLASS_LIST[key].students[i].time,
                    CLASS_LIST[key].students[i].instructor,
                    CLASS_LIST[key].students[i].swimTimeFormatted,
                    inRangeObject.text,
                    CLASS_LIST[key].className,
                    CLASS_LIST[key].ageRange,
                    CLASS_LIST[key].averageAge,
                ]])
            let classAverageObj = checkClassAge(key)
            spreadsheet.getRange(row).setBackground(classAverageObj.color)
            spreadsheet.getRange("G"+rowIndex+":G"+rowIndex).setBackground(inRangeObject.color)
            rowIndex++
            if (flag == false) {
                spreadsheet.getRange(row).setBackground('#DDDDDD')
            }
        }
        if (flag == false) {
            flag = true
        }
        else if (flag == true) {
            flag = false
        }
    }
}
//==========================================================================================================================
function checkIfWithinRange(index, key) {
    if (CLASS_LIST[key].students[index].age < CLASS_LIST[key].minAge || CLASS_LIST[key].students[index].age > CLASS_LIST[key].maxAge) {
        return { text: 'No', color: '#d80422' }
    }
    else {
        return { text: 'Yes', color: '#ffffff' }
    }
}
//==========================================================================================================================

function checkClassAge(key){
    if(CLASS_LIST[key].averageAge > CLASS_LIST[key].maxAge){
        return {response: false , color: '#f6a800'} //return orange
    }
    else if(CLASS_LIST[key].averageAge < CLASS_LIST.minAge){
        return {response: false , color : '#3daf2c'} //return green
    }
    else{
        return {response : true , color: '#FFFFFF'} //return white
    }
}