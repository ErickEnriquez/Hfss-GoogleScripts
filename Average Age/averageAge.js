var CLASS_LIST = {} //dictionary holding all of the classes
var STUDENT_LIST = {} //dictionary holding all of the students
const CLASSES_SHEET = 'CLA-1'
const STUDENTS_SHEET = 'STU-1'

function main() {
    getRawData()
}

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
                maxAge: Number(classes[i][3].slice(4, 5))
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
    var counter = { value: 0 }
    for (i = 1; i < studentInfo.length; i++) {
        x = {
            name: studentInfo[i][0],
            age: Number(studentInfo[i][1]),
            studentKeywords: studentInfo[i][2],
            level: studentInfo[i][3],
            time: studentInfo[i][4],
            instructor: studentInfo[i][5],
            classReference: getClass(studentInfo[i][5], studentInfo[i][4], counter)
        }
        key = x.name
        STUDENT_LIST[key] = x
    }

    Logger.log('DONE\n' + "COUNTER " + counter.value)

}

function getClass(instructor, time, counter) {
    let end = time.indexOf('y', 5) + 3 //find the first y which is end of the day of week, ie monda(y) and add 3 to get start index of time
    let schedule = time.slice(0, 3) + '-' + time.slice(end)//format the time to be like class schedule
    let tempKey = schedule + instructor.trim()
    let keyList = Object.keys(CLASS_LIST)
    if (CLASS_LIST[tempKey] !== undefined) {
        counter.value++
        return CLASS_LIST[tempKey]
    } else {
        return undefined
    }

}
