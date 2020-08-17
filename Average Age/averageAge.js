var CLASS_LIST = {} //dictionary holding all of the classes
var STUDENT_LIST = {} //dictionary holding all of the students
const CLASSES_SHEET = 'CLA-1'
const STUDENTS_SHEET = 'STU-1'
var count = { found: 0, notFound: 0 }
function main() {
    getRawData()//grab the raw user data and link students with their correct classes
    Logger.log('\n\nLINKING DONE\n FOUND = ' + count.found + '\nNOT FOUND = ' + count.notFound)
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
    Logger.log('DONE\n' + "COUNTER " + counter.value + '\nNOT Found ' + counter.notFound)
}

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
//FORMAT THE STUDENTS TIME TO BE SAME AS INSTRUCTOR TIME
function formatTime(instructor, time) {
    let end = time.indexOf('y', 5) + 3 //find the first y which is end of the day of week, ie monda(y), tuesda(y) and add 3 to get start index of time
    let schedule = time.slice(0, 3) + '-' + time.slice(end)//format the time to be like class schedule
    return schedule + instructor.trim()
}

function linkStudents(count, key) {
    if (CLASS_LIST[STUDENT_LIST[key].swimTimeFormatted] !== undefined) { // if key if valid
        CLASS_LIST[[STUDENT_LIST[key].swimTimeFormatted]].students.push(STUDENT_LIST[key])// put students to their respective classes
        count.found++
    }
    else {
        count.notFound++
        Logger.log(STUDENT_LIST[key].name)
        Logger.log(STUDENT_LIST[key].swimTimeFormatted)
    }
}
