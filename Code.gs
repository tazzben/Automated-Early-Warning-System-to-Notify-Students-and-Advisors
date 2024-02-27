/*
Google Apps Script to Extract Current Scores and Notify Students and Advisors

This script uses the Canvas API to extract the current scores from a list of courses (defined in COURSE_LIST).  It notifies the students of their current score in the these courses over email.  If the student's score is below a threshold (defined as the script property THRESHOLD), the student's name, email, score, and course are written to a spreadsheet (defined in DANGER_SHEET).  EMAIL_STORAGE is used to store emails for future sending if the script executing user is over the daily email quota.

The script must have the following script properties defined (set under File->Project Properties->Script Properties within the Apps Scripts editor): 

+-----------------+---------------------------------------------------------------------------------------------------------------+
| Script Property | Example                                                                                                       |
+-----------------+---------------------------------------------------------------------------------------------------------------+
| THRESHOLD       | 75                                                                                                            |
+-----------------+---------------------------------------------------------------------------------------------------------------+
| DOMAIN          | unomaha.edu                                                                                                   |
+-----------------+---------------------------------------------------------------------------------------------------------------+
| MESSAGE_SUBJECT | CBA Periodic Grade Report                                                                                     |
+-----------------+---------------------------------------------------------------------------------------------------------------+
| CANVAS_API      | https://unomaha.instructure.com                                                                               |
+-----------------+---------------------------------------------------------------------------------------------------------------+
| DEVELOPER_KEY   | <Developer Key From Canvas>                                                                                   |
+-----------------+---------------------------------------------------------------------------------------------------------------+
| MESSAGE_HEADER  | This message is to inform you of your standing in select CBA classes.  As of right now, your score in         |
+-----------------+---------------------------------------------------------------------------------------------------------------+
| MESSAGE_FOOTER  | This score is based on the current data in Canvas.  Please contact your instructor if you have any questions. |
+-----------------+---------------------------------------------------------------------------------------------------------------+
| COURSE_LIST     | https://docs.google.com/spreadsheets/d/1r1byAiO_6KhUSyJcVXAvTqOP0Uqw9eyQi-AIIDU7WBY/edit#gid=0                |
+-----------------+---------------------------------------------------------------------------------------------------------------+
| DANGER_SHEET    | https://docs.google.com/spreadsheets/d/1lM-bomPSIGyYm0Myt-T2KQXIYsAcZB3jUrUlci5H_Gk/edit#gid=0                |
+-----------------+---------------------------------------------------------------------------------------------------------------+
| EMAIL_STORAGE   | https://docs.google.com/spreadsheets/d/13tdBNFECF-6nMxAMhCFwbg2m1PB1iicYNpxRopJskZM/edit#gid=0                |
+-----------------+---------------------------------------------------------------------------------------------------------------+
The function main could be setup as a time driven trigger inside of the G Suite Developer Hub (the script must be run manually by the user once to authorize the OAuth scopes).  As a result this information can be sent on a period schedule (e.g. monthly).

For help, please contact Ben Smith <bosmith@unomaha.edu>.

*/


var extractDataFromCanvas = {};
extractDataFromCanvas.developerKey = "";
extractDataFromCanvas.canvasAPI = "";
extractDataFromCanvas.dangerSpreadsheet = "";
extractDataFromCanvas.courseList = "";
extractDataFromCanvas.domain = "";
extractDataFromCanvas.threshold = 100.0;


extractDataFromCanvas.loadSettings = function () {
    var scriptProperties = PropertiesService.getScriptProperties();
    extractDataFromCanvas.developerKey = "Bearer " + scriptProperties.getProperty('DEVELOPER_KEY');
    extractDataFromCanvas.canvasAPI = scriptProperties.getProperty('CANVAS_API');
    extractDataFromCanvas.dangerSpreadsheet = scriptProperties.getProperty('DANGER_SHEET');
    extractDataFromCanvas.courseList = scriptProperties.getProperty('COURSE_LIST');
    extractDataFromCanvas.domain = scriptProperties.getProperty('DOMAIN');
    createContent.message = scriptProperties.getProperty('MESSAGE_HEADER');
    createContent.footer = scriptProperties.getProperty('MESSAGE_FOOTER');
    createContent.subject = scriptProperties.getProperty('MESSAGE_SUBJECT');
    newEmailClass.sheetStorageURL = scriptProperties.getProperty('EMAIL_STORAGE');
    extractDataFromCanvas.threshold = parseFloat(scriptProperties.getProperty('THRESHOLD'));
    if (isNaN(extractDataFromCanvas.threshold)) {
        extractDataFromCanvas.threshold = 100.0;
    }
};


extractDataFromCanvas.getCourses = function () {
    var url = extractDataFromCanvas.canvasAPI + "/api/v1/courses?state[]=available&per_page=500";
    return extractDataFromCanvas.extractFromCanvas(url, []);
};

extractDataFromCanvas.findCourses = function (data) {
    var courseURLS = [];
    for (var i = 0; i < data.length; i++) {
        if (data[i].course_code) {
            var course = data[i].course_code.toString().substr(0, 8);
            if (loadCourseList.couseList.indexOf(course) > -1) {
                courseURLS.push({
                    course: course,
                    id: data[i].id
                });
            }
        }
    }
    return courseURLS;
};

extractDataFromCanvas.extractFromCanvas = function (apiCall, data) {
    var developerKey = extractDataFromCanvas.developerKey;
    var cData = data || [];
    var headers = {
        Authorization: developerKey
    };
    var options = {
        headers: headers
    };
    var response = UrlFetchApp.fetch(apiCall, options);
    var responseheaders = response.getHeaders();
    var jsondata = response.getContentText();
    var parsedata = JSON.parse(jsondata);
    if (Array.isArray(parsedata)) {
        cData = cData.concat(parsedata);
    }
    var links = linkheaderparser.parseLinkHeader(responseheaders.Link);
    if (links.next) {
        return extractDataFromCanvas.extractFromCanvas(links.next.href, cData);
    } else {
        return cData;
    }
};

extractDataFromCanvas.loopCourses = function (courses) {
    var data = [];
    for (var i = 0; i < courses.length; i++) {
        var course = courses[i].course;
        var id = courses[i].id;
        var url = extractDataFromCanvas.canvasAPI + "/api/v1/courses/" + id.toString() + "/enrollments?per_page=500&type[]=StudentEnrollment&state[]=active";
        var enrollments = extractDataFromCanvas.extractFromCanvas(url, []);
        if (Array.isArray(enrollments)) {
            data = data.concat(enrollments);
        }
    }
    return data;
};

var loadCourseList = {};
loadCourseList.couseList = [];
loadCourseList.couseDes = [];

loadCourseList.LoadCourseFile = function () {
    var ss = SpreadsheetApp.openByUrl(extractDataFromCanvas.courseList);
    if (ss !== false) {
        var sheets = ss.getSheets();
        if (sheets.length > 0) {
            var firstSheet = sheets[0];
            var data = firstSheet.getDataRange().getValues();
            for (var i = 0; i < data.length; i++) {
                if (data[i][0].toString().trim().length > 0) {
                    loadCourseList.couseList.push(data[i][0].toString().trim());
                    var courseMes = "";
                    if (data[i].length > 1){
                        if (data[i][1].toString().trim().length > 0) {
                            courseMes = data[i][1].toString().trim();
                        }
                    }
                    loadCourseList.couseDes.push(courseMes);
                }
            }
        }
    }
    return ss;
};

var updateSpreadsheetReport = {};

updateSpreadsheetReport.createSpreadsheet = function () {
    var header = [
        ['Date','Student', 'Course', 'E-Mail', 'Current Score', 'Current Grade']
    ];
    var ss = SpreadsheetApp.openByUrl(extractDataFromCanvas.dangerSpreadsheet);
    if (ss !== false) {
        var sheet = ss.getActiveSheet();
        sheet.activate();
        if (sheet.getLastRow() === 0) {
            sheet.getRange(1, 1, 1, header[0].length).setValues(header);
        }
        return sheet;
    }
    return ss;
};

updateSpreadsheetReport.addContentToSpreadSheet = function (sheet, contentData) {
    if (contentData.length > 0) {
        if (contentData[0].length > 0) {
            var setActiveRow = sheet.getLastRow() + 1;
            var range = sheet.getRange(setActiveRow, 1, contentData.length, contentData[0].length);
            range.setValues(contentData);
        }
    }
};


var createContent = {};
createContent.message = "";
createContent.footer = "";
createContent.subject = "";

createContent.createMessages = function (data) {
    var sheet = updateSpreadsheetReport.createSpreadsheet();
    var contentToWrite = [];
    var today = new Date();
    var dd = String(today.getDate());
    var mm = String(today.getMonth() + 1);
    var yyyy = today.getFullYear();
    var formattedDate = mm + '/' + dd + '/' + yyyy;
    for (var i = 0; i < data.length; i++) {
        // Based on Enrollments API

        if (data[i].grades && data[i].user && data[i].sis_course_id) {

            var course = data[i].sis_course_id.toString().substr(0, 8).trim();
            var name = data[i].user.name;
            var email = data[i].user.login_id;
            var score = data[i].grades.current_score;

            if (typeof score !== 'undefined' && score !== null && typeof name !== 'undefined' && name !== null && typeof email !== 'undefined' && email !== null) {
                if (data[i].grades.current_grade) {
                    var cgrade = data[i].grades.current_grade.toString().trim();
                    var cgrademes = "Your current grade is a(n) \"" + cgrade + ".\" ";
                } else {
                    var cgrademes = "";
                    var cgrade = "";
                }
                var message = name + ",\r\n\r\n" + createContent.message + " " + course + " is " + score + ". " + cgrademes + createContent.footer;
                if (loadCourseList.couseList.indexOf(course) > -1){
                    var mesPos = loadCourseList.couseList.indexOf(course);
                    var message = message + "\r\n\r\n" + loadCourseList.couseDes[mesPos];
                }
                if (score <= extractDataFromCanvas.threshold) {
                    var content = [formattedDate, name, data[i].sis_course_id.toString(), email, score, cgrade];
                    contentToWrite.push(content);
                }
                var emailSubject = createContent.subject + " - " + course;

                if (newEmailClass.quota > 0) {
                    newEmailClass.sendEmail(emailSubject, email, message);
                    newEmailClass.quota = newEmailClass.quota - 1;
                } else {
                    newEmailClass.writeToSheet(emailSubject, email, message);
                }
            }
        }
    }
    updateSpreadsheetReport.addContentToSpreadSheet(sheet, contentToWrite);
    newEmailClass.saveToSheet();
};

newEmailClass = {};
newEmailClass.quota = MailApp.getRemainingDailyQuota();
newEmailClass.sheetStorage = false;
newEmailClass.sheetStorageURL = "";
newEmailClass.storageRows = [];

newEmailClass.sendEmail = function (subjectLine, email, emailContent) {
    if (extractDataFromCanvas.domain.length > 0 && email.length > 0 && email.indexOf("@") === -1) {
        email += "@" + extractDataFromCanvas.domain;
    }
    if (emailContent.length > 0 && validateEmail.validate(email)) {
        subjectLine = subjectLine || "(No subject)";
        //Disable while testing
        MailApp.sendEmail(email, subjectLine, emailContent);
    }
};

newEmailClass.openSheet = function () {
    var ss = SpreadsheetApp.openByUrl(newEmailClass.sheetStorageURL);
    if (ss !== false) {
        var sheets = ss.getSheets();
        if (sheets.length > 0) {
            newEmailClass.sheetStorage = sheets[0];
        }
    }
};

newEmailClass.writeToSheet = function (subjectLine, email, emailContent) {
    if (newEmailClass.sheetStorage === false) {
        newEmailClass.openSheet();
        newEmailClass.createTrigger();
    }
    if (newEmailClass.sheetStorage !== false) {
        var datatowrite = [subjectLine, email, emailContent];
        newEmailClass.storageRows.push(datatowrite)
    }
};

newEmailClass.saveToSheet = function () {
    if (newEmailClass.sheetStorage !== false && newEmailClass.storageRows.length > 0) {
        if (newEmailClass.storageRows[0].length > 0) {
            var range = newEmailClass.sheetStorage.getRange((newEmailClass.sheetStorage.getLastRow() + 1), 1, newEmailClass.storageRows.length, newEmailClass.storageRows[0].length);
            range.setValues(newEmailClass.storageRows);
            newEmailClass.storageRows = [];
        }
    }
};

newEmailClass.createTrigger = function () {
    ScriptApp.newTrigger("cleanOutEmailList")
        .timeBased()
        .after(1500 * 60 * 1000)
        .create();
};


newEmailClass.readSheet = function () {
    var createNewTrigger = false;
    if (newEmailClass.sheetStorage === false) {
        newEmailClass.openSheet();
    }
    if (newEmailClass.sheetStorage !== false) {
        var numberofrows = 0;
        var data = newEmailClass.sheetStorage.getDataRange().getValues();
        for (var i = 0; i < data.length; i++) {
            if (newEmailClass.quota > 0) {
                numberofrows = numberofrows + 1;
            } else if (data[i].length > 2) {
                createNewTrigger = true;
                break;
            } else {
                numberofrows = numberofrows + 1;
            }
            if (data[i].length > 2 && newEmailClass.quota > 0) {
                var subject = data[i][0].toString();
                var email = data[i][1].toString();
                var body = data[i][2].toString();
                newEmailClass.sendEmail(subject, email, body);
                newEmailClass.quota = newEmailClass.quota - 1;
            }
        }
        if (numberofrows > 0) {
            newEmailClass.sheetStorage.deleteRows(1, numberofrows);
        }
    }
    if (createNewTrigger) {
        newEmailClass.createTrigger();
    }
};

function main() {
    extractDataFromCanvas.loadSettings();
    loadCourseList.LoadCourseFile();
    var courseList = extractDataFromCanvas.getCourses();
    var coursestoRun = extractDataFromCanvas.findCourses(courseList);
    var enrollements = extractDataFromCanvas.loopCourses(coursestoRun);
    createContent.createMessages(enrollements);
}

function cleanOutEmailList() {
    extractDataFromCanvas.loadSettings();
    newEmailClass.readSheet();
}