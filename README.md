Automated Early Warning System to Notify Students and Advisors: A Google Apps Script to Extract Current Scores from Canvas

This script uses the Canvas API to extract the current scores from a list of courses (defined in COURSE_LIST) in a given an integer term (defined as the script property TERM).  It notifies the students of their current score in the these courses over email.  If the student's score is below a threshold (defined as the script property THRESHOLD), the student's name, email, score, and course are written to a spreadsheet (defined in DANGER_SHEET).  EMAIL_STORAGE is used to store emails for future sending if the script executing user is over the daily email quota.

The COURSE_LIST file allows you to send additional information to students based on the course.  Therefore, the course should be defined in the first column of the COURSE_LIST sheet and additional information should be provided in the second column. 

The script must have the following script properties defined (set under File->Project Properties->Script Properties within the Apps Scripts editor): 

Script Property | Example
---------------------|------------- 
THRESHOLD   | 75
DOMAIN  | unomaha.edu 
MESSAGE_SUBJECT | CBA Periodic Grade Report 
CANVAS_API      | https://unomaha.instructure.com
DEVELOPER_KEY   | Developer Key From Canvas
MESSAGE_HEADER  | This message is to inform you of your standing in select CBA classes.  As of right now, your score in   
MESSAGE_FOOTER  | This score is based on the current data in Canvas.  Please contact your instructor if you have any questions. 
TERM            | 0 for disabled, or 1, 2, ... 
COURSE_LIST     | https://docs.google.com/spreadsheets/d/1r1byAiO_6KhUSyJcVXAvTqOP0Uqw9eyQi-AIIDU7WBY/edit#gid=0
DANGER_SHEET    | https://docs.google.com/spreadsheets/d/1lM-bomPSIGyYm0Myt-T2KQXIYsAcZB3jUrUlci5H_Gk/edit#gid=0 
EMAIL_STORAGE   | https://docs.google.com/spreadsheets/d/13tdBNFECF-6nMxAMhCFwbg2m1PB1iicYNpxRopJskZM/edit#gid=0
FROM_NAME       | UNO Automated Early Warning System
REPLY_TO        | unocbaadvising@unomaha.edu 

The function main could be setup as a time driven trigger inside of the G Suite Developer Hub (the script must be run manually by the user once to authorize the OAuth scopes).  As a result this information can be sent on a period schedule (e.g. monthly).

For help, please contact Ben Smith <bosmith@unomaha.edu>.
