const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const notifyAboutTaskStatuses = ["Ongoing", "Delayed"];
const urlRegex = /((([A-Za-z]{3,9}:(?:\/\/)?)(?:[-;:&=\+\$,\w]+@)?[A-Za-z0-9.-]+|(?:www.|[-;:&=\+\$,\w]+@)[A-Za-z0-9.-]+)((?:\/[\+~%\/.\w-_]*)?\??(?:[-\+=&;%@.\w_]*)#?(?:[\w]*))?)/;


const errorEmailId = "paras@rivant.in"; // Will send the mail to this email if an error occured
const requiredMailIds = [];  // Mail id to add to send mail


const actualCalendaSheetName = "7 Day Assignment - Actual";
const alloacatedCalendaSheetName = "7 Day Assignment - Allocated";
const employeeEmailIdIndex = 0;
const completedOnDateIndex = 11;
const ss = SpreadsheetApp.getActiveSpreadsheet();


// Configuration Variables
const subject = "Your Personalised Morning Summary"
// For Employee's Sheet
const employeesSheetName = "Company Employees";
const emailAddressIndex = 7;
const employeeNameIndex = 1;
// For Task's Sheet
const companyNameIndex = 0;
const taskDomainIndex = 1;
const taskSummaryIndex = 2;
const taskBriefIndex = 3;
const allocatedHoursIndex = 4;
const taskEmailAddressIndex = 6;
const assignedOnIndex = 7;
const deadlineDateIndex = 8;
const inchargeIndex = 9;
const taskStatusIndex = 10;
