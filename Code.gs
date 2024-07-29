const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const notifyAboutTaskStatuses = ["Ongoing", "Delayed"];
const urlRegex = /((([A-Za-z]{3,9}:(?:\/\/)?)(?:[-;:&=\+\$,\w]+@)?[A-Za-z0-9.-]+|(?:www.|[-;:&=\+\$,\w]+@)[A-Za-z0-9.-]+)((?:\/[\+~%\/.\w-_]*)?\??(?:[-\+=&;%@.\w_]*)#?(?:[\w]*))?)/;


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

function sendGeneratedEmail() {
  const todayDate = new Date();

  // Do not send email if it's sunday
  if (todayDate.getDay() === 0) {
    return;
  }

  var currMonth = months[todayDate.getMonth()];
  var employeeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(employeesSheetName);
  var taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(currMonth);
  var employeeDataRange = employeeSheet.getDataRange();
  var employeeData = employeeDataRange.getValues();   // Gets Data from Employee Sheet
  var taskDataRange = taskSheet.getDataRange();
  var taskData = taskDataRange.getValues();   // Gets Data from Task Sheet

  // Loops Through All Active Employees
  for (var i = 0; i < employeeData.length; i++) {

    var emailAddress = employeeData[i][emailAddressIndex];
    var employeeName = employeeData[i][employeeNameIndex];
    var tasks = [];
    var taskAssigned = false;

    // Loops Through Task Data
    for (var j = 1; j < taskData.length; j++) {
      
      if (taskData[j][taskEmailAddressIndex] === emailAddress) {
        var taskStatus = taskData[j][taskStatusIndex];
        if (notifyAboutTaskStatuses.includes(taskStatus)) { 
          taskAssigned = true;
          var companyName = taskData[j][companyNameIndex];
          var taskDomain = taskData[j][taskDomainIndex];
          var taskBrief = taskData[j][taskBriefIndex];
          var assignedOn = taskData[j][assignedOnIndex].toLocaleDateString();
          var allocatedHours = taskData[j][allocatedHoursIndex];
          var taskDuration = Math.ceil(allocatedHours/3);
          var taskSummary = taskData[j][taskSummaryIndex];
          var deadlineDate = taskData[j][deadlineDateIndex];
          var deadline = deadlineDate.toLocaleDateString();
          var daysRemaining = Math.floor((Math.abs(deadlineDate - todayDate)/1000)/(60*60*24));
          var isDeadlineCrossed = todayDate > deadlineDate
          var incharge = taskData[j][inchargeIndex];
          var taskBriefExist = urlRegex.test(taskBrief);

          var task = {
            companyName,
            taskDomain,
            assignedOn,
            taskBriefExist,
            taskBrief,
            allocatedHours,
            taskDuration,
            taskSummary,
            deadline,
            daysRemaining,
            incharge,
            isDeadlineCrossed
          }

          tasks.push(task);
        }
      }
    }

    // Code To send emails
    var htmlBody = generateEmail(employeeName, taskAssigned ,tasks);
      
    if (emailAddress) {
      MailApp.sendEmail({
        to: emailAddress,
        subject,
        htmlBody: htmlBody
      });
    }

  }
}

function generateEmail(employeeName, taskAssigned, tasks) {
  const htmlTemplate = HtmlService.createTemplateFromFile('index');
  htmlTemplate.employeeName = employeeName;
  htmlTemplate.taskAssigned = taskAssigned;
  htmlTemplate.tasks = tasks;
  return htmlTemplate.evaluate().getContent();
}

// It will make a Trigger The sendGeneratedEmail function on regular interval
function activateSendGeneratedEmail() {
  ScriptApp.newTrigger('sendGeneratedEmail')
           .timeBased()
           .everyDays(1)
           .atHour(5)
           .create();
}

function deactivateSendGeneratedEmail() {
  const triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendGeneratedEmail") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
