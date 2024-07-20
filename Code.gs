const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const notifyAboutTaskStatuses = ["Ongoing", "Delayed"];

// Configuration Variables Index
// For Employee's Sheet
var employeesSheetName = "Company Employees";
var emailAddressIndex = 7;
var employeeNameIndex = 1;
// For Task's Sheet
var companyNameIndex = 0;
var taskNameIndex = 1;
var taskDescIndex = 2;
var allocatedHoursIndex = 3;
var taskEmailAddressIndex = 5;
var assignedOnIndex = 6;
var deadlineDateIndex = 7;
var inchargeIndex = 8;
var taskStatusIndex = 9;

function sendGeneratedEmail() {
  const todayDate = new Date();
  var currMonth = months[todayDate.getMonth()];
  var activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var employeeSheetIndex = activeSpreadSheet.findIndex((sheet) => sheet.getName() === employeesSheetName);
  var taskSheetIndex = activeSpreadSheet.findIndex((sheet) => sheet.getName() === currMonth);
  var employeeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[employeeSheetIndex];
  var taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[taskSheetIndex];
  var employeeDataRange = employeeSheet.getDataRange();
  var employeeData = employeeDataRange.getValues();   // Gets Data from Employee Sheet
  var taskDataRange = taskSheet.getDataRange();
  var taskData = taskDataRange.getValues();   // Gets Data from Task Sheet

  // Loops Through All Active Employees
  for (var i = 0; i < employeeData.length; i++) {

    var emailAddress = employeeData[i][emailAddressIndex];
    var employeeName = employeeData[i][employeeNameIndex];
    var taskCardsDisplay = "";
    var taskAssigned = false;
    var taskCount = 0;

    // Loops Through Task Data
    for (var j = 1; j < taskData.length; j++) {
      
      if (taskData[j][taskEmailAddressIndex] === emailAddress) {
        var taskStatus = taskData[j][taskStatusIndex];
        if (notifyAboutTaskStatuses.includes(taskStatus)) {
          taskAssigned = true;
          var companyName = taskData[j][companyNameIndex];
          var taskName = taskData[j][taskNameIndex];
          var assignedOn = taskData[j][assignedOnIndex].toLocaleDateString();
          var allocatedHours = taskData[j][allocatedHoursIndex];
          var taskDuration = Math.ceil(allocatedHours/3);
          var taskDesc = taskData[j][taskDescIndex];
          var deadlineDate = taskData[j][deadlineDateIndex];
          var deadline = deadlineDate.toLocaleDateString();
          var daysRemaining = Math.floor((Math.abs(deadlineDate - todayDate)/1000)/(60*60*24));
          var isDeadlineCrossed = todayDate > deadlineDate
          var incharge = taskData[j][inchargeIndex];

          taskCount++;  // Increment No of Task Assigned
          taskCardsDisplay = taskCardsDisplay + " " + taskCards(companyName,taskName,assignedOn,allocatedHours,taskDuration,deadline,daysRemaining,incharge, taskDesc, isDeadlineCrossed);  // Create and Concatenate Task Card Component
        }
      }
    }

    // Code To send emails
    var subject = "Morning Summary";
    var htmlBody = generateEmail(employeeName, taskCardsDisplay, taskCount, taskAssigned);
      
    if (emailAddress) {
      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: htmlBody
      });
    }

  }
}

// Function To Generate HTML Component of Email
function generateEmail(employeeName, taskCardsDisplay, taskCount, taskAssigned) {
  return `${compulsoryHtml(employeeName)} ${taskHtml(taskCardsDisplay, taskCount, taskAssigned)}`;
}

// HTML Component that should be send to everyone
function compulsoryHtml(employeeName) {
  return `
  <div
      style="
        font-family: Arial, sans-serif;
        max-width: 600px;
        margin: auto;
        padding: 20px;
        background-color: #121212;
        color: #fff;
        border-radius: 8px;
      "
    >
      <h1 style="text-align: center">Good Morning <br />${employeeName}</h1>
      <p style="text-align: center">
        Before you get your day started, please mark your daily attendance by
        clicking the button below.
      </p>
      <div style="text-align: center; margin: 35px 0">
        <a
          href="https://dash.rvnt.in/attendance"
          style="
            padding: 10px 30px;
            background-color: #f6e8d0;
            color: #1a1a1a;
            text-decoration: none;
            border-radius: 5px;
            font-weight: 600;
          "
          >Daily Attendance</a
        >
      </div>
      <div style="text-align: center; color: #4b4842">
        <p>
          If you are planning to take a leave today, please inform HR as well as
          your Department Head about the same.
        </p>
        <p>Click the below links to write an email to HR about your leave.</p>
        <div style="text-align: center">
          <a
            href="https://mail.google.com/a/rivant.in/mail/?view=cm&to=hr@rivant.in&su=Leave%20of%20Absence%20-%20%5BYour%20Name%5D%20%28Start%20Date%20-%20End%20Date%29&body=Dear%20Sir%2FMa%27am%2C%0A%0AThis%20email%20is%20to%20inform%20you%20that%20I%20will%20be%20taking%20a%20leave%20of%20absence%20from%20today%2C%20%5BDate%5D%2C%20until%20%5BEnd%20Date%5D%2C%20returning%20to%20work%20on%20%5BReturn%20Date%5D.%0A%0A%5BOptional%3A%20Briefly%20state%20the%20reason%20for%20your%20leave%2C%20but%20this%20is%20not%20mandatory.%5D%0A%0AI%20have%20completed%20%5Bbriefly%20mention%20any%20urgent%20tasks%20you%20have%20finished%5D%20and%20have%20informed%20%5Bcolleague%27s%20name%5D%20about%20my%20ongoing%20projects.%20%20%5BColleague%27s%20name%5D%20is%20familiar%20with%20the%20work%20and%20can%20be%20contacted%20at%20%5Bcolleague%27s%20email%20address%5D%20if%20anything%20urgent%20arises.%0A%0AI%20will%20be%20checking%20my%20email%20periodically%20in%20case%20of%20emergencies.%20You%20can%20reach%20me%20at%20%5Byour%20work%20email%20address%5D%20or%20%5Byour%20phone%20number%5D%20if%20absolutely%20necessary.%0A%0AThank%20you%20for%20your%20understanding.%0A%0ASincerely%2C%0A%0A%5BYour%20Name%5D"
            style="color: #2d5173; margin: 0 10px"
            >personal leave</a
          >
          |
          <a
            href="https://mail.google.com/a/rivant.in/mail/?view=cm&to=hr@rivant.in&su=Medical%20Leave%20of%20Absence%20-%20%5BYour%20Name%5D%20%28Start%20Date%20-%20End%20Date%29&body=Dear%20Sir%2FMa%27am%2C%0A%0AThis%20email%20is%20to%20inform%20you%20that%20I%20will%20be%20taking%20a%20medical%20leave%20of%20absence%20from%20today%2C%20%5BDate%5D%2C%20until%20%5BEnd%20Date%5D%2C%20returning%20to%20work%20on%20%5BReturn%20Date%5D.%0A%0APlease%20note%20that%20this%20is%20a%20medical%20leave.%0A%0A%5BOptional%3A%20You%20can%20choose%20to%20disclose%20more%20information%20here%2C%20but%20you%20are%20not%20obligated%20to%20do%20so.%20Examples%3A%20%22I%20will%20be%20undergoing%20a%20medical%20procedure%22%20or%20%22I%20am%20not%20feeling%20well%20enough%20to%20work.%22%20%5D%0A%0AI%20have%20completed%20%5Bbriefly%20mention%20any%20urgent%20tasks%20you%20have%20finished%5D%20and%20have%20informed%20%5Bcolleague%27s%20name%5D%20about%20my%20ongoing%20projects.%20%20%5BColleague%27s%20name%5D%20is%20familiar%20with%20the%20work%20and%20can%20be%20contacted%20at%20%5Bcolleague%27s%20email%20address%5D%20if%20anything%20urgent%20arises.%0A%0AI%20will%20be%20checking%20my%20email%20periodically%20in%20case%20of%20emergencies.%20You%20can%20reach%20me%20at%20%5Byour%20work%20email%20address%5D%20or%20%5Byour%20phone%20number%5D%20if%20absolutely%20necessary.%0A%0AThank%20you%20for%20your%20understanding.%0A%0ASincerely%2C%0A%0A%5BYour%20Name%5D"
            style="color: #2d5173; margin: 0 10px"
            >medical leave</a
          >
        </div>
      </div>
      <hr
        style="
          border-color: #fff;
          margin: 30px 0;
          height: 1px;
          background-color: #fff;
        "
      />`;
}

// HTML Component For Task Display
function taskHtml(taskCardsDisplay, taskCount, taskAssigned) {
  return taskAssigned ? `
  <div>
        <h2 style="text-align: center">
          Here are the tasks you will be working on today.
        </h2>
        <p style="text-align: center; color: #9e9a9a">
          (You have ${taskCount} tasks assigned today)
        </p>
      </div>
      ${taskCardsDisplay}
    </div>` : `
   <div>
        <h2 style="text-align: center">
          There No task assigned to you today.
        </h2>
    </div>
    </div> `;
}

// HTML Component of Task Card
function taskCards(companyName,taskName,assignedOn,allocatedHours,taskDuration,deadline,daysRemaining,incharge,taskDesc,isDeadlineCrossed) {

  if (daysRemaining === 0 && isDeadlineCrossed) {
    var msg = "due today";
    var color = "#2D5173";
  } else if (daysRemaining === 0 && !isDeadlineCrossed) {
    var msg = "due tomorrow";
    var color = "#4B4842";
  } else if (isDeadlineCrossed) {
    var msg = `${daysRemaining} day overdue`;
    var color = "#823D3F";
  } else {
    var msg = `${daysRemaining} days remaining`;
    var color = "#4B4842";
  }

  return `
  <div
        style="
          background-color: #2b282c;
          padding: 20px;
          border-radius: 8px;
          margin: 30px auto 0 auto;
          width: 85%;
        "
      >
        <table style="width: 90%; margin: auto">
          <tr>
            <td>
              <p font-size: 16px; margin: 0 0 5px 0">
                ${companyName}
              </p>
            </td>
            <td rowspan="2" style="text-align: end">
              <a
                href="${taskDesc}"
                style="
                  padding: 5px 15px;
                  background-color: #f6e8d0;
                  color: #1a1a1a;
                  text-decoration: none;
                  border-radius: 5px;
                  font-weight: 600;
                  font-size: 13px;
                  margin: auto;
                "
                >View Details</a
              >
            </td>
          </tr>
          <tr>
            <td>
              <p
                style="
                  font-size: 14px;
                  color: #9e9a9a;
                  margin: 0;
                  width: 6rem;
                "
              >
                ${taskName}
              </p>
            </td>
          </tr>
        </table>
        <div style="width: 90%; margin: 25px auto 0 auto; color: #9e9a9a">
          <p style="margin: 0 0 4px 0">
            <span class="icon"></span>Assigned On: ${assignedOn}
          </p>
          <p style="margin: 0 0 4px 0">
            <span class="icon"></span>Allocated Hours: ${allocatedHours}
          </p>
          <p style="margin: 0 0 4px 0">
            <span class="icon"></span>Task Duration: ${taskDuration} days
          </p>
          <p style="margin: 0 0 4px 0">
            <span class="icon"></span>Deadline: ${deadline}
            <span style="color: ${color}; font-size: 11px"
              >(${msg})</span
            >
          </p>
          <p style="margin: 0 0 4px 0">
            <span class="icon"></span>Incharge:
            <a href="mailto:${incharge.split(" ")[0].toLowerCase()}@rivant.in" style="color: #4891d4">${incharge}</a>
          </p>
        </div>
      </div>`;
}

// It will make a Trigger The sendGeneratedEmail function on regular interval
function createTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('sendGeneratedEmail')
           .timeBased()
           .atHour(8)
           .everyDays(1)
           .create();
}
