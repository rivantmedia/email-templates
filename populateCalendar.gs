const calendaSheetName = "7 Day Assignment";
const employeeEmailIdIndex = 0;
const completedOnDateIndex = 11;
const linkFormula = `=QUERY('Company Employees'!A1:H, "select Col8")`;

function populateCalendar() {
  const todayDate = new Date();
  var currMonth = months[todayDate.getMonth()];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName(currMonth);
  const calendarSheet = ss.getSheetByName(calendaSheetName);
  var taskData = taskSheet.getDataRange().getValues();
  var calendarData = calendarSheet.getDataRange().getValues();
  
  // Loop through the calendar sheet
  for(var i = 1; i < calendarData.length; i++) {
    for(var j = 1; j < calendarData[i].length; j++) {
      var employeeEmailId = calendarData[i][0];
      var date = calendarData[0][j];
      var tasks = 0;
      
      // Loop through the task data
      for(var k = 1; k < taskData.length; k++) {
        var assignedOnDate = taskData[k][assignedOnIndex];
        var deadlineDate = taskData[k][deadlineDateIndex];
        var completedOnDate = taskData[k][completedOnDateIndex];

        if (!completedOnDate) {
          if (taskData[k][taskEmailAddressIndex] == employeeEmailId && assignedOnDate <= date && deadlineDate >= date) {
            tasks++;
          }
        } else {
          if (taskData[k][taskEmailAddressIndex] == employeeEmailId && assignedOnDate <= date && completedOnDate >= date) {
            tasks++;
          }
        }
      }

      // Join tasks and set in the calendar cell
      calendarData[i][j] = tasks;
    }
    calendarData[i][0] = "";
  }
  calendarData[1][0] = linkFormula;
  calendarData[0][1] = "=C1-1";
  calendarData[0][2] = "=D1-1";
  calendarData[0][3] = "=E1-1";
  calendarData[0][4] = "=TODAY()";
  calendarData[0][5] = "=E1+1";
  calendarData[0][6] = "=F1+1";
  calendarData[0][7] = "=G1+1";

  calendarSheet.getDataRange().setValues(calendarData);
}

function generateTriggger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  ScriptApp.newTrigger('populateCalendar')
           .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
           .onChange()
           .create();
  
  ScriptApp.newTrigger('populateCalendar')
           .timeBased()
           .everyDays(1)
           .atHour(0)
           .create();
}
