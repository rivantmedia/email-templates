const actualCalendaSheetName = "7 Day Assignment - Actual";
const alloacatedCalendaSheetName = "7 Day Assignment - Allocated";
const employeeEmailIdIndex = 0;
const completedOnDateIndex = 11;
const ss = SpreadsheetApp.getActiveSpreadsheet();

function initiatePopulateCalendar() {
  populateData(actualCalendaSheetName);
  populateData(alloacatedCalendaSheetName);
}

function populateData(sheetName) {
  const todayDate = new Date();
  var currMonth = months[todayDate.getMonth()];
  const taskSheet = ss.getSheetByName(currMonth);
  var taskData = taskSheet.getDataRange().getValues();
  const calendarSheet = ss.getSheetByName(sheetName);
  const lastRow = calendarSheet.getLastRow();
  const lastColumn = calendarSheet.getLastColumn();
  const datesArray = calendarSheet.getRange(1,2,1,lastColumn-1).getValues();
  const employeeArray = calendarSheet.getRange(2,1,lastRow-1,1).getValues();
  var calendarData = calendarSheet.getRange(2,2,lastRow-1,lastColumn-1).getValues();

  // Loop through the calendar sheet
  for(var i = 0; i < calendarData.length; i++) {
    for(var j = 0; j < calendarData[i].length; j++) {
      var employeeEmailId = employeeArray[i][0];
      var date = datesArray[0][j];
      var tasks = 0;
      
      // Loop through the task data
      for(var k = 1; k < taskData.length; k++) {
        var assignedOnDate = taskData[k][assignedOnIndex];
        var deadlineDate = taskData[k][deadlineDateIndex];
        var completedOnDate = taskData[k][completedOnDateIndex];

        if (sheetName === actualCalendaSheetName) {
          if (!completedOnDate) {
            if (taskData[k][taskEmailAddressIndex] == employeeEmailId && assignedOnDate <= date && todayDate >= date) {
              tasks++;
            }
          } else {
            if (taskData[k][taskEmailAddressIndex] == employeeEmailId && assignedOnDate <= date && completedOnDate >= date) {
              tasks++;
            }
          }
        } else {
          if (taskData[k][taskEmailAddressIndex] == employeeEmailId && assignedOnDate <= date && deadlineDate >= date) {
            tasks++;
          }
        }
      }

      // Join tasks and set in the calendar cell
      calendarData[i][j] = tasks === 0 ? "" : tasks;
    }
  }

  calendarSheet.getRange(2,2,lastRow-1,lastColumn-1).setValues(calendarData);
}

function activateInitiatePopulateCalendar() {
  ScriptApp.newTrigger('initiatePopulateCalendar')
           .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
           .onChange()
           .create();
  
  ScriptApp.newTrigger('initiatePopulateCalendar')
           .timeBased()
           .everyDays(1)
           .atHour(0)
           .create();
}

function deactivateInitiatePopulateCalendar() {
  const triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "initiatePopulateCalendar") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
