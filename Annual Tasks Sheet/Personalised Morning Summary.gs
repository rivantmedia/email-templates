function sendPersonalisedGeneratedEmail() {
	sendGeneratedEmail(requiredMailIds);
}

function sendGeneratedEmail(mailIds = []) {
	const todayDate = new Date();
	const errorOccurredIds = [];

	// Do not send email if it's sunday
	if (todayDate.getDay() === 0) {
		return;
	}

	var currMonth = months[todayDate.getMonth()];
	var employeeSheet = ss.getSheetByName(employeesSheetName);
	var taskSheet = ss.getSheetByName(currMonth);
	var employeeDataRange = employeeSheet.getDataRange();
	var employeeData = employeeDataRange.getValues();
	var taskDataRange = taskSheet.getDataRange();
	var taskData = taskDataRange.getValues();

	if (mailIds.length != 0) {
		employeeData = employeeData.filter((employee) => mailIds.includes(employee[emailAddressIndex]));
	}

	// Loops Through All Active Employees
	for (var i = 0; i < employeeData.length; i++) {
		try {
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
						var assignedOn = taskData[j][assignedOnIndex].toLocaleDateString("en-GB");
						var allocatedHours = taskData[j][allocatedHoursIndex];
						var taskDuration = Math.ceil(allocatedHours / 3);
						var taskSummary = taskData[j][taskSummaryIndex];
						var deadlineDate = taskData[j][deadlineDateIndex];
						var deadline = deadlineDate.toLocaleDateString("en-GB");
						var daysRemaining = Math.floor(Math.abs(deadlineDate - todayDate) / 1000 / (60 * 60 * 24));
						var isDeadlineCrossed = todayDate > deadlineDate;
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
						};

						tasks.push(task);
					}
				}
			}

			// Code To send emails
			var htmlBody = generateEmail(employeeName, taskAssigned, tasks);

			if (emailAddress) {
				MailApp.sendEmail({
					to: emailAddress,
					subject,
					htmlBody: htmlBody
				});
			}
		} catch (err) {
			const errorDetails = {
				emailAddress,
				error: err.message
			};
			errorOccurredIds.push(errorDetails);
		}
	}
	if (errorOccurredIds != 0) {
		htmlBody = generateErrorEmail(errorOccurredIds);
		MailApp.sendEmail({
			to: errorEmailId,
			subject: "An Error Occurred During Execution of sendGeneratedEmail",
			htmlBody: htmlBody
		});
	}
}

function generateEmail(employeeName, taskAssigned, tasks) {
	const htmlTemplate = HtmlService.createTemplateFromFile("morning-summary-email-template");
	htmlTemplate.employeeName = employeeName;
	htmlTemplate.taskAssigned = taskAssigned;
	htmlTemplate.tasks = tasks;
	return htmlTemplate.evaluate().getContent();
}

function generateErrorEmail(errorOccurredIds) {
	const htmlTemplate = HtmlService.createTemplateFromFile("error-logs-email-template");
	htmlTemplate.data = errorOccurredIds;
	return htmlTemplate.evaluate().getContent();
}

// It will make a Trigger The sendGeneratedEmail function on regular interval
function activateSendGeneratedEmail() {
	ScriptApp.newTrigger("sendGeneratedEmail").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(5).create();
	ScriptApp.newTrigger("sendGeneratedEmail").timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(5).create();
	ScriptApp.newTrigger("sendGeneratedEmail").timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(5).create();
	ScriptApp.newTrigger("sendGeneratedEmail").timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(5).create();
	ScriptApp.newTrigger("sendGeneratedEmail").timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(5).create();
	ScriptApp.newTrigger("sendGeneratedEmail").timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(5).create();
}

function deactivateSendGeneratedEmail() {
	const triggers = ScriptApp.getProjectTriggers();
	for (var i = 0; i < triggers.length; i++) {
		if (triggers[i].getHandlerFunction() === "sendGeneratedEmail") {
			ScriptApp.deleteTrigger(triggers[i]);
		}
	}
}
