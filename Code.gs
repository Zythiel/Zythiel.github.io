function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Project Entry')
    .addItem('Open Entry Form', 'showEntryForm')
    .addItem('Edit Existing Entries', 'showEditForm')
    .addItem('Manage People/Roles', 'showPeopleManager')
    .addItem('Manage Escalation Options', 'showEscalationManager')
    .addToUi();
}

function showEntryForm() {
  var html = HtmlService.createHtmlOutputFromFile('EntryForm')
    .setWidth(800)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Project Task Entry');
}

function showEditForm() {
  var html = HtmlService.createHtmlOutputFromFile('EditForm')
    .setWidth(800)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Edit Existing Entries');
}

function showPeopleManager() {
  var html = HtmlService.createHtmlOutputFromFile('PeopleManager')
    .setWidth(500)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manage People/Roles');
}

function showEscalationManager() {
  var html = HtmlService.createHtmlOutputFromFile('EscalationManager')
    .setWidth(500)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manage Escalation Options');
}

function getBudgetDataSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Budget Data');

  if (!sheet) {
    sheet = ss.insertSheet('Budget Data');
    // Add header row
    sheet.getRange('A1:J1').setValues([[
      'Milestone',
      'Task',
      'Project Year',
      'Start Date',
      'End Date',
      'Person/Role',
      'Rate',
      'Hours/Quantity',
      'Escalation',
      'Notes'
    ]]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function getPeopleRolesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('People_Roles');

  if (!sheet) {
    sheet = ss.insertSheet('People_Roles');
    sheet.getRange('A1:C1').setValues([['Name', 'Rate/Unit Cost', 'Type']]).setFontWeight('bold');
  }

  return sheet;
}

function getAllPeopleRoles() {
  var sheet = getPeopleRolesSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var results = [];

  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) {
      results.push({
        id: i,
        person: data[i][0],
        rate: data[i][1],
        type: data[i][2] || 'Person/Role' // Default to Person/Role for backward compatibility
      });
    }
  }

  return results;
}

function addPersonRole(person, rate, type) {
  var sheet = getPeopleRolesSheet();
  sheet.appendRow([person, parseFloat(rate), type || 'Person/Role']);
  return { success: true };
}

function deletePersonRole(id) {
  var sheet = getPeopleRolesSheet();
  var rowToDelete = id + 2;
  sheet.deleteRow(rowToDelete);
  return { success: true };
}

function getEscalationSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Escalation_Options');

  if (!sheet) {
    sheet = ss.insertSheet('Escalation_Options');
    sheet.getRange('A1').setValue('Percentage').setFontWeight('bold');
    sheet.getRange('A2:A5').setValues([[0], [2], [3], [5]]);
  }

  return sheet;
}

function getProjectSettingsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Project_Settings');

  if (!sheet) {
    sheet = ss.insertSheet('Project_Settings');
    sheet.getRange('A1:B1').setValues([['Setting', 'Value']]).setFontWeight('bold');
  }

  return sheet;
}

function getSavedEscalation() {
  var sheet = getProjectSettingsSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'LastEscalation') {
      return data[i][1];
    }
  }

  return null;
}

function saveEscalation(escalationPercentage) {
  var sheet = getProjectSettingsSheet();
  var data = sheet.getDataRange().getValues();
  var found = false;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'LastEscalation') {
      sheet.getRange(i + 1, 2).setValue(escalationPercentage);
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.appendRow(['LastEscalation', escalationPercentage]);
  }

  return { success: true };
}

function getAllEscalationOptions() {
  var sheet = getEscalationSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var results = [];

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== '' && data[i][0] !== null) {
      results.push({
        id: i,
        percentage: data[i][0]
      });
    }
  }

  return results;
}

function addEscalationOption(percentage) {
  var sheet = getEscalationSheet();
  sheet.appendRow([parseFloat(percentage)]);
  return { success: true };
}

function deleteEscalationOption(id) {
  var sheet = getEscalationSheet();
  var rowToDelete = id + 2;
  sheet.deleteRow(rowToDelete);
  return { success: true };
}

function submitFormData(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getBudgetDataSheet();

    var selectedPeople = formData.selectedPeople;
    var tasks = formData.tasks;
    var milestone = formData.milestone;
    var escalationPercentage = formData.escalationPercentage;
    var notes = formData.notes;
    var differentHours = formData.differentHours;
    var customDates = formData.customDates || false;

    saveEscalation(escalationPercentage);

    var rows = [];

    // Each task now has its own projectYears array
    for (var t = 0; t < tasks.length; t++) {
      var taskItem = tasks[t];
      var taskProjectYears = taskItem.projectYears || [];

      for (var y = 0; y < taskProjectYears.length; y++) {
        var year = taskProjectYears[y];
        var yearOffset = year - 1;

        var adjustedStartDate, adjustedEndDate;

        if (customDates && taskItem.yearDates && taskItem.yearDates[year]) {
          // Use custom dates for this year
          adjustedStartDate = new Date(taskItem.yearDates[year].startDate);
          adjustedEndDate = new Date(taskItem.yearDates[year].endDate);
        } else {
          // Use base dates and advance by year offset
          var baseStartDate = new Date(taskItem.startDate);
          var baseEndDate = new Date(taskItem.endDate);

          adjustedStartDate = new Date(baseStartDate);
          adjustedStartDate.setFullYear(baseStartDate.getFullYear() + yearOffset);

          adjustedEndDate = new Date(baseEndDate);
          adjustedEndDate.setFullYear(baseEndDate.getFullYear() + yearOffset);
        }

        for (var p = 0; p < selectedPeople.length; p++) {
          var pr = selectedPeople[p];
          var escalationMultiplier = Math.pow(1 + (escalationPercentage / 100), yearOffset);
          var escalatedRate = parseFloat(pr.rate) * escalationMultiplier;
          var roundedRate = Math.round(escalatedRate * 100) / 100;

          // Get hours for this person/task/year combination
          var hours = '';
          if (pr.hoursData && pr.hoursData[taskItem.id]) {
            if (differentHours && pr.hoursData[taskItem.id][year] !== undefined) {
              hours = pr.hoursData[taskItem.id][year];
            } else if (!differentHours && pr.hoursData[taskItem.id]['all'] !== undefined) {
              hours = pr.hoursData[taskItem.id]['all'];
            }
          }

          rows.push([
            milestone,
            taskItem.task,
            year,
            adjustedStartDate,
            adjustedEndDate,
            pr.person,
            roundedRate,
            hours,
            escalationPercentage + '%',
            notes
          ]);
        }
      }
    }

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { success: true, rowsAdded: rows.length };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function getAllUniqueMilestones() {
  var sheet = getBudgetDataSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var milestonesObj = {};

  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) {
      milestonesObj[data[i][0]] = true;
    }
  }

  var milestones = [];
  for (var key in milestonesObj) {
    milestones.push(key);
  }

  return milestones.sort();
}

function getTasksForMilestone(milestone) {
  var sheet = getBudgetDataSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var tasksObj = {};

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === milestone && data[i][1]) {
      tasksObj[data[i][1]] = true;
    }
  }

  var tasks = [];
  for (var key in tasksObj) {
    tasks.push(key);
  }

  return tasks.sort();
}

function searchEntries(milestone, task) {
  try {
    var sheet = getBudgetDataSheet();
    var lastRow = sheet.getLastRow();

    Logger.log("Searching for: '" + milestone + "' and '" + task + "'");
    Logger.log("Last row: " + lastRow);

    if (lastRow <= 1) {
      return { found: false, message: "No data in sheet" };
    }

    var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    var matchingRows = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowMilestone = String(row[0] || '').trim();
      var rowTask = String(row[1] || '').trim();
      var searchMilestone = String(milestone || '').trim();
      var searchTask = String(task || '').trim();

      if (rowMilestone === searchMilestone && rowTask === searchTask) {
        matchingRows.push({
          rowIndex: i + 2,
          milestone: row[0],
          task: row[1],
          projectYear: row[2],
          startDate: row[3],
          endDate: row[4],
          person: row[5],
          rate: row[6],
          escalation: row[7],
          notes: row[8]
        });
      }
    }

    Logger.log("Found " + matchingRows.length + " matching rows");

    if (matchingRows.length === 0) {
      return {
        found: false,
        message: "No matching entries found for '" + milestone + "' and '" + task + "'"
      };
    }

    var projectYearsObj = {};
    var peopleObj = {};

    for (var i = 0; i < matchingRows.length; i++) {
      projectYearsObj[matchingRows[i].projectYear] = true;
      peopleObj[matchingRows[i].person] = true;
    }

    var projectYears = [];
    for (var key in projectYearsObj) {
      projectYears.push(Number(key));
    }
    projectYears.sort(function(a, b) { return a - b; });

    var people = [];
    for (var key in peopleObj) {
      people.push(key);
    }

    var startDate = null;
    var endDate = null;

    for (var i = 0; i < matchingRows.length; i++) {
      if (matchingRows[i].projectYear === projectYears[0]) {
        startDate = matchingRows[i].startDate;
        endDate = matchingRows[i].endDate;
        break;
      }
    }

    var escalation = String(matchingRows[0].escalation || '0%');
    // If it's already a string with %, keep it
    // If it's a number like 0.03, convert to 3%
    if (escalation.indexOf('%') === -1) {
      // It's a decimal number, convert to percentage
      escalation = (parseFloat(escalation) * 100) + '%';
    }

    var notes = matchingRows[0].notes || '';

    var result = {
      found: true,
      rowCount: matchingRows.length,
      rowIndices: [],
      milestone: milestone,
      task: task,
      projectYears: projectYears,
      people: people,
      startDate: startDate,
      endDate: endDate,
      escalation: escalation,
      notes: notes
    };

    for (var i = 0; i < matchingRows.length; i++) {
      result.rowIndices.push(matchingRows[i].rowIndex);
    }

    Logger.log("Returning result with " + result.rowCount + " matches");

    return result;

  } catch (error) {
    Logger.log("Error in searchEntries: " + error.toString());
    return {
      found: false,
      message: "Error: " + error.toString()
    };
  }
}

function updateEntries(searchCriteria, updates) {
  try {
    var sheet = getBudgetDataSheet();

    // First get the existing data to find rows to delete
    var existingData = searchEntriesForUI(searchCriteria.milestone, searchCriteria.task);

    if (!existingData.found) {
      return { success: false, error: "No matching entries found" };
    }

    // Use the hoursMap from the form if provided, otherwise fall back to existing
    var hoursMap = updates.hoursMap || existingData.hoursMap || {};

    // Use the datesMap from the form if provided
    var datesMap = updates.datesMap || {};

    var rowIndices = existingData.rowIndices.slice();
    rowIndices.sort(function(a, b) { return b - a; });

    for (var i = 0; i < rowIndices.length; i++) {
      sheet.deleteRow(rowIndices[i]);
    }

    var escalationPercentage = parseFloat(updates.escalation.replace('%', ''));
    var rows = [];

    for (var y = 0; y < updates.projectYears.length; y++) {
      var year = updates.projectYears[y];
      var yearOffset = year - 1;

      var adjustedStartDate, adjustedEndDate;

      // Check if we have custom dates for this year
      if (datesMap[year] && datesMap[year].startDate && datesMap[year].endDate) {
        adjustedStartDate = new Date(datesMap[year].startDate);
        adjustedEndDate = new Date(datesMap[year].endDate);
      } else {
        // Fall back to base dates with year offset
        var baseStartDate = new Date(updates.startDate);
        var baseEndDate = new Date(updates.endDate);

        adjustedStartDate = new Date(baseStartDate);
        adjustedStartDate.setFullYear(baseStartDate.getFullYear() + yearOffset);

        adjustedEndDate = new Date(baseEndDate);
        adjustedEndDate.setFullYear(baseEndDate.getFullYear() + yearOffset);
      }

      for (var p = 0; p < updates.selectedPeople.length; p++) {
        var pr = updates.selectedPeople[p];
        var escalationMultiplier = Math.pow(1 + (escalationPercentage / 100), yearOffset);
        var escalatedRate = parseFloat(pr.rate) * escalationMultiplier;
        var roundedRate = Math.round(escalatedRate * 100) / 100;

        // Get hours from the hoursMap (either from form input or preserved from existing data)
        var hoursKey = pr.person + '|' + year;
        var hours = hoursMap[hoursKey];
        if (hours === undefined || hours === null) {
          hours = '';
        }

        rows.push([
          updates.milestone,
          updates.task,
          year,
          adjustedStartDate,
          adjustedEndDate,
          pr.person,
          roundedRate,
          hours,
          updates.escalation,
          updates.notes
        ]);
      }
    }

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return {
      success: true,
      deletedRows: rowIndices.length,
      addedRows: rows.length
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function searchEntriesForUI(milestone, task) {
  Logger.log("searchEntriesForUI called with: " + milestone + ", " + task);

  try {
    var sheet = getBudgetDataSheet();
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return { found: false, message: "No data in sheet" };
    }

    var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var matchingRows = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowMilestone = String(row[0] || '').trim();
      var rowTask = String(row[1] || '').trim();
      var searchMilestone = String(milestone || '').trim();
      var searchTask = String(task || '').trim();

      if (rowMilestone === searchMilestone && rowTask === searchTask) {
        matchingRows.push({
          rowIndex: i + 2,
          milestone: row[0],
          task: row[1],
          projectYear: row[2],
          startDate: row[3],
          endDate: row[4],
          person: row[5],
          rate: row[6],
          hours: row[7],
          escalation: row[8],
          notes: row[9]
        });
      }
    }

    Logger.log("Found " + matchingRows.length + " matching rows");

    if (matchingRows.length === 0) {
      return {
        found: false,
        message: "No matching entries found for '" + milestone + "' and '" + task + "'"
      };
    }

    var projectYearsObj = {};
    var peopleObj = {};
    var hoursMap = {}; // Store hours for each person/year combination
    var datesMap = {}; // Store dates for each year

    for (var i = 0; i < matchingRows.length; i++) {
      var row = matchingRows[i];
      projectYearsObj[row.projectYear] = true;
      peopleObj[row.person] = true;

      // Store hours for this person and year
      var hoursKey = row.person + '|' + row.projectYear;
      hoursMap[hoursKey] = row.hours || '';

      // Store dates for this year (only need one per year since all rows for a year have same dates)
      if (!datesMap[row.projectYear]) {
        datesMap[row.projectYear] = {
          startDate: row.startDate,
          endDate: row.endDate
        };
      }
    }

    var projectYears = [];
    for (var key in projectYearsObj) {
      projectYears.push(Number(key));
    }
    projectYears.sort(function(a, b) { return a - b; });

    var people = [];
    for (var key in peopleObj) {
      people.push(key);
    }

    // Get base dates from Year 1 (or first year)
    var startDate = null;
    var endDate = null;

    for (var i = 0; i < matchingRows.length; i++) {
      if (matchingRows[i].projectYear === projectYears[0]) {
        startDate = matchingRows[i].startDate;
        endDate = matchingRows[i].endDate;
        break;
      }
    }

    var escalation = String(matchingRows[0].escalation || '0%');
    if (escalation.indexOf('%') === -1) {
      escalation = (parseFloat(escalation) * 100) + '%';
    }

    var notes = matchingRows[0].notes || '';

    var result = {
      found: true,
      rowCount: matchingRows.length,
      rowIndices: [],
      milestone: milestone,
      task: task,
      projectYears: projectYears,
      people: people,
      hoursMap: hoursMap, // Include the hours data
      datesMap: datesMap, // Include the dates for each year
      startDate: startDate ? startDate.toString() : null,
      endDate: endDate ? endDate.toString() : null,
      escalation: escalation,
      notes: notes
    };

    for (var i = 0; i < matchingRows.length; i++) {
      result.rowIndices.push(matchingRows[i].rowIndex);
    }

    Logger.log("Returning result with " + result.rowCount + " matches");

    return result;

  } catch (error) {
    Logger.log("Error in searchEntriesForUI: " + error.toString());
    return {
      found: false,
      message: "Error: " + error.toString()
    };
  }
}
