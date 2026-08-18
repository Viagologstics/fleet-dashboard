function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Vehicle Log & Dashboard Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getVehicleList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trip sheet");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var data = sheet.getRange("C2:C" + lastRow).getValues();
  var vehicles = [...new Set(data.flat().map(function(v) { return String(v).trim(); }).filter(function(v) { return v !== ""; }))];
  
  return vehicles.sort();
}

function submitNonRunData(vehicleNo, startDateStr, endDateStr, entryType, detailsInput) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trip sheet");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return "Error: No data rows found in Trip sheet.";

  var dataRange = sheet.getRange(2, 1, lastRow - 1, 39).getValues();
  var datesArray = getDatesBetween(startDateStr, endDateStr);
  var totalDaysCount = datesArray.length;
  var successCount = 0;

  datesArray.forEach(function(targetDate) {
    var targetRow = -1;
    
    for (var i = 0; i < dataRange.length; i++) {
      if (dataRange[i][2] == vehicleNo) {
        var rowDate = dataRange[i][1];
        var formattedRowDate = "";
        
        if (rowDate instanceof Date) {
          formattedRowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (rowDate) {
          formattedRowDate = String(rowDate).trim();
        }

        if (formattedRowDate === targetDate) {
          targetRow = i + 2; 
          break;
        }
      }
    }
    
    if (targetRow === -1) {
      var lastMatchingRow = -1;
      for (var i = 0; i < dataRange.length; i++) {
        if (dataRange[i][2] == vehicleNo) {
          lastMatchingRow = i + 2;
        }
      }
      targetRow = lastMatchingRow; 
    }

    if (targetRow !== -1) {
      if (entryType === 'Non-Run') {
        var existingNonRun = sheet.getRange(targetRow, 30).getValue();
        var entryText = "*" + targetDate + " (Days: 1) - Remarks: " + detailsInput.trim() + "*";
        var newVal = existingNonRun ? existingNonRun + ", " + entryText : entryText;
        sheet.getRange(targetRow, 30).setValue(newVal);
      } else {
        if (detailsInput && detailsInput.trim() !== "") {
          var existingIncident = sheet.getRange(targetRow, 39).getValue();
          var formattedLog = "[" + targetDate + "] " + entryType + " " + detailsInput.trim();
          var newIncidentVal = existingIncident ? existingIncident + " | " + formattedLog : formattedLog;
          sheet.getRange(targetRow, 39).setValue(newIncidentVal);
        }
      }
      successCount++;
    }
  });

  if (successCount > 0) {
    return "Success: Log updated for " + vehicleNo + " across " + totalDaysCount + " date entry(ies).";
  } else {
    return "Error: Vehicle number " + vehicleNo + " not found in Trip sheet Column C.";
  }
}

// Generates high-level macro analytics for the default Director view
function getFleetMacroData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trip sheet");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { totalFleetSize: 0, totalNonRuns: 0, totalIncidents: 0, fleetSummary: [] };

  var dataRange = sheet.getRange(2, 1, lastRow - 1, 39).getValues();
  var vehicleMap = {};
  var totalNonRuns = 0;
  var totalIncidents = 0;

  dataRange.forEach(function(row) {
    var vehicle = String(row[2] || "").trim(); // Column C
    if (!vehicle) return;

    if (!vehicleMap[vehicle]) {
      vehicleMap[vehicle] = { vehicle: vehicle, nonRunCount: 0, incidentCount: 0 };
    }

    var nonRunVal = String(row[29] || "").trim(); // Column AD
    var incidentVal = String(row[38] || "").trim(); // Column AM

    if (nonRunVal) {
      var nonRunChunks = nonRunVal.split(/[,|]/);
      nonRunChunks.forEach(function(chunk) {
        if (chunk.replace(/\*/g, '').trim()) {
          vehicleMap[vehicle].nonRunCount++;
          totalNonRuns++;
        }
      });
    }

    if (incidentVal) {
      var incidentChunks = incidentVal.split('|');
      incidentChunks.forEach(function(chunk) {
        if (chunk.trim()) {
          vehicleMap[vehicle].incidentCount++;
          totalIncidents++;
        }
      });
    }
  });

  var fleetSummaryArray = [];
  for (var v in vehicleMap) {
    fleetSummaryArray.push(vehicleMap[v]);
  }

  // Sort by highest incidents/non-runs descending to surface problem areas immediately
  fleetSummaryArray.sort(function(a, b) {
    return (b.nonRunCount + b.incidentCount) - (a.nonRunCount + a.incidentCount);
  });

  return {
    totalFleetSize: fleetSummaryArray.length,
    totalNonRuns: totalNonRuns,
    totalIncidents: totalIncidents,
    fleetSummary: fleetSummaryArray
  };
}

// Drill-down data for a specific vehicle selected by management
function getVehicleDashboardData(vehicleNo) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trip sheet");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { nonRunCount: 0, incidentCount: 0, logs: [] };

  var dataRange = sheet.getRange(2, 1, lastRow - 1, 39).getValues();
  var parsedLogs = [];
  var nonRunCount = 0;
  var incidentCount = 0;

  for (var i = 0; i < dataRange.length; i++) {
    if (dataRange[i][2] == vehicleNo) {
      var nonRunVal = String(dataRange[i][29] || "").trim(); // Column AD
      var incidentVal = String(dataRange[i][38] || "").trim(); // Column AM

      if (nonRunVal) {
        var nonRunChunks = nonRunVal.split(/[,|]/);
        nonRunChunks.forEach(function(chunk) {
          var cleanChunk = chunk.replace(/\*/g, '').trim();
          if (cleanChunk) {
            nonRunCount++;
            parsedLogs.push({
              date: extractDateFromString(cleanChunk) || "General",
              type: "Non-Run",
              text: cleanChunk
            });
          }
        });
      }

      if (incidentVal) {
        var incidentChunks = incidentVal.split('|');
        incidentChunks.forEach(function(chunk) {
          var cleanChunk = chunk.trim();
          if (cleanChunk) {
            incidentCount++;
            parsedLogs.push({
              date: extractDateFromString(cleanChunk) || "General",
              type: "Incident / Remark",
              text: cleanChunk
            });
          }
        });
      }
    }
  }

  parsedLogs.sort(function(a, b) {
    return b.date.localeCompare(a.date);
  });

  return {
    nonRunCount: nonRunCount,
    incidentCount: incidentCount,
    logs: parsedLogs
  };
}

function extractDateFromString(str) {
  var match = str.match(/\d{4}-\d{2}-\d{2}/);
  return match ? match[0] : "";
}

function getDatesBetween(startDateString, endDateString) {
  var arr = [];
  var dt = new Date(startDateString);
  var endDt = endDateString ? new Date(endDateString) : new Date(startDateString);
  
  if (endDt < dt) endDt = dt; 

  while (dt <= endDt) {
    var month = '' + (dt.getMonth() + 1);
    var day = '' + dt.getDate();
    var year = dt.getFullYear();

    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;

    arr.push([year, month, day].join('-'));
    dt.setDate(dt.getDate() + 1);
  }
  return arr;
}
