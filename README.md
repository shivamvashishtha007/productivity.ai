/**
 * Logic 0: Track Last Processed Row for Each Logic
 * Stores last processed row to optimize execution time.
 */
function trackLastProcessedRow(sheetName, logicNumber, lastRow) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataProcessingTracker");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("DataProcessingTracker");
    sheet.appendRow(["Sheet Name", "Last Row - Logic 1", "Last Row - Logic 2", "Last Row - Logic 3", "Last Data Sync", "Error Logs"]);
  }
 
  var data = sheet.getDataRange().getValues();
  var rowFound = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === sheetName) {
      data[i][logicNumber] = lastRow;
      data[i][4] = new Date();
      rowFound = true;
      break;
    }
  }
  if (!rowFound) {
    sheet.appendRow([sheetName, logicNumber === 1 ? lastRow : "", logicNumber === 2 ? lastRow : "", logicNumber === 3 ? lastRow : "", new Date(), ""]);
  } else {
    sheet.getRange(2, 1, data.length - 1, data[0].length).setValues(data.slice(1));
  }
  Logger.log(`✅ Logic 0: Updated last processed row for ${sheetName} - Logic ${logicNumber} at row ${lastRow}`);
}








/**
 * Logic 0: Retrieve Last Processed Row for a Sheet
 */
function getLastProcessedRow(sheetName, logicNumber) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataProcessingTracker");
  if (!sheet) {
    Logger.log(`⚠️ No DataProcessingTracker sheet found. Defaulting to row 1.`);
    return 1; // If tracker does not exist, start from the first row
  }
 
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === sheetName) {
      Logger.log(`🔄 Logic 0: Retrieved last processed row ${data[i][logicNumber] || 1} for ${sheetName} - Logic ${logicNumber}`);
      return data[i][logicNumber] || 1; // Default to row 1 if no record
    }
  }
  Logger.log(`⚠️ No record found for ${sheetName}. Defaulting to row 1.`);
  return 1;
}








/**
 * Get Last Non-Empty Row in a Given Column
 */
function getLastNonEmptyRow(sheet, columnIndex) {
  var data = sheet.getRange(1, columnIndex, sheet.getLastRow()).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== "") {
      return i + 1; // Convert 0-based index to row number
    }
  }
  return 1; // Default to first row if all values are empty
}




/**
 * Logic 1: Process Data from 'Source Inputs' to Geo-Specific Sheets (Optimized)
 */
function processSourceToGeoSheets() {
    Logger.log("🚀 Starting Logic 1: Processing data from 'Source Inputs' to Geo-specific sheets...");
   
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName("Source Inputs");
    if (!sourceSheet) {
        Logger.log("❌ Source Inputs sheet not found. Exiting Logic 1.");
        return;
    }
   
    var sourceData = sourceSheet.getDataRange().getValues();
    var lastRow = getLastNonEmptyRow(sourceSheet, 3); // ✅ Ensure last row is correctly identified
    var startRow = getLastProcessedRow("Source Inputs", 1);
    var headerRow = 1;




    let geoDataMap = {}; // Stores existing data for each geo-sheet to reduce `getValues()` calls
    let dataToAppend = {}; // Stores data to be written in batch mode




    for (var i = Math.max(headerRow, startRow); i < lastRow; i++) {
        var row = sourceData[i];
        if (row[9] === "Processed" || row[2] === "") {
          Logger.log(`⏭️ Skipping row ${i + 1}: Already processed or missing Game Name.`);
            continue;
        }




        // Assign Serial Number and Date
        if (!row[1]) row[1] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy");




        var gameName = row[2].trim();
        var gplayPackage = row[4] ? row[4].trim() : "";
        var priority = row[5] ? row[5].trim() : "";
        var geoList = row[7] ? row[7].split(",").map(geo => geo.trim()) : [];
        var blogTypes = row[6] ? row[6].split(",").map(blog => blog.trim()) : [];
        var contentMap = {};




        var contentTypeEntries = row[8] ? row[8].split("\n") : [];
        for (var entry of contentTypeEntries) {
            var parts = entry.split("- ");
            if (parts.length === 2) {
                var type = parts[0].trim();
                var applicableGeos = parts[1].split(",").map(geo => geo.trim());
                applicableGeos.forEach(geo => contentMap[geo] = type);
            }
        }




        if (!gameName || geoList.length === 0 || blogTypes.length === 0) continue;
        for (var j = 0; j < geoList.length; j++) {
            var geoSheetName = geoList[j].trim();
            var contentType = contentMap[geoSheetName] || "Original";
            var geoSheet = ss.getSheetByName(geoSheetName);
            if (!geoSheet) {
                Logger.log(`⚠️ Geo sheet '${geoSheetName}' not found. Skipping.`);
                continue;
            }
   
            if (!geoDataMap[geoSheetName]) {
                geoDataMap[geoSheetName] = new Set(geoSheet.getDataRange().getValues().map(row => row[2] + "|" + row[4] + "|" + row[6]));
            }




            if (!dataToAppend[geoSheetName]) {
                dataToAppend[geoSheetName] = [];
            }




            // ✅ Loop through blog types (keeps serial numbers increasing correctly)
            for (var k = 0; k < blogTypes.length; k++) {
                var blogType = blogTypes[k].trim();
                var key = gameName + "|" + gplayPackage + "|" + blogType;




                if (geoDataMap[geoSheetName].has(key)) {
                    Logger.log(`⏭️ Skipping duplicate entry for ${gameName} in ${geoSheetName}.`);
                    continue;
                }




                dataToAppend[geoSheetName].push([
                    "",  // Empty column for Serial Number (Column A)
                    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy"),  // Date (Column B)
                    gameName,  // Game Name (Column C)
                    row[3],  // Source (Column D)
                    gplayPackage,  // Gplay Package (Column E)
                    priority,  // Priority (Column F)
                    blogType,  // Blog Type (Column G)
                    "", "", "", "",  // Empty columns (Columns H to K)
                    contentType  // Content Type (Column L)
                ]);
            }
        }




        // ✅ Assign Serial Number in Column A if missing
        if (!row[0]) {
            row[0] = i + 1; // Ensure continuity of serial numbers
        }
        row[9] = "Processed"; // Mark as processed
    }




    for (var sheetName in dataToAppend) {
        var geoSheet = ss.getSheetByName(sheetName);
        if (geoSheet) {
            var startRow = getLastNonEmptyRow(geoSheet, 3) + 1; // ✅ Correct last row before appending
            geoSheet.getRange(startRow, 1, dataToAppend[sheetName].length, 12).setValues(dataToAppend[sheetName]);
            trackLastProcessedRow(sheetName, 1, startRow + dataToAppend[sheetName].length - 1);
        }
    }




    trackLastProcessedRow("Source Inputs", 1, lastRow);
    sourceSheet.getDataRange().setValues(sourceData);
    Logger.log("✅ Logic 1 Execution Completed Successfully.");
    SpreadsheetApp.flush();
}




/**
 * ✅ Get Last Non-Empty Row in a Given Column
 */
function getLastNonEmptyRow(sheet, columnIndex) {
    var data = sheet.getRange(1, columnIndex, sheet.getLastRow()).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
        if (data[i][0] !== "") {
            return i + 1; // Convert 0-based index to row number
        }
    }
    return 1; // Default to first row if all values are empty
}




/**
 * 🚀 Initial Logic 2 Start - Adds a custom menu to the Google Sheets UI for easy execution of Logic 2.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Controls')
    .addItem('Run Logic 2 for Specific Sheet', 'showSheetSelectionDialog')
    .addToUi();
}




/**
 * ✅ Displays a dialog box for the user to select a Geowise Sheet to run Logic 2.
 */
function showSheetSelectionDialog() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Run Logic 2',
    'Enter the Geowise Sheet name (e.g., EN, PL, PTBR, IT, FR, etc.) or type "ALL" to process all sheets:',
    ui.ButtonSet.OK_CANCEL
  );




  if (response.getSelectedButton() == ui.Button.OK) {
    var sheetName = response.getResponseText().trim().toUpperCase();
    var allowedSheets = ["EN", "PL", "PTBR", "IT", "FR", "DE", "ES", "RU", "TR", "JA", "KO", "TW", "VI", "ID", "TH", "AR"];




    if (sheetName === "ALL" || allowedSheets.includes(sheetName)) {
      Logger.log(`▶️ Running Logic 2 for: ${sheetName}`);
      processGeoToWriterSheets(sheetName === "ALL" ? null : sheetName);
      ui.alert(`✅ Logic 2 has been executed for: ${sheetName}`);
    } else {
      ui.alert(`❌ Invalid sheet name. Please enter a valid Geowise Sheet name or "ALL".`);
    }
  }
}


/**
 * 🚀 Initial Logic 2 Start - Adds a custom menu to the Google Sheets UI for easy execution of Logic 2.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Controls')
    .addItem('Run Logic 2 for Specific Sheet', 'showSheetSelectionDialog')
    .addToUi();
}


/**
 * ✅ Displays a dialog box for the user to select a Geowise Sheet to run Logic 2.
 */
function showSheetSelectionDialog() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Run Logic 2',
    'Enter the Geowise Sheet name (e.g., EN, PL, PTBR, IT, FR, etc.) or type "ALL" to process all sheets:',
    ui.ButtonSet.OK_CANCEL
  );


  if (response.getSelectedButton() == ui.Button.OK) {
    var sheetName = response.getResponseText().trim().toUpperCase();
    var allowedSheets = ["EN", "PL", "PTBR", "IT", "FR", "DE", "ES", "RU", "TR", "JA", "KO", "TW", "VI", "ID", "TH", "AR"];


    if (sheetName === "ALL" || allowedSheets.includes(sheetName)) {
      Logger.log(`▶️ Running Logic 2 for: ${sheetName}`);
      processGeoToWriterSheets(sheetName === "ALL" ? null : sheetName);
      ui.alert(`✅ Logic 2 has been executed for: ${sheetName}`);
    } else {
      ui.alert(`❌ Invalid sheet name. Please enter a valid Geowise Sheet name or "ALL".`);
    }
  }
}


/**
 * 🚀 Optimized Logic 2: Transfers data from Geowise Sheets to External Writer Workbooks
 */
function processGeoToWriterSheets(geoSheetFilter = null) {
    Logger.log("🚀 Starting Optimized Logic 2...");


    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var writerConfigSheet = ss.getSheetByName("Writer Configuration Sheet");


    if (!writerConfigSheet) {
        Logger.log("❌ Writer Configuration Sheet not found. Exiting Logic 2.");
        return;
    }


    // ✅ Load writer configuration
    var writerConfigData = writerConfigSheet.getDataRange().getValues();
    var writerMap = {};
    var writerSheetNames = {};


    for (var i = 1; i < writerConfigData.length; i++) {
        var writerName = writerConfigData[i][1]?.trim().toLowerCase();
        var writerSheetURL = writerConfigData[i][3]?.trim();
        var writerSubsheetName = writerConfigData[i][2]?.trim();


        if (writerName && writerSheetURL && writerSubsheetName) {
            writerMap[writerName] = { url: writerSheetURL, sheetName: writerSubsheetName };
        }
    }


    var allowedSheets = ["EN", "PL", "PTBR", "IT", "FR", "DE", "ES", "RU", "TR", "JA", "KO", "TW", "VI", "ID", "TH", "AR"];
    var sheetsToProcess = geoSheetFilter
        ? (allowedSheets.includes(geoSheetFilter) ? [geoSheetFilter] : [])
        : ss.getSheets().map(s => s.getName()).filter(name => allowedSheets.includes(name));


    sheetsToProcess.forEach(sheetName => {
        if (geoSheetFilter && !allowedSheets.includes(geoSheetFilter)) {
            Logger.log(`⚠️ Invalid Geowise sheet '${geoSheetFilter}'. Skipping.`);
            return;
        }


        var geoSheet = ss.getSheetByName(sheetName);
        if (!geoSheet) {
            Logger.log(`⚠️ Geowise Sheet '${sheetName}' not found. Skipping.`);
            return;
        }


        var lastProcessedRow = getLastProcessedRow(sheetName, 2);
        var geoData = geoSheet.getDataRange().getValues();
        var newProcessedRow = lastProcessedRow;


        let writerDataToAppend = {};
        let writerSheetsCache = {};  
        let writerExistingEntries = {};
        let timestampsToUpdate = [];


        for (var i = lastProcessedRow; i < geoData.length; i++) {
            var row = geoData[i];
            var writer = row[7] ? row[7].trim().toLowerCase() : "";


            if (!writer || !writerMap[writer]?.url) {
                continue;
            }


            if (!writerSheetsCache[writer]) {
                if (!writerMap[writer]?.url || !writerMap[writer]?.sheetName) {
                    Logger.log(`❌ Missing URL or sheet name for writer '${writer}'. Skipping.`);
                    continue;
                }


                try {
                    var writerDoc = SpreadsheetApp.openByUrl(writerMap[writer].url);
                    var writerSheet = writerDoc.getSheetByName(writerMap[writer].sheetName);


                    if (!writerSheet) {
                        Logger.log(`❌ Writer sheet '${writerMap[writer].sheetName}' not found for '${writer}'. Skipping.`);
                        continue;
                    }


                    writerSheetsCache[writer] = writerSheet;
                    var writerData = writerSheet.getDataRange().getValues();
                    writerExistingEntries[writer] = new Set(
                        writerData.slice(1).map(r => `${r[1]?.trim()}|${r[3]?.trim()}|${r[5]?.trim()}`)
                    );
                } catch (e) {
                    Logger.log(`❌ Error opening Writer Sheet for '${writer}': ${e.message}. Skipping.`);
                    continue;
                }
            }


            var duplicateKey = `${row[2]?.trim()}|${row[4]?.trim()}|${row[6]?.trim()}`;
            if (writerExistingEntries[writer].has(duplicateKey)) {
                continue;
            }


            if (!writerDataToAppend[writer]) writerDataToAppend[writer] = [];


            let dataRow = [
                row[2].trim(),
                row[3].trim(),
                row[4].trim(),
                mapPriorityValue(row[5]),
                geoSheet.getRange(i + 1, 7).getDisplayValue(),
                new Date(),
                "", "",
                row[11] ? row[11].trim() : ""
            ];


            writerDataToAppend[writer].push(dataRow);
            timestampsToUpdate.push([dataRow[5]]); // Column G value (Assigned Timestamp) to Column I
            newProcessedRow = i + 1;
        }


        for (var writer in writerDataToAppend) {
            var writerSheet = writerSheetsCache[writer];


            if (writerSheet && writerDataToAppend[writer].length > 0) {
                var data = writerSheet.getDataRange().getValues();
var startRow = data.length;


// Ensure we find the actual last non-empty row in column B (Game Name / Topic)
while (startRow > 0 && (!data[startRow - 1][1] || data[startRow - 1][1].toString().trim() === "")) {
    startRow--;
}
startRow += 1; // Move to the next row for new data


var dataCheck = writerSheet.getRange(startRow, 2).getValue().trim(); // Check Column B (Game Name)


if (!dataCheck) {
    startRow = startRow; // If the last row is empty, overwrite it
} else {
    startRow += 1; // Otherwise, add a new row
}




                writerSheet.getRange(startRow, 2, writerDataToAppend[writer].length, writerDataToAppend[writer][0].length)
                    .setValues(writerDataToAppend[writer]);
            }
        }


        if (timestampsToUpdate.length > 0) {
            geoSheet.getRange(lastProcessedRow + 1, 9, timestampsToUpdate.length, 1).setValues(timestampsToUpdate);
        }


        SpreadsheetApp.flush();
        trackLastProcessedRow(sheetName, 2, newProcessedRow);
        Logger.log(`✅ Optimized Logic 2 Execution Completed for '${sheetName}'.`);
    });
}


/**
 * ✅ Maps Priority values from Geowise Sheet to Writer Sheet dropdown options.
 */
function mapPriorityValue(priority) {
    var priorityMapping = {
        "Weekly Focused": "IP0",
        "Partial Focused": "IP0",
        "IP0": "IP0",
        "CPI / DAU": "IP1",
        "IP1": "IP1",
        "B2B": "IP1",
        "IP3": "IP3",
        "Regular": "IP2",
        "IP2": "IP2"
    };
    return priorityMapping[priority.trim()] || "IP0";
}


//Logic 3 starts


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation Controls')
    .addItem('Run Logic 2 for Specific Sheet', 'showSheetSelectionDialog')
    .addItem('Run Logic 3 for Specific Writer', 'showGeoSheetSelectionDialog')
    .addToUi();
}


/**
 * Displays a dialog box for the user to enter the Geowise Sheet name.
 */
function showGeoSheetSelectionDialog() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Run Logic 3',
    'Enter the Geowise Sheet name (e.g., EN, PL, PTBR, IT, FR, etc.) or type "ALL" to process all sheets:',
    ui.ButtonSet.OK_CANCEL
  );


  if (response.getSelectedButton() == ui.Button.OK) {
    var sheetName = response.getResponseText().trim().toUpperCase();
    var allowedSheets = ["EN", "PL", "PTBR", "IT", "FR", "DE", "ES", "RU", "TR", "JA", "KO", "TW", "VI", "ID", "TH", "AR"];


    if (sheetName === "ALL" || allowedSheets.includes(sheetName)) {
      Logger.log(`▶️ Running Logic 3 for: ${sheetName}`);
      processWriterToGeoSheets(sheetName === "ALL" ? null : sheetName);
      ui.alert(`✅ Logic 3 has been executed for: ${sheetName}`);
    } else {
      ui.alert(`❌ Invalid sheet name. Please enter a valid Geowise Sheet name.`);
    }
  }
}


/**
 * 🚀 Optimized Logic 3: Transfers data from Writer Sheets back to Geowise Sheets.
 */
function processWriterToGeoSheets(geoSheetName) {
  var allowedSheets = ["EN", "PL", "PTBR", "IT", "FR", "DE", "ES", "RU", "TR", "JA", "KO", "TW", "VI", "ID", "TH", "AR"];
  var sheetsToProcess = geoSheetName ? [geoSheetName] : allowedSheets;


  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var writerConfigSheet = ss.getSheetByName("Writer Configuration Sheet");
  if (!writerConfigSheet) {
    Logger.log("❌ Writer Configuration Sheet not found. Exiting Logic 3.");
    return;
  }


  var writerConfigData = writerConfigSheet.getDataRange().getValues();
  var writerSheetMap = {};
  for (var i = 1; i < writerConfigData.length; i++) {
    var writerName = writerConfigData[i][1]?.trim();
    var writerSheetName = writerConfigData[i][2]?.trim();
    var writerSheetUrl = writerConfigData[i][3]?.trim();
    if (writerName && writerSheetName && writerSheetUrl) {
      writerSheetMap[writerName.toLowerCase()] = { sheetName: writerSheetName, sheetUrl: writerSheetUrl };
    }
  }


  sheetsToProcess.forEach(function(sheet) {
    Logger.log(`🚀 Starting Optimized Logic 3 for Geowise Sheet: '${sheet}'`);


    var geoSheet = ss.getSheetByName(sheet);
    if (!geoSheet) {
      Logger.log(`❌ Geowise Sheet '${sheet}' not found. Skipping.`);
      return;
    }


    var lastProcessedRow = getLastProcessedRow(sheet, 3);
    var geoData = geoSheet.getDataRange().getValues();
    var totalRows = geoData.length;
    if (totalRows <= 1) {
      Logger.log(`⚠️ No data found in '${sheet}'. Skipping.`);
      return;
    }


    var updates = [];
    var newProcessedRow = lastProcessedRow;
    var writerSheetsCache = {};


    for (var i = Math.max(1, lastProcessedRow); i < totalRows; i++) {
      var row = geoData[i];
      var writerName = row[7]?.trim().toLowerCase();
      var assignedDate = row[9]?.toString().trim();
      var deliveredDate = row[10]?.toString().trim();


      if (!writerName || (assignedDate && deliveredDate)) {
        continue;
      }


      if (!writerSheetMap[writerName]) {
        Logger.log(`⚠️ Writer '${writerName}' not found in Writer Configuration Sheet. Skipping row ${i + 1}.`);
        continue;
      }


      if (!writerSheetsCache[writerName]) {
        var writerDoc = SpreadsheetApp.openByUrl(writerSheetMap[writerName].sheetUrl);
        var writerSheet = writerDoc.getSheetByName(writerSheetMap[writerName].sheetName);
        if (!writerSheet) {
          Logger.log(`❌ Error accessing Writer Sheet for '${writerName}'. Skipping row ${i + 1}.`);
          continue;
        }
        writerSheetsCache[writerName] = writerSheet.getDataRange().getValues();
      }


      var writerData = writerSheetsCache[writerName];
      var writerMap = {};
      for (var j = 1; j < writerData.length; j++) {
        var key = `${writerData[j][1]?.trim()}|${writerData[j][3]?.trim()}|${writerData[j][5]?.trim()}`;
        writerMap[key] = [writerData[j][7], writerData[j][8]];
      }


      var geoKey = `${row[2]?.trim()}|${row[4]?.trim()}|${row[6]?.trim()}`;
      if (writerMap[geoKey]) {
        var [writerAssignedDate, writerDeliveredDate] = writerMap[geoKey];


        if (!assignedDate && writerAssignedDate) row[9] = writerAssignedDate;
        if (!deliveredDate && writerDeliveredDate) row[10] = writerDeliveredDate;


        updates.push([row[9], row[10]]);
        newProcessedRow = i + 1;
      } else {
        updates.push([row[9], row[10]]);
      }
    }


    if (updates.length > 0) {
      geoSheet.getRange(lastProcessedRow + 1, 10, updates.length, 2).setValues(updates);
      Logger.log(`✅ Successfully updated ${updates.length} rows in '${sheet}'.`);
    } else {
      Logger.log(`⚠️ No updates made in '${sheet}'. All rows were already filled.`);
    }


    trackLastProcessedRow(sheet, 3, newProcessedRow);
    Logger.log(`✅ Optimized Logic 3 Execution Completed for '${sheet}'.`);
  });
}



