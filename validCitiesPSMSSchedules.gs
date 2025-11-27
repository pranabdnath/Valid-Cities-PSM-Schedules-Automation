/*
  Script: validCitiesPSMSSchedules
  Purpose:
    - Import zipped CSV data from an email.
    - Filter records into North, South, and No-Date groups based on city and matching conditions.
    - Clean, write, and format data into dedicated sheets.
    - Apply 30-day Y/N flag logic.
    - Log all activities to ScriptLogs sheet.

  Note:
    Replace the subject line, folder IDs, and sheet names with your own references.
*/

function validCitiesPSMSSchedules() {
  const SHEET_NAME_NORTH = "Schedules-Noida";
  const SHEET_NAME_SOUTH = "Schedules-South";
  const SHEET_NAME_NODATE = "NoDate-BothCenters";

  const FOLDER_ID = "YOUR_FOLDER_ID_HERE"; 
  const SUBJECT_LINE = "your email subject here";

  const STATUS_FILTER = [
    "Site Visit Confirmed",
    "Condition Match 1",
    "Condition Match 2",
    "Condition Match 3",
    "Condition Match 4"
  ];

  const NODATE_STATUS = "Scheduled-No Date";

  const CITY_FILTER_NORTH = ["Gandhinagar", "Ahmedabad", "Mumbai", "Navi Mumbai", "Pune", "Thane"];
  const CITY_FILTER_SOUTH = ["Bangalore", "Hyderabad", "Chennai", "Kolkata"];

  const logs = [];
  const start = new Date();
  logs.push("Script started");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shNorth = ss.getSheetByName(SHEET_NAME_NORTH);
  const shSouth = ss.getSheetByName(SHEET_NAME_SOUTH);
  const shNoDate = ss.getSheetByName(SHEET_NAME_NODATE);

  // Validate sheets
  if (!shNorth || !shSouth || !shNoDate) {
    logs.push("Target sheets missing");
    writeLogs_Clean(logs);
    throw new Error("Target sheets not found");
  }

  // Fetch latest ZIP file from Gmail
  const zipFile = getLatestZipFromGmail_Clean(SUBJECT_LINE, FOLDER_ID, logs);
  if (!zipFile) {
    logs.push("ZIP attachment missing");
    writeLogs_Clean(logs);
    throw new Error("No ZIP attachment found");
  }

  // Extract CSV from ZIP
  let blobs;
  try {
    blobs = Utilities.unzip(zipFile.getBlob());
  } catch (e) {
    logs.push("Error unzipping ZIP: " + e.message);
    writeLogs_Clean(logs);
    throw e;
  }

  const csvBlob = blobs.find(b => b.getName().toLowerCase().includes(".csv")) || blobs[0];
  const csvData = Utilities.parseCsv(csvBlob.getDataAsString("UTF-8"));
  if (!csvData.length) {
    logs.push("CSV empty");
    writeLogs_Clean(logs);
    throw new Error("CSV is empty");
  }

  const header = csvData[0];
  const rows = csvData.slice(1);

  // Filter for North and South
  const filteredNorth = rows.filter(r => {
    const city = (r[6] || "").toString().trim().toLowerCase();
    const status = (r[9] || "").toString().trim();
    return (
      STATUS_FILTER.indexOf(status) !== -1 &&
      CITY_FILTER_NORTH.some(c => c.toLowerCase() === city)
    );
  });

  const filteredSouth = rows.filter(r => {
    const city = (r[6] || "").toString().trim().toLowerCase();
    const status = (r[9] || "").toString().trim();
    return (
      STATUS_FILTER.indexOf(status) !== -1 &&
      CITY_FILTER_SOUTH.some(c => c.toLowerCase() === city)
    );
  });

  const filteredNoDate = rows.filter(r => {
    const status = (r[9] || "").toString().trim();
    return status === NODATE_STATUS;
  });

  // Write into sheets
  processSheetBlock(shNorth, header, filteredNorth, logs);
  processSheetBlock(shSouth, header, filteredSouth, logs);
  processSheetBlockNoDate(shNoDate, header, filteredNoDate, logs);

  const end = new Date();
  logs.push("Script completed");
  writeLogs_Clean(logs);
}

function processSheetBlock(sh, header, filtered, logs) {
  const headerRow = 1;
  const lastRow = sh.getLastRow();

  // Clear older rows except header
  if (lastRow > headerRow) {
    sh.getRange(headerRow + 1, 1, lastRow - headerRow, sh.getLastColumn()).clearContent();
  }

  // Add cleaned data
  const finalData = filtered.map(r => {
    const newRow = r.slice(0, 26);
    const mDate = parseCSVDate_Clean(r[12]);
    const aaDate = mDate ? new Date(mDate.getFullYear(), mDate.getMonth(), mDate.getDate()) : "";
    newRow.push(aaDate);
    return newRow;
  });

  if (finalData.length > 0) {
    sh.getRange(2, 1, finalData.length, 27).setValues(finalData);
  }

  // Format dates
  if (finalData.length) {
    const startRow = 2;
    sh.getRange(startRow, 27, finalData.length).setNumberFormat("dd-MMM-yyyy");
  }

  // Mark Y/N flags
  markYEvery30Days_Clean(sh, logs);
}

function processSheetBlockNoDate(sh, header, filtered) {
  const headerRow = 1;
  const lastRow = sh.getLastRow();

  if (lastRow > headerRow) {
    sh.getRange(headerRow + 1, 1, lastRow - headerRow, sh.getLastColumn()).clearContent();
  }

  if (filtered.length > 0) {
    sh.getRange(2, 1, filtered.length, header.length).setValues(filtered);
  }
}

function getLatestZipFromGmail_Clean(subject, folderId, logs) {
  try {
    const threads = GmailApp.search('subject:"' + subject + '"', 0, 5);
    if (!threads.length) return null;

    const lastMsg = threads[0].getMessages().pop();
    const zipAttachment = lastMsg.getAttachments().find(a => a.getName().toLowerCase().endsWith(".zip"));
    if (!zipAttachment) return null;

    const folder = DriveApp.getFolderById(folderId);

    const oldFiles = folder.getFiles();
    while (oldFiles.hasNext()) {
      oldFiles.next().setTrashed(true);
    }

    const saved = folder.createFile(zipAttachment);
    saved.setName("import_" + new Date().getTime() + "_" + zipAttachment.getName());
    return saved;
  } catch (e) {
    logs.push("Gmail error: " + e.message);
    return null;
  }
}

function parseCSVDate_Clean(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value)) return value;

  const s = String(value).trim();
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{2}))?$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], m[4] || 0, m[5] || 0);

  const d = Date.parse(s);
  return isNaN(d) ? null : new Date(d);
}

function markYEvery30Days_Clean(sh, logs) {
  const HEADER_ROW = 1;
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return;

  const COL_BUYER = 3;
  const COL_PROJECT = 5;
  const COL_DATE = 27;
  const COL_OUT = 28;

  const numRows = lastRow - HEADER_ROW;

  const buyers = sh.getRange(HEADER_ROW + 1, COL_BUYER, numRows, 1).getValues();
  const projects = sh.getRange(HEADER_ROW + 1, COL_PROJECT, numRows, 1).getValues();
  const dates = sh.getRange(HEADER_ROW + 1, COL_DATE, numRows, 1).getValues();

  const flags = [];
  const lastYDate = {};

  for (let i = 0; i < numRows; i++) {
    const buyer = String(buyers[i][0] || "").trim();
    const project = String(projects[i][0] || "").trim();

    let date = dates[i][0];
    if (!(date instanceof Date)) date = parseCSVDate_Clean(date);

    if (!buyer || !project || !date) {
      flags.push([""]);
      continue;
    }

    const key = buyer + "|" + project;

    if (!lastYDate[key]) {
      flags.push(["Y"]);
      lastYDate[key] = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    } else {
      const prev = lastYDate[key];
      const diffDays = Math.floor((date - prev) / (1000 * 60 * 60 * 24));

      if (diffDays >= 30) {
        flags.push(["Y"]);
        lastYDate[key] = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      } else {
        flags.push(["N"]);
      }
    }
  }

  sh.getRange(HEADER_ROW + 1, COL_OUT, numRows, 1).setValues(flags);
}

function writeLogs_Clean(logs) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("ScriptLogs");

  if (!sh) {
    sh = ss.insertSheet("ScriptLogs");
    sh.appendRow(["Timestamp", "Message"]);
  }

  const ts = new Date();
  logs.forEach(msg => sh.appendRow([ts, msg]));
}
