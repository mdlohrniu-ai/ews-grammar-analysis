// ================================================================
//  EWS Blog Study Guide — Google Apps Script
//
//  SETUP INSTRUCTIONS:
//  1. Create a blank Google Sheet
//  2. Open Extensions > Apps Script
//  3. Paste this entire file, replacing any existing code
//  4. Click Save
//  5. Run setupSheet() once from the Run menu
//     — this builds the roster tabs and Drive folders
//  6. Deploy > New deployment > Web App
//       Execute as: Me
//       Who has access: Anyone
//  7. Copy the Web App URL into the HTML page config
//
//  WEEKLY WORKFLOW:
//  1. Export Study Guide from Excel (purple button)
//  2. Upload the JSON file to the matching Drive folder
//     (URLs shown in the Settings tab after setup)
//  3. Students visit the HTML page and log in with their ID
// ================================================================


// ================================================================
//  CONFIG — reads folder IDs from Settings tab
// ================================================================

function getConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName('Settings');
  if (!settings) throw new Error('Settings tab not found. Run setupSheet() first.');
  return {
    ewsiFolderId:  settings.getRange('B2').getValue(),
    ewsiiiFolderId: settings.getRange('B3').getValue()
  };
}


// ================================================================
//  SETUP — run once on a blank spreadsheet
// ================================================================

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Settings tab
  var settingsWs = ss.getSheetByName('Settings');
  if (!settingsWs) settingsWs = ss.insertSheet('Settings');
  else { settingsWs.clearContents(); settingsWs.clearFormats(); }
  buildSettingsTab(settingsWs);

  // EWSI roster tab
  var ewsiWs = ss.getSheetByName('EWSI');
  if (!ewsiWs) ewsiWs = ss.insertSheet('EWSI');
  else { ewsiWs.clearContents(); ewsiWs.clearFormats(); }
  buildRosterTab(ewsiWs, 'EWSI');

  // EWSIII roster tab
  var ewsiiiWs = ss.getSheetByName('EWSIII');
  if (!ewsiiiWs) ewsiiiWs = ss.insertSheet('EWSIII');
  else { ewsiiiWs.clearContents(); ewsiiiWs.clearFormats(); }
  buildRosterTab(ewsiiiWs, 'EWSIII');

  // Create Drive folders
  var root = DriveApp.getRootFolder();
  var ewsiFolder   = getOrCreateFolder(root, 'EWS_StudyGuide_EWSI');
  var ewsiiiFolder = getOrCreateFolder(root, 'EWS_StudyGuide_EWSIII');

  // Store folder IDs and URLs in Settings
  settingsWs.getRange('B2').setValue(ewsiFolder.getId());
  settingsWs.getRange('B3').setValue(ewsiiiFolder.getId());
  settingsWs.getRange('B4').setValue(ewsiFolder.getUrl());
  settingsWs.getRange('B5').setValue(ewsiiiFolder.getUrl());

  // Order tabs: Settings first
  ss.setActiveSheet(settingsWs);
  ss.moveActiveSheet(1);

  SpreadsheetApp.getUi().alert(
    'Setup complete!\n\n' +
    'Two Drive folders created:\n' +
    '  EWSI:   ' + ewsiFolder.getUrl() + '\n' +
    '  EWSIII: ' + ewsiiiFolder.getUrl() + '\n\n' +
    'Upload your exported JSON files to those folders.\n\n' +
    'Next step:\n' +
    '  Deploy > New deployment > Web App\n' +
    '  Execute as: Me\n' +
    '  Who has access: Anyone\n\n' +
    'Then paste the Web App URL into the HTML page config.'
  );
}

function getOrCreateFolder(parent, name) {
  var iter = parent.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return parent.createFolder(name);
}

function buildSettingsTab(ws) {
  var darkGreen  = '#2d6a2d';
  var lightGreen = '#e8f5e9';

  ws.getRange('A1:B1').merge()
    .setValue('EWS Study Guide — Settings')
    .setBackground(darkGreen).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(12);
  ws.setRowHeight(1, 28);

  var rows = [
    ['EWSI Drive Folder ID',   ''],
    ['EWSIII Drive Folder ID', ''],
    ['EWSI Folder URL',        ''],
    ['EWSIII Folder URL',      '']
  ];
  for (var i = 0; i < rows.length; i++) {
    var r = i + 2;
    ws.getRange('A' + r).setValue(rows[i][0])
      .setBackground(lightGreen).setFontWeight('bold').setFontSize(10);
    ws.getRange('B' + r).setValue(rows[i][1]).setFontSize(10);
    ws.setRowHeight(r, 20);
  }
  ws.getRange('A7')
    .setValue('Do not edit this tab manually.')
    .setFontStyle('italic').setFontColor('#888888').setFontSize(9);

  ws.setColumnWidth(1, 200);
  ws.setColumnWidth(2, 500);
}

function buildRosterTab(ws, className) {
  var darkGreen  = '#2d6a2d';
  var midGreen   = '#4caf50';
  var lightGreen = '#e8f5e9';
  var altGreen   = '#f1f8f1';

  // Title
  ws.getRange('A1:C1').merge()
    .setValue(className + ' — Student Roster')
    .setBackground(darkGreen).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(12)
    .setHorizontalAlignment('center');
  ws.setRowHeight(1, 28);

  // Headers
  var headers = ['Student ID', 'Name (Eng.)', 'Textbook'];
  for (var h = 0; h < headers.length; h++) {
    ws.getRange(2, h + 1)
      .setValue(headers[h])
      .setBackground(midGreen).setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center');
  }
  ws.setRowHeight(2, 20);

  // Textbook dropdown validation
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Basic', 'Intermediate', 'Advanced'], true)
    .setAllowInvalid(false).build();
  ws.getRange('C3:C52').setDataValidation(rule);

  // Alternating row colours
  for (var r = 3; r <= 52; r++) {
    var bg = (r % 2 === 1) ? lightGreen : altGreen;
    ws.getRange(r, 1, 1, 3).setBackground(bg).setFontSize(10);
    ws.getRange(r, 1).setHorizontalAlignment('center');
    ws.getRange(r, 2).setHorizontalAlignment('left');
    ws.getRange(r, 3).setHorizontalAlignment('center');
    ws.setRowHeight(r, 20);
  }

  ws.setColumnWidth(1, 120);
  ws.setColumnWidth(2, 220);
  ws.setColumnWidth(3, 120);
  ws.setFrozenRows(2);

  // Notes
  ws.getRange('A54')
    .setValue('Name (Eng.) must match the student sheet tab name in the Excel workbook exactly (e.g. Shione SAZA).')
    .setFontStyle('italic').setFontColor('#888888').setFontSize(9);
  ws.getRange('A55')
    .setValue('Student ID must match what is in the Excel Roster sheet.')
    .setFontStyle('italic').setFontColor('#888888').setFontSize(9);
}


// ================================================================
//  WEB APP — handles all requests from the HTML page
// ================================================================

function doGet(e) {
  var action = e.parameter.action || '';
  var result;

  try {
    if (action === 'getStudent') {
      result = getStudent(e.parameter.id);
    } else if (action === 'getWeeks') {
      result = getWeeks(e.parameter.id);
    } else if (action === 'getWeek') {
      result = getWeekData(e.parameter.id, parseInt(e.parameter.week));
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON)
    .setCharset('UTF-8');
}


// ── getStudent ─────────────────────────────────────────────────
// Searches both class tabs for a student ID.
// Returns: { id, nameEng, textbook, class } or { error }

function getStudent(id) {
  if (!id) return { error: 'No ID provided.' };
  id = id.toString().trim();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classes = ['EWSI', 'EWSIII'];

  for (var c = 0; c < classes.length; c++) {
    var ws = ss.getSheetByName(classes[c]);
    if (!ws) continue;
    var data = ws.getDataRange().getValues();
    // Row 0 = title, row 1 = headers, rows 2+ = data
    for (var r = 2; r < data.length; r++) {
      var rowId = data[r][0].toString().trim();
      if (rowId === id) {
        var nameEng = data[r][1].toString().trim();
        if (nameEng === '') continue; // skip blank rows
        return {
          id:       rowId,
          nameEng:  nameEng,
          textbook: data[r][2].toString().trim(),
          class:    classes[c]
        };
      }
    }
  }
  return { error: 'Student ID not found. Please check your ID and try again.' };
}


// ── getWeeks ───────────────────────────────────────────────────
// Returns sorted list of weeks that have data in the JSON.
// Returns: { nameEng, textbook, class, weeks: [1,2,3,...] } or { error }

function getWeeks(id) {
  if (!id) return { error: 'No ID provided.' };

  var student = getStudent(id);
  if (student.error) return student;

  var json = loadClassJson(student.class);
  if (json.error) return json;

  var stuData = findStudentInJson(json, student.nameEng);
  if (!stuData) return {
    nameEng:  student.nameEng,
    textbook: student.textbook,
    class:    student.class,
    weeks:    []
  };

  var weeks = stuData.weeks.map(function(w) { return w.week; });
  weeks.sort(function(a, b) { return a - b; });

  return {
    nameEng:  student.nameEng,
    textbook: student.textbook,
    class:    student.class,
    weeks:    weeks
  };
}


// ── getWeekData ────────────────────────────────────────────────
// Returns annotated blog + errors for one student/week.
// Returns: { week, date, annotatedText, errors:[...] } or { error }

function getWeekData(id, weekNum) {
  if (!id) return { error: 'No ID provided.' };
  if (!weekNum) return { error: 'No week provided.' };

  var student = getStudent(id);
  if (student.error) return student;

  var json = loadClassJson(student.class);
  if (json.error) return json;

  var stuData = findStudentInJson(json, student.nameEng);
  if (!stuData) return { error: 'No study guide data found for this student yet.' };

  for (var w = 0; w < stuData.weeks.length; w++) {
    if (stuData.weeks[w].week === weekNum) {
      return stuData.weeks[w];
    }
  }
  return { error: 'No data found for Week ' + weekNum + '.' };
}


// ── loadClassJson ──────────────────────────────────────────────
// Reads the most recently modified JSON file from the class Drive folder.

function loadClassJson(className) {
  var config = getConfig();
  var folderId = (className === 'EWSI') ? config.ewsiFolderId : config.ewsiiiFolderId;

  if (!folderId) {
    return { error: 'Drive folder not configured. Run setupSheet() again.' };
  }

  try {
    var folder = DriveApp.getFolderById(folderId);

    // Search for .json files (uploaded as plain text)
    var latestFile = null;
    var latestDate = new Date(0);

    // Try plain text type first (how Drive stores uploaded .json)
    var fileTypes = [MimeType.PLAIN_TEXT, 'application/json'];
    for (var t = 0; t < fileTypes.length; t++) {
      try {
        var files = folder.getFilesByType(fileTypes[t]);
        while (files.hasNext()) {
          var f = files.next();
          if (f.getName().toLowerCase().indexOf('.json') > -1) {
            var modified = f.getLastUpdated();
            if (modified > latestDate) {
              latestDate = modified;
              latestFile = f;
            }
          }
        }
      } catch(e) { /* mime type may not exist */ }
    }

    // Fallback: search all files in folder
    if (!latestFile) {
      var allFiles = folder.getFiles();
      while (allFiles.hasNext()) {
        var f2 = allFiles.next();
        if (f2.getName().toLowerCase().indexOf('.json') > -1) {
          var mod2 = f2.getLastUpdated();
          if (mod2 > latestDate) {
            latestDate = mod2;
            latestFile = f2;
          }
        }
      }
    }

    if (!latestFile) {
      return { error: 'No JSON file found in ' + className + ' folder. Upload the Study Guide JSON first.' };
    }

    var content = latestFile.getBlob().getDataAsString('UTF-8');
    // Strip UTF-8 BOM if present (added by Windows/Excel ADODB.Stream)
    if (content.charCodeAt(0) === 0xFEFF) content = content.slice(1);
    content = content.trim();
    return JSON.parse(content);

  } catch (err) {
    return { error: 'Could not read Drive folder: ' + err.message };
  }
}


// ── findStudentInJson ──────────────────────────────────────────
// Finds a student by English name — word-order independent.
// "Gaito ABE" matches "ABE GAITO" but not "Fuchi ABE".

function findStudentInJson(json, nameEng) {
  if (!json.students) return null;
  var wordsA = nameEng.toLowerCase().split(' ')
    .filter(function(w) { return w !== ''; });

  for (var s = 0; s < json.students.length; s++) {
    var stu = json.students[s];
    var wordsB = (stu.nameEng || '').toLowerCase().split(' ')
      .filter(function(w) { return w !== ''; });
    if (wordsA.length !== wordsB.length) continue;
    var allMatch = wordsA.every(function(w) { return wordsB.indexOf(w) > -1; });
    if (allMatch) return stu;
  }
  return null;
}
