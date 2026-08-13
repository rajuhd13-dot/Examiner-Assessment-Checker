// ====================================================================================
// GOOGLE APPS SCRIPT WEB APP - FOR HIGH PERFORMANCE REAL-TIME STUDENT ASSESSMENT DATA
// ====================================================================================
// This script enables instant search and bi-directional real-time database sync.
// Store this code in your Google Sheet's script editor (Extensions -> Apps Script).

var CONFIG = {
  // Replace this ID if you change your spreadsheet
  SPREADSHEET_ID:  '1R_O4llA1K43Y97GAgkK97WMvWbqg-tftz_FXpcUSZPU',
  SHEET_NAME:      'Examiner Information', // Spreadsheet active tab name
  DATA_START_ROW:  2,                       // Header is on Row 1, Data starts on Row 2
  TOTAL_COLS:      96                       // Total columns to pull / write back
};

// 1-Based Index Column Mappings
var COL = {
  NICK_NAME: 2, 
  TPIN: 4, 
  INST: 5, 
  DEPT: 6, 
  HSC_BATCH: 7, 
  RM: 8,
  MOBILE_1: 10, 
  MOBILE_2: 11, 
  MOBILE_BANKING: 12,
  RUNNING_PROGRAM: 16, 
  PREVIOUS_PROGRAM: 17,
  EMAIL: 22, 
  TEAMS_ID: 23,
  HSC_ROLL: 28, 
  HSC_REG: 29, 
  HSC_BOARD: 30, 
  HSC_GPA: 31,
  SUBJECT_1: 34, 
  SUBJECT_2: 35, 
  SUBJECT_3: 36, 
  SUBJECT_4: 37, 
  SUBJECT_5: 38,
  VERSION_INTERESTED: 39,
  FULL_NAME: 43, 
  RELIGION: 45, 
  GENDER: 46, 
  DATE_OF_BIRTH: 47,
  FATHERS_NAME: 52, 
  MOTHERS_NAME: 56, 
  HOME_DISTRICT: 61,
  ENGLISH_PCT: 62, 
  ENGLISH_SET: 63, 
  ENGLISH_DATE: 64,
  BANGLA_PCT: 65, 
  BANGLA_SET: 66, 
  BANGLA_DATE: 67,
  PHYSICS_PCT: 68, 
  PHYSICS_SET: 69, 
  PHYSICS_DATE: 70,
  CHEMISTRY_PCT: 71, 
  CHEMISTRY_SET: 72, 
  CHEMISTRY_DATE: 73,
  MATH_PCT: 74, 
  MATH_SET: 75, 
  MATH_DATE: 76,
  BIOLOGY_PCT: 77, 
  BIOLOGY_SET: 78, 
  BIOLOGY_DATE: 79,
  ICT_PCT: 80, 
  ICT_SET: 81, 
  ICT_DATE: 82,
  TRAINING_REPORT: 83, 
  TRAINING_DATE: 84,
  ID_CHECKED: 86, 
  FORM_FILL_DATE: 88, 
  PHYSICAL_CAMPUS_PREF: 89,
  SELECTED_SUBJECT: 92,
  REMARK_COMMENT: 93,
  REMARK_COUNT: 94, 
  REMARK_TEXT: 94, 
  REMARK_BY: 95, 
  REMARK_DATE: 96
};

// Automatic Header Column Index Detection
function detectColumns_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol <= 0) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  var headerList = [];
  for (var i = 0; i < headers.length; i++) {
    headerList.push(String(headers[i] || '').trim().toLowerCase());
  }
  
  for (var c = 0; c < headerList.length; c++) {
    var h = headerList[c];
    var colNum = c + 1;
    
    if (h === 'tpin' || h === 't-pin' || h.indexOf('tpin') > -1) {
      COL.TPIN = colNum;
    } else if (h.indexOf('roll') > -1 && h.indexOf('hsc') > -1) {
      COL.HSC_ROLL = colNum;
    } else if (h.indexOf('reg') > -1 && h.indexOf('hsc') > -1) {
      COL.HSC_REG = colNum;
    } else if (h === 'roll' || h === 'roll no' || h === 'roll_no' || h === 'roll number') {
      if (!COL.HSC_ROLL || COL.HSC_ROLL === 28) {
        COL.HSC_ROLL = colNum;
      }
    } else if (h === 'reg' || h === 'reg no' || h === 'reg_no' || h === 'registration' || h === 'registration no' || h === 'registration number') {
      if (!COL.HSC_REG || COL.HSC_REG === 29) {
        COL.HSC_REG = colNum;
      }
    } else if (h.indexOf('full name') > -1 || h === 'name') {
      COL.FULL_NAME = colNum;
    } else if (h.indexOf('nick name') > -1) {
      COL.NICK_NAME = colNum;
    } else if (h.indexOf('rm') > -1) {
      COL.RM = colNum;
    } else if (h.indexOf('mobile') > -1 && (h.indexOf('1') > -1 || h.indexOf('one') > -1)) {
      COL.MOBILE_1 = colNum;
    } else if (h.indexOf('mobile') > -1 && (h.indexOf('2') > -1 || h.indexOf('two') > -1)) {
      COL.MOBILE_2 = colNum;
    }
  }
}

// Locate Sheet Safely
function getSheet() {
  var ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {}
  if (!ss && CONFIG.SPREADSHEET_ID) {
    try {
      ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    } catch (err) {}
  }
  if (!ss) {
    throw new Error('Spreadsheet not found. Please bind this script to your Google Sheet or configure CONFIG.SPREADSHEET_ID in Code.gs.');
  }
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.getSheets()[0];
  }
  if (!sheet) {
    throw new Error('No sheet found inside your spreadsheet.');
  }
  try {
    detectColumns_(sheet);
  } catch (e) {}
  return sheet;
}

// ------------------------------------------------------------------------------------
// GET REQUESTS ROUTER (GET ENTRANCES)
// ------------------------------------------------------------------------------------
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'debug') {
      var sheet = getSheet();
      var ss = sheet.getParent();
      var sheetNames = ss.getSheets().map(function(s) { return s.getName(); });
      return ContentService.createTextOutput(JSON.stringify({
        ok: true,
        spreadsheetId: ss.getId(),
        spreadsheetName: ss.getName(),
        sheetName: sheet.getName(),
        lastRow: sheet.getLastRow(),
        sheetNames: sheetNames,
        configSheetName: CONFIG.SHEET_NAME,
        configSpreadsheetId: CONFIG.SPREADSHEET_ID
      })).setMimeType(ContentService.MimeType.JSON);
    }
    if (e && e.parameter && e.parameter.action === 'sync') {
      return ContentService.createTextOutput(JSON.stringify(syncAllData()))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (e && e.parameter && e.parameter.q) {
      var result = searchExaminer(e.parameter.q);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (e && e.parameter && e.parameter.action === 'search') {
      var rollQuery = e.parameter.roll;
      var regQuery = e.parameter.reg;
      var res = searchByHscRollReg(rollQuery, regQuery);
      return ContentService.createTextOutput(JSON.stringify(res))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: 'No query provided.' }))
        .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: 'Script Error: ' + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ------------------------------------------------------------------------------------
// POST REQUESTS ROUTER (WRITE ENTRANCES)
// ------------------------------------------------------------------------------------
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var params = JSON.parse(e.postData.contents);
    if (params.action === 'update') {
      return ContentService.createTextOutput(JSON.stringify(updateRow(params.tpin, params.updates)))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: 'Invalid action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, message: 'Update Error: ' + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------------------------------
// CORE DB WRITE FUNCTION
// ------------------------------------------------------------------------------------
function updateRow(tpin, updates) {
  if (!tpin) return { ok: false, message: 'TPIN required' };
  
  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  if (lastRow < CONFIG.DATA_START_ROW) return { ok: false, message: 'No data found in sheet' };
  
  var tpinData = sheet.getRange(CONFIG.DATA_START_ROW, COL.TPIN, lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues();
  
  var rowIdx = -1;
  for (var i = 0; i < tpinData.length; i++) {
    if (String(tpinData[i][0]).trim() === String(tpin).trim()) {
      rowIdx = i + CONFIG.DATA_START_ROW;
      break;
    }
  }
  
  if (rowIdx === -1) return { ok: false, message: 'Examiner with TPIN ' + tpin + ' not found' };
  
  var rowRange = sheet.getRange(rowIdx, 1, 1, lastCol);
  var rowValues = rowRange.getValues();
  var row = rowValues[0];
  
  var changed = false;
  for (var key in updates) {
    if (COL[key]) {
      var colIdx = COL[key] - 1;
      if (colIdx < lastCol) {
        row[colIdx] = updates[key];
        changed = true;
      }
    }
  }
  
  if (changed) {
    rowRange.setValues([row]);
  }
  
  return { ok: true, message: 'Updated successfully' };
}

// ------------------------------------------------------------------------------------
// SYNC ALL DATA EXPORTER
// ------------------------------------------------------------------------------------
function syncAllData() {
  var sheet = getSheet();
  var ss = sheet.getParent();
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return { ok: true, data: [] };
  
  var data = sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, CONFIG.TOTAL_COLS).getValues();
  
  var filtered = data.filter(function(row) {
    var rollVal = row[COL.HSC_ROLL - 1];
    var nameVal = row[COL.FULL_NAME - 1];
    var tpinVal = row[COL.TPIN - 1];
    return (rollVal && String(rollVal).trim() !== '') || 
           (nameVal && String(nameVal).trim() !== '') || 
           (tpinVal && String(tpinVal).trim() !== '');
  });
  
  return { 
    ok: true, 
    data: filtered,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    sheetName: sheet.getName(),
    lastRow: lastRow,
    totalRows: filtered.length
  };
}

// ------------------------------------------------------------------------------------
// CLEANSE ROLL / REG TO STANDARDIZED DIGITS
// ------------------------------------------------------------------------------------
function normalizeRollReg_(v) {
  if (v === undefined || v === null) return "";
  var str = String(v).trim();
  
  var banglaDigits = ["০", "১", "২", "৩", "৪", "৫", "৬", "৭", "৮", "৯"];
  for (var i = 0; i < banglaDigits.length; i++) {
    var re = new RegExp(banglaDigits[i], "g");
    str = str.replace(re, String(i));
  }
  
  str = str.replace(/,/g, "");
  
  if (/[eE][+-]?\d+/.test(str)) {
    var num = Number(str);
    if (!isNaN(num)) {
      str = num.toFixed(0);
    }
  }
  
  if (str.indexOf('.') > -1) {
    var pNum = parseFloat(str);
    if (!isNaN(pNum)) {
      str = String(Math.round(pNum));
    }
  }
  
  return str.replace(/\D/g, "");
}

// ------------------------------------------------------------------------------------
// TARGETED SEARCH: ROLL & REGISTRATION NO
// ------------------------------------------------------------------------------------
function searchByHscRollReg(rollQuery, regQuery) {
  try {
    rollQuery = normalizeRollReg_(rollQuery);
    regQuery = normalizeRollReg_(regQuery);

    if (!rollQuery || !regQuery) {
      return { ok: false, message: 'HSC Roll No এবং Registration No দুটোই দিন।' };
    }

    var sheet = getSheet();
    var lastRow = sheet.getLastRow();
    
    if (lastRow < CONFIG.DATA_START_ROW) {
      return { ok: false, message: 'শীটে কোনো ডাটা পাওয়া যায়নি।' };
    }

    // single-batch API read for extreme efficiency
    var totalRows = lastRow - CONFIG.DATA_START_ROW + 1;
    var allRows = sheet.getRange(CONFIG.DATA_START_ROW, 1, totalRows, CONFIG.TOTAL_COLS).getDisplayValues();

    var foundIdx = -1;
    var rollColIdx = COL.HSC_ROLL - 1;
    var regColIdx = COL.HSC_REG - 1;

    // Search pass 1: Exact matches
    for (var i = 0; i < allRows.length; i++) {
      var rVal = normalizeRollReg_(allRows[i][rollColIdx]);
      var rgVal = normalizeRollReg_(allRows[i][regColIdx]);
      if (rVal === rollQuery && rgVal === regQuery) {
        foundIdx = i;
        break;
      }
    }

    // Search pass 2: Swapped input fallback
    if (foundIdx === -1) {
      for (var i = 0; i < allRows.length; i++) {
        var rVal = normalizeRollReg_(allRows[i][rollColIdx]);
        var rgVal = normalizeRollReg_(allRows[i][regColIdx]);
        if (rVal === regQuery && rgVal === rollQuery) {
          foundIdx = i;
          break;
        }
      }
    }

    // Search pass 3: Loose substring search fallback across the entire row values
    if (foundIdx === -1) {
      for (var i = 0; i < allRows.length; i++) {
        var row = allRows[i];
        var rowStr = "";
        for (var j = 0; j < row.length; j++) {
          rowStr += normalizeRollReg_(row[j]) + "|";
        }
        if (rowStr.indexOf(rollQuery) > -1 && rowStr.indexOf(regQuery) > -1) {
          foundIdx = i;
          break;
        }
      }
    }

    var ss = sheet.getParent();
    if (foundIdx === -1) {
      return { 
        ok: false, 
        message: 'Roll ও Registration নম্বর মিলছে না। সঠিক নম্বর দিন।',
        metadata: {
          spreadsheetId: ss.getId(),
          spreadsheetName: ss.getName(),
          sheetName: sheet.getName(),
          lastRow: lastRow
        }
      };
    }

    var rowData = allRows[foundIdx];
    return { 
      ok: true, 
      data: mapRow_(rowData),
      metadata: {
        spreadsheetId: ss.getId(),
        spreadsheetName: ss.getName(),
        sheetName: sheet.getName(),
        lastRow: lastRow
      }
    };

  } catch (err) {
    return { ok: false, message: 'Script Error: ' + err.message };
  }
}

function norm_(v) {
  v = String(v || '').trim();
  if (!v) return '';
  var d = v.replace(/\D/g, '');
  if (d) {
    if (d.length >= 12 && d.slice(0,3) === '880') return d;
    if (d[0] === '0' && d.length === 11) return '88' + d;
    if (d[0] === '1' && d.length === 10) return '880' + d;
    return d;
  }
  return v.toUpperCase();
}

function searchExaminer(query) {
  query = String(query || '').trim();
  if (!query) return { ok: false, message: 'Search value is empty.' };

  var key = norm_(query);
  if (!key) return { ok: false, message: 'Invalid search key.' };

  var sheet = getSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    return { ok: false, message: 'No examiner found.' };
  }

  var totalRows = lastRow - CONFIG.DATA_START_ROW + 1;
  var allRows = sheet.getRange(CONFIG.DATA_START_ROW, 1, totalRows, CONFIG.TOTAL_COLS).getDisplayValues();

  var foundIdx = -1;
  var tpinColIdx = COL.TPIN - 1;
  var mob1ColIdx = COL.MOBILE_1 - 1;
  var mob2ColIdx = COL.MOBILE_2 - 1;

  for (var i = 0; i < allRows.length; i++) {
    var row = allRows[i];
    var tpinVal = norm_(row[tpinColIdx]);
    var mob1Val = norm_(row[mob1ColIdx]);
    var mob2Val = norm_(row[mob2ColIdx]);

    if (tpinVal === key || mob1Val === key || mob2Val === key) {
      foundIdx = i;
      break;
    }
  }

  if (foundIdx === -1) {
    for (var i = 0; i < allRows.length; i++) {
      var row = allRows[i];
      var tpinVal = String(row[tpinColIdx]).trim();
      var mob1Val = String(row[mob1ColIdx]).trim();
      var mob2Val = String(row[mob2ColIdx]).trim();

      if (tpinVal.indexOf(query) > -1 || mob1Val.indexOf(query) > -1 || mob2Val.indexOf(query) > -1) {
        foundIdx = i;
        break;
      }
    }
  }

  if (foundIdx === -1) {
    return { ok: false, message: 'No examiner found.' };
  }

  var rowData = allRows[foundIdx];
  var mappedData = mapRow_(rowData);
  return { ok: true, data: mappedData };
}

function mapRow_(row) {
  var g = function(c) { return row[c - 1] || ''; };
  var rm = String(g(COL.RM)).trim();
  var rmNum = extractNum_(rm);
  var remarkRaw = String(g(COL.REMARK_COMMENT)).trim();
  var parsedRemark = parseRemarkCell_(remarkRaw, rmNum);

  return {
    quick: {
      tpin: g(COL.TPIN), rm: rm, nickName: g(COL.NICK_NAME),
      fullName: g(COL.FULL_NAME), mobile1: g(COL.MOBILE_1), mobile2: g(COL.MOBILE_2),
      nagadNumber: g(COL.MOBILE_BANKING), institute: g(COL.INST), department: g(COL.DEPT),
      hscGpa: g(COL.HSC_GPA), hscBatch: fmtBatch_(g(COL.HSC_BATCH)),
      trainingReport: g(COL.TRAINING_REPORT), trainingDate: g(COL.TRAINING_DATE),
      physicalCampus: g(COL.PHYSICAL_CAMPUS_PREF)
    },
    assessments: [
      mkAs_('English',   g(COL.ENGLISH_PCT),   g(COL.ENGLISH_SET),   g(COL.ENGLISH_DATE),   60),
      mkAs_('Bangla',    g(COL.BANGLA_PCT),     g(COL.BANGLA_SET),    g(COL.BANGLA_DATE),    50),
      mkAs_('Physics',   g(COL.PHYSICS_PCT),    g(COL.PHYSICS_SET),   g(COL.PHYSICS_DATE),   50),
      mkAs_('Chemistry', g(COL.CHEMISTRY_PCT),  g(COL.CHEMISTRY_SET), g(COL.CHEMISTRY_DATE), 50),
      mkAs_('Math',      g(COL.MATH_PCT),       g(COL.MATH_SET),      g(COL.MATH_DATE),      50),
      mkAs_('Biology',   g(COL.BIOLOGY_PCT),    g(COL.BIOLOGY_SET),   g(COL.BIOLOGY_DATE),   50),
      mkAs_('ICT',       g(COL.ICT_PCT),        g(COL.ICT_SET),       g(COL.ICT_DATE),       50)
    ],
    remark: parsedRemark,
    personal: {
      fathersName: g(COL.FATHERS_NAME), mothersName: g(COL.MOTHERS_NAME),
      religion: g(COL.RELIGION), gender: g(COL.GENDER), dateOfBirth: g(COL.DATE_OF_BIRTH),
      hscRoll: g(COL.HSC_ROLL), hscReg: g(COL.HSC_REG), teamsId: g(COL.TEAMS_ID),
      hscBoard: g(COL.HSC_BOARD), email: g(COL.EMAIL), regDate: g(COL.FORM_FILL_DATE),
      homeDistrict: g(COL.HOME_DISTRICT),
      subjectsChoice: [g(COL.SUBJECT_1),g(COL.SUBJECT_2),g(COL.SUBJECT_3),g(COL.SUBJECT_4),g(COL.SUBJECT_5)].filter(Boolean).join(', '),
      selectedSub: g(COL.SELECTED_SUBJECT), versionInterested: g(COL.VERSION_INTERESTED),
      idChecked: g(COL.ID_CHECKED), runningProgram: g(COL.RUNNING_PROGRAM),
      previousProgram: g(COL.PREVIOUS_PROGRAM)
    }
  };
}

function parseRemarkCell_(raw, rmNum) {
  var show = (rmNum >= 4) || (raw.length > 0 && rmNum > 0);
  if (!show || !raw) return { count: rmNum, show: false, body: '', byLine: '', dateLine: '' };

  var lines = raw.replace(/\r/g, '').split('\n');
  var bodyLines = [], byLine = '', dateLine = '';

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    if (line.charAt(0) === '#') byLine = line;
    else if (/^date\s*:/i.test(line)) dateLine = line;
    else bodyLines.push(line);
  }

  var body = bodyLines.join('\n').trim();
  if (!body) body = 'সমস্যাঃ\n** খাতা দেখার নিয়ম না মেনে খাতা দেখা।\n** প্রিন্টিং কমেন্ট করা।\n** কনসেপ্ট দুর্বল।\n** একাধিকবার সুযোগ দেয়া সত্ত্বেও শুধরাতে পারেননি।';

  return { count: rmNum, show: true, body: body, byLine: byLine, dateLine: dateLine };
}

function mkAs_(name, pct, set, date, pass) {
  var p = String(pct||'').trim(), s = String(set||'').trim(), d = String(date||'').trim();
  var sc = parseScore_(p);
  var st = (p||s||d) ? ((sc !== null && sc >= pass) ? 'Allow' : 'Not Allow') : 'No Exam';
  return { subject: name + ' (%)', percent: p, set: s, date: d, status: st };
}

function parseScore_(v) {
  v = String(v||'').trim();
  if (!v) return null;
  var fm = v.match(/(\d+(?:\.\d+)?)\s*\/\s*\d+/);
  if (fm) return Number(fm[1]);
  var m = v.match(/-?\d+(?:\.\d+)?/);
  return m ? Number(m[0]) : null;
}

// Extract number from RM count string
function extractNum_(v) {
  var m = String(v||'').match(/\d+/);
  return m ? Number(m[0]) : 0;
}

// Format year prefix for batch (e.g. 23 -> 2023)
function fmtBatch_(v) {
  v = String(v||'').trim();
  return /^\d{2}$/.test(v) ? '20' + v : v;
}

// ------------------------------------------------------------------------------------
// AUTOMATIC TRIGGER (ON EDIT) TO CLEAR WEB SERVER CACHE INSTANTLY
// ------------------------------------------------------------------------------------
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() === CONFIG.SHEET_NAME) {
     var tpin = sheet.getRange(e.range.getRow(), COL.TPIN).getValue();
     
     // Dynamic Endpoint Ingestion URL
     var endpointUrl = "https://ais-dev-ddmcf52xgr6udwnqohb35b-192410877328.asia-southeast1.run.app/api/refresh?tpin=" + tpin;
     
     UrlFetchApp.fetch(endpointUrl, {
       'method': 'get',
       'muteHttpExceptions': true
     });
  }
}
