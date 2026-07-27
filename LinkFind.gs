// ============================================================
// CONFIG
// ============================================================
var CONFIG = {
  INPUT_SHEET_NAME: 'FileID一覧',
  RESULT_SHEET_NAME: '結果',
  ERROR_SHEET_NAME: 'エラー',
  OUTPUT_FOLDER_ID: '',
  TIME_LIMIT_SEC: 240,
  TRIGGER_INTERVAL_MIN: 5,
  ROWS_PER_BATCH: 1000,
  HEADER_COLOR: '#4A86C8',
  HEADER_FONT_COLOR: '#FFFFFF'
};

// ============================================================
// PROP_KEYS
// ============================================================
var PROP_KEYS = {
  RESUME_FILE_INDEX: 'LINKFIND_RESUME_FILE_INDEX',
  RESUME_SHEET_INDEX: 'LINKFIND_RESUME_SHEET_INDEX',
  RESUME_ROW_INDEX: 'LINKFIND_RESUME_ROW_INDEX',
  TOTAL_COUNT: 'LINKFIND_TOTAL_COUNT',
  COMPLETED: 'LINKFIND_COMPLETED',
  DISCOVERED_IDS: 'LINKFIND_DISCOVERED_IDS'
};

// ============================================================
// Menu
// ============================================================

// onOpen - add custom menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('LinkFind')
    .addItem('手動実行', 'runLinkFind')
    .addSeparator()
    .addItem('自動調査を開始', 'startAutoScan')
    .addItem('自動調査を停止', 'stopAutoScan')
    .addSeparator()
    .addItem('再開位置を設定', 'setResumeIndex')
    .addSeparator()
    .addItem('キャッシュクリア', 'clearAllProgress')
    .addToUi();
}

// setResumeIndex - set resume file index manually
function setResumeIndex() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Resume Index',
    'Enter the file index to resume from (0-based):',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() === ui.Button.OK) {
    var index = parseInt(response.getResponseText(), 10);
    if (isNaN(index) || index < 0) {
      ui.alert('Invalid number.');
      return;
    }
    var props = PropertiesService.getScriptProperties();
    props.setProperty(PROP_KEYS.RESUME_FILE_INDEX, String(index));
    props.deleteProperty(PROP_KEYS.RESUME_SHEET_INDEX);
    props.deleteProperty(PROP_KEYS.RESUME_ROW_INDEX);
    ui.alert('Resume file index set to ' + index + '.');
  }
}

// ============================================================
// Manual execution
// ============================================================

// runLinkFind - manual scan with lock
function runLinkFind() {
  var lock = LockService.getScriptLock();
  var hasLock = lock.tryLock(10000);
  if (!hasLock) {
    SpreadsheetApp.getUi().alert('別の処理が実行中です。しばらく待ってから再実行してください。');
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var fileIds = getFileIds_(ss, true);

    if (fileIds.length === 0) {
      SpreadsheetApp.getUi().alert('チェック対象のFileIDが見つかりませんでした。');
      return;
    }

    var resultSheet = prepareResultSheet_(ss, true);
    var errorSheet = prepareErrorSheet_(ss, true);

    var allResults = [];
    var allErrors = [];

    for (var i = 0; i < fileIds.length; i++) {
      var fileId = fileIds[i].trim();
      if (fileId === '') continue;

      Logger.log('Processing: ' + (i + 1) + '/' + fileIds.length + ' - ' + fileId);
      writeStatus_(resultSheet, i, fileIds.length, 'Processing...');

      var scanResult = scanFileForImportRange_(fileId, ss, null, null, null);
      if (scanResult.error) {
        allErrors.push(scanResult.error);
      }
      if (scanResult.results.length > 0) {
        allResults = allResults.concat(scanResult.results);
      }

      addDiscoveredIds_(ss, scanResult.newIds);
    }

    fileIds = getFileIds_(ss, false);

    if (allResults.length > 0) {
      writeResultRows_(resultSheet, allResults, 2);
    }
    if (allErrors.length > 0) {
      writeErrorRows_(errorSheet, allErrors);
    }
    writeStatus_(resultSheet, fileIds.length, fileIds.length, 'Completed');

    outputCsvFiles_(allResults);

    SpreadsheetApp.getUi().alert(
      '調査完了' + '\n' +
      'IMPORTRANGE検出数: ' + allResults.length + '件' + '\n' +
      '結果は' + CONFIG.RESULT_SHEET_NAME + 'シートをご確認ください。'
    );
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// Auto scan
// ============================================================

// startAutoScan - start auto scan with trigger
function startAutoScan() {
  var ui = SpreadsheetApp.getUi();

  var existingTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existingTriggers.length; i++) {
    if (existingTriggers[i].getHandlerFunction() === 'runAutoScan') {
      ui.alert('既に自動調査が実行中です。' + '\n' + '停止してから再度開始してください。');
      return;
    }
  }

  clearAllProgressProps_();

  ScriptApp.newTrigger('runAutoScan')
    .timeBased()
    .everyMinutes(CONFIG.TRIGGER_INTERVAL_MIN)
    .create();

  ui.alert(
    '自動調査を開始します。' + '\n' +
    CONFIG.TRIGGER_INTERVAL_MIN + '分毎にトリガーが実行されます。' + '\n' +
    '初回実行を開始します。'
  );

  runAutoScan();
}

// stopAutoScan - stop auto scan
function stopAutoScan() {
  deleteLinkFindTriggers_();
  clearAllProgressProps_();

  SpreadsheetApp.getUi().alert(
    '自動調査を停止しました。' + '\n' + 'トリガーと進捗データを削除しました。'
  );
}

// runAutoScan - triggered auto scan with resume support
function runAutoScan() {
  var lock = LockService.getScriptLock();
  var hasLock = lock.tryLock(10000);
  if (!hasLock) {
    Logger.log('LinkFind: Another instance is running. Skipping.');
    return;
  }

  try {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty(PROP_KEYS.COMPLETED) === '1') {
      Logger.log('LinkFind: Already completed. Deleting triggers.');
      deleteLinkFindTriggers_();
      return;
    }

    var startTime = new Date().getTime();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var fileIds = getFileIds_(ss, false);
    if (fileIds.length === 0) {
      Logger.log('LinkFind: No FileIDs found.');
      deleteLinkFindTriggers_();
      clearAllProgressProps_();
      return;
    }

    var resumeFileIdx = getResumeInt_(PROP_KEYS.RESUME_FILE_INDEX);
    var resumeSheetIdx = getResumeInt_(PROP_KEYS.RESUME_SHEET_INDEX);
    var resumeRowIdx = getResumeInt_(PROP_KEYS.RESUME_ROW_INDEX);
    var totalCount = fileIds.length;
    var isFirstRun = (resumeFileIdx === 0 && resumeSheetIdx === 0 && resumeRowIdx === 0);

    var resultSheet;
    var errorSheet;
    if (isFirstRun) {
      resultSheet = prepareResultSheet_(ss, true);
      errorSheet = prepareErrorSheet_(ss, true);
      props.setProperty(PROP_KEYS.TOTAL_COUNT, String(totalCount));
      Logger.log('LinkFind: Starting auto scan - total ' + totalCount + ' files');
    } else {
      resultSheet = prepareResultSheet_(ss, false);
      errorSheet = prepareErrorSheet_(ss, false);
      Logger.log('LinkFind: Resuming from file ' + resumeFileIdx + ', sheet ' + resumeSheetIdx + ', row ' + resumeRowIdx);
    }

    writeStatus_(resultSheet, resumeFileIdx, totalCount, 'Processing...');

    var batchResults = [];
    var batchErrors = [];

    for (var i = resumeFileIdx; i < fileIds.length; i++) {
      var elapsed = (new Date().getTime() - startTime) / 1000;
      if (elapsed >= CONFIG.TIME_LIMIT_SEC) {
        props.setProperty(PROP_KEYS.RESUME_FILE_INDEX, String(i));
        props.deleteProperty(PROP_KEYS.RESUME_SHEET_INDEX);
        props.deleteProperty(PROP_KEYS.RESUME_ROW_INDEX);
        Logger.log('LinkFind: Time limit at file boundary - ' + i + '/' + totalCount);

        if (batchResults.length > 0) {
          appendResultRows_(resultSheet, batchResults);
        }
        if (batchErrors.length > 0) {
          appendErrorRows_(errorSheet, batchErrors);
        }
        writeStatus_(resultSheet, i, totalCount, 'Suspended - next trigger');
        return;
      }

      var fileId = fileIds[i].trim();
      if (fileId === '') continue;

      Logger.log('Processing: ' + (i + 1) + '/' + totalCount + ' - ' + fileId);
      writeStatus_(resultSheet, i, totalCount, 'Processing...');

      var sheetStart = (i === resumeFileIdx && resumeSheetIdx > 0) ? resumeSheetIdx : null;
      var rowStart = (i === resumeFileIdx && resumeRowIdx > 0) ? resumeRowIdx : null;

      var scanResult = scanFileForImportRange_(fileId, ss, startTime, sheetStart, rowStart);

      if (scanResult.error) {
        batchErrors.push(scanResult.error);
      }
      if (scanResult.results.length > 0) {
        batchResults = batchResults.concat(scanResult.results);
      }

      addDiscoveredIds_(ss, scanResult.newIds);

      if (scanResult.interrupted) {
        props.setProperty(PROP_KEYS.RESUME_FILE_INDEX, String(i));
        props.setProperty(PROP_KEYS.RESUME_SHEET_INDEX, String(scanResult.lastSheetIdx));
        props.setProperty(PROP_KEYS.RESUME_ROW_INDEX, String(scanResult.lastRowIdx));
        Logger.log('LinkFind: Time limit mid-file - file ' + i + ', sheet ' + scanResult.lastSheetIdx + ', row ' + scanResult.lastRowIdx);

        if (batchResults.length > 0) {
          appendResultRows_(resultSheet, batchResults);
        }
        if (batchErrors.length > 0) {
          appendErrorRows_(errorSheet, batchErrors);
        }
        writeStatus_(resultSheet, i, totalCount, 'Suspended - next trigger');
        return;
      }

      resumeSheetIdx = 0;
      resumeRowIdx = 0;
    }

    fileIds = getFileIds_(ss, false);
    totalCount = fileIds.length;

    if (batchResults.length > 0) {
      appendResultRows_(resultSheet, batchResults);
    }
    if (batchErrors.length > 0) {
      appendErrorRows_(errorSheet, batchErrors);
    }

    var allResults = readAllResultsFromSheet_(resultSheet);
    outputCsvFiles_(allResults);

    writeStatus_(resultSheet, totalCount, totalCount, 'Completed');
    Logger.log('LinkFind: All done - ' + totalCount + ' files');

    props.setProperty(PROP_KEYS.COMPLETED, '1');
    deleteLinkFindTriggers_();
    clearAllProgressProps_();
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// Core scanning logic
// ============================================================

// scanFileForImportRange_ - scan a single spreadsheet for IMPORTRANGE formulas
function scanFileForImportRange_(fileId, ss, startTime, resumeSheetIdx, resumeRowIdx) {
  var output = {
    results: [],
    error: null,
    newIds: [],
    interrupted: false,
    lastSheetIdx: 0,
    lastRowIdx: 0
  };

  var fileName = '';
  var targetSs;

  try {
    targetSs = SpreadsheetApp.openById(fileId);
    fileName = targetSs.getName();
  } catch (e) {
    var errMsg = String(e.message || e);
    var errorType = 'openById failed';
    if (errMsg.indexOf('not found') !== -1 || errMsg.indexOf('404') !== -1) {
      errorType = 'File not found';
    } else if (errMsg.indexOf('403') !== -1 || errMsg.indexOf('permission') !== -1 ||
               errMsg.indexOf('access') !== -1 || errMsg.indexOf('forbidden') !== -1) {
      errorType = 'Access denied';
    }
    output.error = {
      fileId: fileId,
      fileName: fileName,
      errorMessage: errorType + ': ' + errMsg,
      timestamp: new Date()
    };
    return output;
  }

  var existingIds = getFileIdSet_(ss);

  var sheets = targetSs.getSheets();
  var sheetStart = (resumeSheetIdx !== null && resumeSheetIdx !== undefined) ? resumeSheetIdx : 0;

  for (var s = sheetStart; s < sheets.length; s++) {
    var sheet = sheets[s];
    var sheetName = sheet.getName();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow === 0 || lastCol === 0) continue;

    var rowStart = (s === sheetStart && resumeRowIdx !== null && resumeRowIdx !== undefined) ? resumeRowIdx : 0;

    for (var batchStart = rowStart; batchStart < lastRow; batchStart += CONFIG.ROWS_PER_BATCH) {
      if (startTime !== null) {
        var elapsed = (new Date().getTime() - startTime) / 1000;
        if (elapsed >= CONFIG.TIME_LIMIT_SEC) {
          output.interrupted = true;
          output.lastSheetIdx = s;
          output.lastRowIdx = batchStart;
          return output;
        }
      }

      var batchEnd = Math.min(batchStart + CONFIG.ROWS_PER_BATCH, lastRow);
      var numRows = batchEnd - batchStart;

      var formulas;
      try {
        formulas = sheet.getRange(batchStart + 1, 1, numRows, lastCol).getFormulas();
      } catch (e2) {
        Logger.log('getFormulas error: ' + fileId + ' / ' + sheetName + ' row ' + batchStart + ': ' + e2.message);
        continue;
      }

      for (var r = 0; r < formulas.length; r++) {
        for (var c = 0; c < formulas[r].length; c++) {
          var formula = formulas[r][c];
          if (formula === '') continue;

          var upperFormula = formula.toUpperCase();
          if (upperFormula.indexOf('IMPORTRANGE') === -1) continue;

          var importRangeMatches = extractImportRangeCalls_(formula);

          for (var m = 0; m < importRangeMatches.length; m++) {
            var irCall = importRangeMatches[m];
            var linkedId = resolveImportRangeId_(irCall, sheet, batchStart + r + 1);
            var cellRef = columnToLetter_(c + 1) + (batchStart + r + 1);
            var linkedName = '';

            if (linkedId) {
              linkedName = getFileNameSafe_(linkedId);

              var isNew = false;
              if (!existingIds[linkedId]) {
                isNew = true;
                existingIds[linkedId] = true;
                output.newIds.push(linkedId);
              }

              output.results.push({
                sourceFileId: fileId,
                sourceFileName: fileName,
                sheetName: sheetName,
                cellRef: cellRef,
                formula: formula,
                linkedFileId: linkedId,
                linkedFileName: linkedName,
                isNew: isNew ? 'Yes' : ''
              });
            } else {
              output.results.push({
                sourceFileId: fileId,
                sourceFileName: fileName,
                sheetName: sheetName,
                cellRef: cellRef,
                formula: formula,
                linkedFileId: '(unresolved)',
                linkedFileName: '',
                isNew: ''
              });
            }
          }
        }
      }
    }
  }

  return output;
}

// ============================================================
// IMPORTRANGE parsing
// ============================================================

// extractImportRangeCalls_ - extract IMPORTRANGE first-argument strings from a formula
function extractImportRangeCalls_(formula) {
  var results = [];
  var upper = formula.toUpperCase();
  var searchPos = 0;

  while (true) {
    var idx = upper.indexOf('IMPORTRANGE', searchPos);
    if (idx === -1) break;

    var parenStart = formula.indexOf('(', idx);
    if (parenStart === -1) break;

    var firstArg = extractFirstArg_(formula, parenStart);
    if (firstArg !== null) {
      results.push(firstArg);
    }

    searchPos = parenStart + 1;
  }

  return results;
}

// extractFirstArg_ - extract the first argument from a function call starting at openParen
function extractFirstArg_(formula, openParen) {
  var pos = openParen + 1;
  var len = formula.length;
  var depth = 0;
  var inQuote = false;

  while (pos < len && formula.charAt(pos) === ' ') {
    pos++;
  }
  var argStart = pos;

  while (pos < len) {
    var ch = formula.charAt(pos);

    if (inQuote) {
      if (ch === '"') {
        if (pos + 1 < len && formula.charAt(pos + 1) === '"') {
          pos += 2;
          continue;
        }
        inQuote = false;
      }
      pos++;
      continue;
    }

    if (ch === '"') {
      inQuote = true;
      pos++;
      continue;
    }

    if (ch === '(') {
      depth++;
      pos++;
      continue;
    }

    if (ch === ')') {
      if (depth > 0) {
        depth--;
        pos++;
        continue;
      }
      return formula.substring(argStart, pos).trim();
    }

    if (ch === ',' && depth === 0) {
      return formula.substring(argStart, pos).trim();
    }

    pos++;
  }

  return null;
}

// resolveImportRangeId_ - resolve the first argument to a spreadsheet ID
function resolveImportRangeId_(firstArg, sheet, rowNum) {
  if (!firstArg || firstArg === '') return null;

  var arg = firstArg.trim();
  if (arg.charAt(0) === '"' && arg.charAt(arg.length - 1) === '"') {
    arg = arg.substring(1, arg.length - 1);
    arg = arg.replace(/""/g, '"');
  }

  if (arg.indexOf('https://') === 0 || arg.indexOf('http://') === 0) {
    return extractIdFromUrl_(arg);
  }

  if (isCellReference_(arg)) {
    return resolveCellReference_(arg, sheet);
  }

  if (arg.length > 5 && /^[A-Za-z0-9_-]+$/.test(arg)) {
    return arg;
  }

  return null;
}

// extractIdFromUrl_ - extract spreadsheet ID from a Google Sheets URL
function extractIdFromUrl_(url) {
  var match = url.match(/\/spreadsheets\/d\/([A-Za-z0-9_-]+)/);
  if (match) {
    return match[1];
  }
  var keyMatch = url.match(/[?&]key=([A-Za-z0-9_-]+)/);
  if (keyMatch) {
    return keyMatch[1];
  }
  return null;
}

// isCellReference_ - check if a string looks like a cell reference
function isCellReference_(str) {
  var cellPattern = /^\$?[A-Za-z]{1,3}\$?[0-9]+$/;
  if (cellPattern.test(str)) return true;
  var sheetCellPattern = /^.+!\$?[A-Za-z]{1,3}\$?[0-9]+$/;
  return sheetCellPattern.test(str);
}

// resolveCellReference_ - get the value from a cell reference and extract ID
function resolveCellReference_(ref, sheet) {
  try {
    var cellValue;
    if (ref.indexOf('!') !== -1) {
      var bangIdx = ref.indexOf('!');
      var sheetRef = ref.substring(0, bangIdx).replace(/'/g, '');
      var cellRef = ref.substring(bangIdx + 1);
      var parentSs = sheet.getParent();
      var targetSheet = parentSs.getSheetByName(sheetRef);
      if (targetSheet) {
        cellValue = String(targetSheet.getRange(cellRef).getValue());
      } else {
        return null;
      }
    } else {
      cellValue = String(sheet.getRange(ref).getValue());
    }

    if (!cellValue || cellValue === '') return null;

    if (cellValue.indexOf('https://') === 0 || cellValue.indexOf('http://') === 0) {
      return extractIdFromUrl_(cellValue);
    }

    cellValue = cellValue.trim();
    if (cellValue.length > 5 && /^[A-Za-z0-9_-]+$/.test(cellValue)) {
      return cellValue;
    }

    return null;
  } catch (e) {
    Logger.log('resolveCellReference_ error: ' + ref + ' - ' + e.message);
    return null;
  }
}

// ============================================================
// FileID management
// ============================================================

// getFileIds_ - get FileID list from input sheet
function getFileIds_(ss, useUi) {
  var sheet = ss.getSheetByName(CONFIG.INPUT_SHEET_NAME);
  if (!sheet) {
    if (useUi) {
      SpreadsheetApp.getUi().alert(
        CONFIG.INPUT_SHEET_NAME + 'シートが見つかりません。' + '\n' +
        'シートを作成してA列にFileIDを記載してください。'
      );
    } else {
      Logger.log('LinkFind: Input sheet not found - ' + CONFIG.INPUT_SHEET_NAME);
    }
    return [];
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var ids = [];
  for (var i = 0; i < data.length; i++) {
    var val = String(data[i][0]).trim();
    if (val !== '' && val !== 'undefined' && val !== 'null') {
      ids.push(val);
    }
  }
  return ids;
}

// getFileIdSet_ - get existing FileIDs as a lookup set
function getFileIdSet_(ss) {
  var ids = getFileIds_(ss, false);
  var set = {};
  for (var i = 0; i < ids.length; i++) {
    set[ids[i].trim()] = true;
  }
  return set;
}

// addDiscoveredIds_ - append newly discovered IDs to the input sheet
function addDiscoveredIds_(ss, newIds) {
  if (!newIds || newIds.length === 0) return;

  var sheet = ss.getSheetByName(CONFIG.INPUT_SHEET_NAME);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  for (var i = 0; i < newIds.length; i++) {
    sheet.getRange(lastRow + 1 + i, 1).setValue(newIds[i]);
  }
}

// getFileNameSafe_ - try to get file name, return empty on failure
function getFileNameSafe_(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    return file.getName();
  } catch (e) {
    return '';
  }
}

// ============================================================
// Progress management
// ============================================================

// getResumeInt_ - get integer from properties
function getResumeInt_(key) {
  var val = PropertiesService.getScriptProperties().getProperty(key);
  return val ? parseInt(val, 10) : 0;
}

// clearAllProgress - UI wrapper for clearing progress
function clearAllProgress() {
  clearAllProgressProps_();
  SpreadsheetApp.getUi().alert('Progress data cleared.');
}

// clearAllProgressProps_ - clear all progress properties
function clearAllProgressProps_() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_KEYS.RESUME_FILE_INDEX);
  props.deleteProperty(PROP_KEYS.RESUME_SHEET_INDEX);
  props.deleteProperty(PROP_KEYS.RESUME_ROW_INDEX);
  props.deleteProperty(PROP_KEYS.TOTAL_COUNT);
  props.deleteProperty(PROP_KEYS.COMPLETED);
  props.deleteProperty(PROP_KEYS.DISCOVERED_IDS);
}

// ============================================================
// Trigger management
// ============================================================

// deleteLinkFindTriggers_ - delete all runAutoScan triggers
function deleteLinkFindTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runAutoScan') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// ============================================================
// Status display
// ============================================================

// writeStatus_ - show progress in G1 of result sheet
function writeStatus_(sheet, processed, total, status) {
  var statusText = processed + '/' + total + '件処理済み - ' + status;
  sheet.getRange(1, 7).setValue(statusText);
  sheet.setColumnWidth(7, 350);
  Logger.log('LinkFind: ' + statusText);
}

// ============================================================
// Result sheet setup
// ============================================================

// prepareResultSheet_ - prepare or create the result sheet
function prepareResultSheet_(ss, clearSheet) {
  var sheet = ss.getSheetByName(CONFIG.RESULT_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.RESULT_SHEET_NAME);
    clearSheet = true;
  }

  if (clearSheet) {
    sheet.clear();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .setBackground(null);

    var headers = [
      '元FileID',
      '元ファイル名',
      'シート名',
      'セル位置',
      'IMPORTRANGE数式',
      'リンク先FileID',
      'リンク先ファイル名',
      '新規追加'
    ];
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground(CONFIG.HEADER_COLOR);
    headerRange.setFontColor(CONFIG.HEADER_FONT_COLOR);
    headerRange.setFontWeight('bold');

    sheet.setColumnWidth(1, 320);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(5, 400);
    sheet.setColumnWidth(6, 320);
    sheet.setColumnWidth(7, 250);
    sheet.setColumnWidth(8, 80);

    sheet.setFrozenRows(1);
  }

  return sheet;
}

// prepareErrorSheet_ - prepare or create the error sheet
function prepareErrorSheet_(ss, clearSheet) {
  var sheet = ss.getSheetByName(CONFIG.ERROR_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.ERROR_SHEET_NAME);
    clearSheet = true;
  }

  if (clearSheet) {
    sheet.clear();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .setBackground(null);

    var headers = ['FileID', 'ファイル名', 'エラー内容', '発生日時'];
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground(CONFIG.HEADER_COLOR);
    headerRange.setFontColor(CONFIG.HEADER_FONT_COLOR);
    headerRange.setFontWeight('bold');

    sheet.setColumnWidth(1, 320);
    sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 400);
    sheet.setColumnWidth(4, 160);

    sheet.setFrozenRows(1);
  }

  return sheet;
}

// ============================================================
// Result output
// ============================================================

// writeResultRows_ - write results from specified start row
function writeResultRows_(sheet, results, startRow) {
  if (results.length === 0) return;

  var rows = buildResultRows_(results);
  var dataRange = sheet.getRange(startRow, 1, rows.length, rows[0].length);
  dataRange.setValues(rows);
}

// appendResultRows_ - append results below existing data
function appendResultRows_(sheet, results) {
  if (results.length === 0) return;

  var rows = buildResultRows_(results);
  var lastRow = sheet.getLastRow();
  var startRow = (lastRow < 1) ? 2 : lastRow + 1;

  var dataRange = sheet.getRange(startRow, 1, rows.length, rows[0].length);
  dataRange.setValues(rows);
}

// buildResultRows_ - build 2D array from results
function buildResultRows_(results) {
  var rows = [];
  for (var i = 0; i < results.length; i++) {
    var r = results[i];
    rows.push([
      r.sourceFileId,
      r.sourceFileName,
      r.sheetName,
      r.cellRef,
      r.formula,
      r.linkedFileId,
      r.linkedFileName,
      r.isNew
    ]);
  }
  return rows;
}

// writeErrorRows_ - write error rows from row 2
function writeErrorRows_(sheet, errors) {
  if (errors.length === 0) return;

  var rows = buildErrorRows_(errors);
  var dataRange = sheet.getRange(2, 1, rows.length, rows[0].length);
  dataRange.setValues(rows);
}

// appendErrorRows_ - append error rows below existing data
function appendErrorRows_(sheet, errors) {
  if (errors.length === 0) return;

  var rows = buildErrorRows_(errors);
  var lastRow = sheet.getLastRow();
  var startRow = (lastRow < 1) ? 2 : lastRow + 1;

  var dataRange = sheet.getRange(startRow, 1, rows.length, rows[0].length);
  dataRange.setValues(rows);
}

// buildErrorRows_ - build 2D array from error objects
function buildErrorRows_(errors) {
  var rows = [];
  for (var i = 0; i < errors.length; i++) {
    var e = errors[i];
    rows.push([
      e.fileId,
      e.fileName,
      e.errorMessage,
      formatDate_(e.timestamp)
    ]);
  }
  return rows;
}

// ============================================================
// CSV output
// ============================================================

// outputCsvFiles_ - output CSV files to the configured Drive folder
function outputCsvFiles_(results) {
  if (!CONFIG.OUTPUT_FOLDER_ID || CONFIG.OUTPUT_FOLDER_ID === '') {
    Logger.log('LinkFind: OUTPUT_FOLDER_ID not set. Skipping CSV output.');
    return;
  }

  if (results.length === 0) {
    Logger.log('LinkFind: No results to output as CSV.');
    return;
  }

  var folder;
  try {
    folder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
  } catch (e) {
    Logger.log('LinkFind: Cannot access output folder - ' + e.message);
    return;
  }

  var dateStr = formatDateCompact_(new Date());
  var bom = '\uFEFF';

  var csvHeader = '元FileID,元ファイル名,シート名,セル位置,IMPORTRANGE数式,リンク先FileID,リンク先ファイル名,新規追加';
  var csvLines = [csvHeader];
  for (var i = 0; i < results.length; i++) {
    var r = results[i];
    csvLines.push(
      escapeCsvField_(r.sourceFileId) + ',' +
      escapeCsvField_(r.sourceFileName) + ',' +
      escapeCsvField_(r.sheetName) + ',' +
      escapeCsvField_(r.cellRef) + ',' +
      escapeCsvField_(r.formula) + ',' +
      escapeCsvField_(r.linkedFileId) + ',' +
      escapeCsvField_(r.linkedFileName) + ',' +
      escapeCsvField_(r.isNew)
    );
  }

  var unifiedFileName = 'LinkFind_result_' + dateStr + '.csv';
  var unifiedContent = bom + csvLines.join('\r\n');
  folder.createFile(unifiedFileName, unifiedContent, MimeType.PLAIN_TEXT);
  Logger.log('LinkFind: Created ' + unifiedFileName);

  var fileGroups = {};
  for (var j = 0; j < results.length; j++) {
    var res = results[j];
    var key = res.sourceFileId;
    if (!fileGroups[key]) {
      fileGroups[key] = {
        fileName: res.sourceFileName,
        rows: []
      };
    }
    fileGroups[key].rows.push(res);
  }

  var fileKeys = Object.keys(fileGroups);
  for (var k = 0; k < fileKeys.length; k++) {
    var group = fileGroups[fileKeys[k]];
    var safeName = sanitizeFileName_(group.fileName);
    var perFileLines = [csvHeader];

    for (var l = 0; l < group.rows.length; l++) {
      var row = group.rows[l];
      perFileLines.push(
        escapeCsvField_(row.sourceFileId) + ',' +
        escapeCsvField_(row.sourceFileName) + ',' +
        escapeCsvField_(row.sheetName) + ',' +
        escapeCsvField_(row.cellRef) + ',' +
        escapeCsvField_(row.formula) + ',' +
        escapeCsvField_(row.linkedFileId) + ',' +
        escapeCsvField_(row.linkedFileName) + ',' +
        escapeCsvField_(row.isNew)
      );
    }

    var perFileName = 'LinkFind_' + safeName + '_' + dateStr + '.csv';
    var perFileContent = bom + perFileLines.join('\r\n');
    folder.createFile(perFileName, perFileContent, MimeType.PLAIN_TEXT);
    Logger.log('LinkFind: Created ' + perFileName);
  }
}

// escapeCsvField_ - escape a field for CSV output
function escapeCsvField_(value) {
  if (value === null || value === undefined) return '';
  var str = String(value);
  if (str.indexOf(',') !== -1 || str.indexOf('"') !== -1 || str.indexOf('\n') !== -1 || str.indexOf('\r') !== -1) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

// sanitizeFileName_ - remove characters unsafe for file names
function sanitizeFileName_(name) {
  if (!name) return 'unknown';
  return name.replace(/[/\:*?"<>|]/g, '_').substring(0, 50);
}

// readAllResultsFromSheet_ - read all result rows back from the sheet
function readAllResultsFromSheet_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var results = [];
  for (var i = 0; i < data.length; i++) {
    results.push({
      sourceFileId: data[i][0],
      sourceFileName: data[i][1],
      sheetName: data[i][2],
      cellRef: data[i][3],
      formula: data[i][4],
      linkedFileId: data[i][5],
      linkedFileName: data[i][6],
      isNew: data[i][7]
    });
  }
  return results;
}

// ============================================================
// Utility
// ============================================================

// formatDate_ - format date for display
function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
}

// formatDateCompact_ - format date for file names (YYYYMMDD)
function formatDateCompact_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
}

// columnToLetter_ - convert column number to letter (1=A, 27=AA)
function columnToLetter_(col) {
  var letter = '';
  var temp;
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}
