/**
 * SheetManager.gs
 * =====================================================================
 * Owns the auto-created spreadsheet, the two data sheets, the Settings
 * sheet, and all writes. Writes are batched so we never blow past
 * Apps Script's URL-fetch or execution-time limits.
 *
 * IMPORTANT: this file MUST NOT call loadConfigOverridesFromSheet_().
 * It uses hard-coded `CFG` only — that's the architectural change that
 * fixes the previous stack-overflow recursion.
 * =====================================================================
 */

const SHEET_HEADERS = {
  Emails: [
    'messageId', 'threadId', 'date', 'dateMs', 'month',
    'from', 'to', 'cc', 'bcc', 'replyTo', 'subject', 'snippet',
    'labels', 'labelIds',
    'isUnread', 'isStarred', 'isImportant', 'isInInbox',
    'isInTrash', 'isInSpam', 'isDraft', 'isSent',
    'messageSize', 'attachmentCount', 'attachmentNames', 'attachmentMimes',
    'hasZip', 'hasNestedEmail',
    'bodyPlain', 'bodyHtml', 'bodySizePlain', 'bodySizeHtml',
    'rawHeaders', 'headerCount',
    'inquiryAccount', 'firstMatchedAddress',
    'ingestedAt', 'extra'
  ],
  Documents: [
    'docId', 'parentDocId', 'source',
    'messageId', 'threadId', 'emailSubject', 'emailFrom',
    'emailDate', 'emailDateMs', 'month',
    'filename', 'extension', 'mimeType', 'sizeBytes',
    'sha256', 'driveFileId', 'driveFileUrl',
    'parentZipName', 'parentEmailSubject', 'path',
    'ingestedAt'
  ]
};

/** Lazily creates the spreadsheet + the three sheets on first use.
 *  Uses hard-coded CFG only (no Settings sheet read, no recursion). */
function getOrCreateSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  let ss = null;
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }
  if (!ss) {
    let folder = null;
    if (CFG.FOLDER_NAME) {
      const it = DriveApp.getFoldersByName(CFG.FOLDER_NAME);
      if (it.hasNext()) folder = it.next();
      else folder = DriveApp.createFolder(CFG.FOLDER_NAME);
    }
    ss = SpreadsheetApp.create(CFG.SPREADSHEET_NAME);
    if (folder) {
      const file = DriveApp.getFileById(ss.getId());
      file.moveTo(folder);
    }
    props.setProperty('SPREADSHEET_ID', ss.getId());
    log_('INFO', 'Created spreadsheet', { id: ss.getId(), url: ss.getUrl() });
  }
  ensureDataSheets_(ss);
  ensureSettingsSheet_(ss); // uses CFG.SETTINGS_SHEET directly — no config read
  return ss;
}

function ensureDataSheets_(ss) {
  _ensureSheetWithHeaders_(ss, CFG.EMAILS_SHEET, SHEET_HEADERS.Emails);
  _ensureSheetWithHeaders_(ss, CFG.DOCUMENTS_SHEET, SHEET_HEADERS.Documents);
}

function _ensureSheetWithHeaders_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1a1d29')
      .setFontColor('#ffffff');
    sh.setTabColor(name === CFG.EMAILS_SHEET ? '#3b82f6' : '#9ca3af');
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function ensureSettingsSheet_(ss) {
  // Uses hardcoded CFG.SETTINGS_SHEET. Does NOT call
  // loadConfigOverridesFromSheet_ — that's what previously caused
  // infinite recursion. The seeded defaults are the hard-coded ones.
  let sh;
  try { sh = ss.getSheetByName(CFG.SETTINGS_SHEET); } catch (e) { return; }
  if (!sh) {
    try {
      sh = ss.insertSheet(CFG.SETTINGS_SHEET);
      sh.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, 2).setFontWeight('bold');
    } catch (e) {
      log_('WARN', 'Could not create Settings sheet', e);
      return;
    }
  }
  // Seed defaults if the sheet is empty
  if (sh.getLastRow() <= 1) {
    const defaults = [
      ['START_DATE', CFG.START_DATE],
      ['TARGET_ADDRESSES', CFG.TARGET_ADDRESSES.join(', ')],
      ['PAGE_SIZE', String(CFG.PAGE_SIZE)],
      ['MAX_MESSAGES_PER_RUN', String(CFG.MAX_MESSAGES_PER_RUN)],
      ['TIME_BUDGET_SECONDS', String(CFG.TIME_BUDGET_SECONDS)],
      ['RESCHEDULE_MINUTES', String(CFG.RESCHEDULE_MINUTES)],
      ['SPREADSHEET_NAME', CFG.SPREADSHEET_NAME],
      ['VERBOSE_LOGGING', String(CFG.VERBOSE_LOGGING)]
    ];
    try { sh.getRange(2, 1, defaults.length, 2).setValues(defaults); } catch (e) {}
  }
}

/** Append a single row, truncated per cell. */
function appendEmailRow_(rowValues) {
  return _appendRow_(CFG.EMAILS_SHEET, SHEET_HEADERS.Emails.length, rowValues);
}

function appendDocumentRow_(rowValues) {
  return _appendRow_(CFG.DOCUMENTS_SHEET, SHEET_HEADERS.Documents.length, rowValues);
}

function _appendRow_(sheetName, expectedLen, values) {
  try {
    const ss = getOrCreateSpreadsheet_();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) { log_('ERROR', 'Sheet missing: ' + sheetName); return 0; }
    const row = new Array(expectedLen).fill('');
    for (let i = 0; i < Math.min(values.length, expectedLen); i++) {
      let v = values[i];
      if (v === null || v === undefined) v = '';
      row[i] = truncateForSheet_(v, CFG.SHEET_CELL_CHAR_LIMIT);
    }
    sh.appendRow(row);
    return 1;
  } catch (e) {
    log_('ERROR', '_appendRow_ failed for ' + sheetName, e);
    return 0;
  }
}

/** Batch-append many rows in a single setValues call. */
function batchAppendRows_(sheetName, expectedLen, rows) {
  if (!rows || rows.length === 0) return 0;
  try {
    const ss = getOrCreateSpreadsheet_();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) { log_('ERROR', 'Sheet missing: ' + sheetName); return 0; }
    const out = new Array(rows.length);
    for (let r = 0; r < rows.length; r++) {
      const row = new Array(expectedLen).fill('');
      const src = rows[r];
      for (let i = 0; i < Math.min(src.length, expectedLen); i++) {
        let v = src[i];
        if (v === null || v === undefined) v = '';
        row[i] = truncateForSheet_(v, CFG.SHEET_CELL_CHAR_LIMIT);
      }
      out[r] = row;
    }
    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, out.length, expectedLen).setValues(out);
    return out.length;
  } catch (e) {
    log_('ERROR', 'batchAppendRows_ failed for ' + sheetName, e);
    return 0;
  }
}

/** Dedupe check: returns true if messageId is already in Emails sheet. */
function messageAlreadyIngested_(messageId) {
  try {
    const ss = getOrCreateSpreadsheet_();
    const sh = ss.getSheetByName(CFG.EMAILS_SHEET);
    if (!sh || sh.getLastRow() < 2) return false;
    const col = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < col.length; i++) {
      if (col[i][0] === messageId) return true;
    }
    return false;
  } catch (e) {
    log_('WARN', 'messageAlreadyIngested_ check failed', e);
    return false;
  }
}
