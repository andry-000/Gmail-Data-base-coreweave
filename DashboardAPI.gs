/**
 * DashboardAPI.gs
 * =====================================================================
 * The web app calls these functions via google.script.run to get
 * data, status, and downloads. All read paths operate on the two
 * data sheets.
 *
 * Search tokens supported by `runSearch`:
 *   Subject: <text>   From: <text>      To: <text>
 *   CC: <text>        BCC: <text>       Label: <text>
 *   PDF: <text>       Doc: <text>       Filename: <text>  Attachment: <text>
 *   Has: <zip|nested|attachment|unread|important|starred|inbox|spam|trash|draft|sent>
 *   Month: <YYYY-MM>  Date: <YYYY-MM-DD or YYYY-MM>
 *   Address: <text>   Body: <text>
 *
 * Multiple tokens are AND-combined. Free text (no colon) is searched
 * across subject + from + to + filename + body.
 * =====================================================================
 */

const PROP_RUNNING = 'INGESTION_RUNNING';
const PROP_TRIGGER_ID = 'INGESTION_TRIGGER_ID';
const PROP_TRIGGER_FIRE_AT = 'INGESTION_TRIGGER_FIRE_AT';

/** Returns the live state of the ingestion cursor + last run info. */
function apiGetStatus() {
  const cfg = loadConfigOverridesFromSheet_();
  const state = loadState_();
  const props = PropertiesService.getScriptProperties();
  let isRunning = props.getProperty(PROP_RUNNING) === 'true';
  const fireAt = Number(props.getProperty(PROP_TRIGGER_FIRE_AT) || 0);
  if (isRunning && fireAt && (Date.now() - fireAt) > 30 * 60 * 1000) {
    log_('WARN', 'Running flag stuck for more than 30 min, clearing');
    try { props.deleteProperty(PROP_RUNNING); } catch (e) {}
    isRunning = false;
  }

  let emailCount = 0, docCount = 0;
  try {
    const ss = getOrCreateSpreadsheet_();
    const es = ss.getSheetByName(cfg.EMAILS_SHEET);
    const ds = ss.getSheetByName(cfg.DOCUMENTS_SHEET);
    if (es) emailCount = Math.max(0, es.getLastRow() - 1);
    if (ds) docCount = Math.max(0, ds.getLastRow() - 1);
  } catch (e) {
    log_('WARN', 'apiGetStatus could not read sheet counts', e);
  }

  return {
    state: state,
    config: {
      START_DATE: cfg.START_DATE,
      TARGET_ADDRESSES: cfg.TARGET_ADDRESSES,
      SPREADSHEET_NAME: cfg.SPREADSHEET_NAME
    },
    counts: { emails: emailCount, documents: docCount },
    query: cfg.buildGmailQuery(),
    isRunning: isRunning
  };
}

/** Trigger a fresh ingestion run. Kicks off a one-shot trigger so the
 *  web app call returns immediately instead of blocking on the
 *  6-min Apps Script execution ceiling. The dashboard polls
 *  apiGetStatus() to see progress. */
function apiStartIngestion() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_RUNNING) === 'true') {
    log_('INFO', 'Ingestion already running (property flag set)');
    return { ok: true, status: 'already-running' };
  }
  props.setProperty(PROP_RUNNING, 'true');
  props.setProperty(PROP_TRIGGER_FIRE_AT, String(Date.now() + 3000));
  try {
    const trig = ScriptApp.newTrigger(CFG.RESUME_HANDLER)
      .timeBased()
      .after(3 * 1000)
      .create();
    props.setProperty(PROP_TRIGGER_ID, trig.getUniqueId());
    log_('INFO', 'Kicked off ingestion via one-shot trigger', { id: trig.getUniqueId() });
  } catch (e) {
    try { props.deleteProperty(PROP_RUNNING); } catch (e2) {}
    log_('ERROR', 'Failed to schedule ingestion trigger', e);
    return { ok: false, error: String(e && e.message || e) };
  }
  return { ok: true, status: 'scheduled' };
}

/** Reset ingestion cursor. Does NOT delete sheet rows. */
function apiResetIngestion() {
  try { resetIngestion(); } catch (e) { log_('WARN', 'resetIngestion', e); }
  try { cancelResumableTrigger(); } catch (e) { log_('WARN', 'cancelResumableTrigger', e); }
  const props = PropertiesService.getScriptProperties();
  try { props.deleteProperty(PROP_RUNNING); } catch (e) {}
  try { props.deleteProperty(PROP_TRIGGER_ID); } catch (e) {}
  try { props.deleteProperty(PROP_TRIGGER_FIRE_AT); } catch (e) {}
  return { ok: true };
}

/** Main filter+search API. tokens: string[] of "Field: value". */
function apiSearch(tokens, dateFrom, dateTo) {
  return runSearch_(tokens, dateFrom, dateTo, false);
}

/** Returns CSV text of the current filtered view. */
function apiExportCsv(tokens, dateFrom, dateTo) {
  const res = runSearch_(tokens, dateFrom, dateTo, true);
  return res.csv;
}

/** Returns a base64-encoded XLSX of the current filtered view. */
function apiExportXlsx(tokens, dateFrom, dateTo) {
  const res = runSearch_(tokens, dateFrom, dateTo, true);
  return _buildXlsxFromRows_(res.headers, res.rows);
}

// ---- Internals ----

function runSearch_(tokens, dateFrom, dateTo, includeDocs) {
  const cfg = loadConfigOverridesFromSheet_();
  let sh = null;
  try { sh = getOrCreateSpreadsheet_().getSheetByName(cfg.EMAILS_SHEET); } catch (e) { log_('WARN', 'runSearch_ open', e); }
  if (!sh || sh.getLastRow() < 2) {
    return _emptySearchResult_();
  }
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, SHEET_HEADERS.Emails.length).getValues();
  const parsedTokens = _parseSearchTokens_(tokens || []);
  const fromMs = dateFrom ? Date.parse(dateFrom) : null;
  const toMs = dateTo ? Date.parse(dateTo) + 86399999 : null;

  const rows = [];
  const monthly = {};
  let emailHits = 0;
  for (let i = 0; i < data.length; i++) {
    const rObj = _rowToObject_(SHEET_HEADERS.Emails, data[i]);
    if (!_matchesAll_(rObj, parsedTokens)) continue;
    const ms = Number(rObj.dateMs) || 0;
    if (fromMs !== null && ms < fromMs) continue;
    if (toMs !== null && ms > toMs) continue;
    rows.push(data[i]);
    emailHits++;
    const m = rObj.month || fmtMonth_(ms);
    if (m) {
      if (!monthly[m]) monthly[m] = { emails: 0, documents: 0 };
      monthly[m].emails++;
    }
  }

  const docSheet = (() => {
    try { return getOrCreateSpreadsheet_().getSheetByName(cfg.DOCUMENTS_SHEET); }
    catch (e) { return null; }
  })();
  const docRows = [];
  let docHits = 0;
  if (docSheet && docSheet.getLastRow() >= 2) {
    const ddata = docSheet.getRange(2, 1, docSheet.getLastRow() - 1, SHEET_HEADERS.Documents.length).getValues();
    const passingIds = {};
    for (let i = 0; i < rows.length; i++) passingIds[rows[i][0]] = true;
    for (let i = 0; i < ddata.length; i++) {
      const drObj = _rowToObject_(SHEET_HEADERS.Documents, ddata[i]);
      if (passingIds[drObj.messageId]) {
        docRows.push(ddata[i]);
        docHits++;
        const m = drObj.month || fmtMonth_(drObj.emailDateMs);
        if (m) {
          if (!monthly[m]) monthly[m] = { emails: 0, documents: 0 };
          monthly[m].documents++;
        }
      } else if (_docMatchesTokens_(drObj, parsedTokens)) {
        docRows.push(ddata[i]);
        docHits++;
        if (!passingIds[drObj.messageId]) {
          for (let j = 0; j < data.length; j++) {
            if (data[j][0] === drObj.messageId) { rows.push(data[j]); passingIds[drObj.messageId] = true; break; }
          }
        }
        const m = drObj.month || fmtMonth_(drObj.emailDateMs);
        if (m) {
          if (!monthly[m]) monthly[m] = { emails: 0, documents: 0 };
          monthly[m].documents++;
        }
      }
    }
  }

  const MAX_ROWS = 2000;
  const trimmedRows = rows.slice(0, MAX_ROWS);
  const trimmedDocRows = docRows.slice(0, MAX_ROWS);

  if (!includeDocs) {
    return {
      headers: SHEET_HEADERS.Emails,
      rows: trimmedRows,
      totals: { emails: emailHits, documents: docHits, returned: trimmedRows.length, truncated: rows.length > MAX_ROWS },
      monthly: monthly
    };
  }

  const csv = _buildCsv_(SHEET_HEADERS.Emails, rows, SHEET_HEADERS.Documents, docRows);
  return {
    headers: SHEET_HEADERS.Emails,
    rows: trimmedRows,
    docHeaders: SHEET_HEADERS.Documents,
    docRows: trimmedDocRows,
    csv: csv,
    totals: { emails: emailHits, documents: docHits, returned: trimmedRows.length, truncated: rows.length > MAX_ROWS },
    monthly: monthly
  };
}

function _emptySearchResult_() {
  return {
    headers: SHEET_HEADERS.Emails, rows: [],
    docHeaders: SHEET_HEADERS.Documents, docRows: [],
    totals: { emails: 0, documents: 0, returned: 0, truncated: false },
    monthly: {}
  };
}

function _rowToObject_(headers, row) {
  const o = {};
  for (let i = 0; i < headers.length; i++) o[headers[i]] = row[i];
  return o;
}

function _parseSearchTokens_(tokens) {
  const out = [];
  for (let i = 0; i < tokens.length; i++) {
    const t = String(tokens[i] || '').trim();
    if (!t) continue;
    const colon = t.indexOf(':');
    if (colon < 0) {
      out.push({ field: 'freetext', value: t.toLowerCase() });
    } else {
      out.push({ field: t.slice(0, colon).trim().toLowerCase(), value: t.slice(colon + 1).trim() });
    }
  }
  return out;
}

function _matchesAll_(rObj, tokens) {
  for (let i = 0; i < tokens.length; i++) {
    if (!_rowMatchesToken_(rObj, tokens[i])) return false;
  }
  return true;
}

function _docMatchesTokens_(dObj, tokens) {
  for (let i = 0; i < tokens.length; i++) {
    if (!_docMatchesToken_(dObj, tokens[i])) return false;
  }
  return true;
}

function _rowMatchesToken_(r, t) {
  const v = (t.value || '').toLowerCase();
  switch (t.field) {
    case 'subject': return _ci_(r.subject).indexOf(v) >= 0;
    case 'from':    return _ci_(r.from).indexOf(v) >= 0;
    case 'to':      return _ci_(r.to).indexOf(v) >= 0;
    case 'cc':      return _ci_(r.cc).indexOf(v) >= 0;
    case 'bcc':     return _ci_(r.bcc).indexOf(v) >= 0;
    case 'reply':   return _ci_(r.replyTo).indexOf(v) >= 0;
    case 'label':   return _ci_(r.labels).indexOf(v) >= 0;
    case 'body':    return _ci_(r.bodyPlain).indexOf(v) >= 0;
    case 'month':   return _ci_(r.month).indexOf(v) >= 0;
    case 'date':    return _ci_(r.month).indexOf(v) >= 0 || _ci_(r.date).indexOf(v) >= 0;
    case 'address': return _ci_((r.from || '') + ' ' + (r.to || '') + ' ' + (r.cc || '') + ' ' + (r.bcc || '') + ' ' + (r.replyTo || '')).indexOf(v) >= 0;
    case 'has': {
      switch (v) {
        case 'zip': return !!r.hasZip;
        case 'nested': case 'nestedemail': return !!r.hasNestedEmail;
        case 'attachment': case 'attachments': return Number(r.attachmentCount) > 0;
        case 'unread': return !!r.isUnread;
        case 'important': return !!r.isImportant;
        case 'starred': return !!r.isStarred;
        case 'inbox': return !!r.isInInbox;
        case 'spam': return !!r.isInSpam;
        case 'trash': return !!r.isInTrash;
        case 'draft': return !!r.isDraft;
        case 'sent': return !!r.isSent;
      }
      return false;
    }
    case 'freetext': {
      const hay = _ci_(
        (r.subject || '') + ' ' + (r.from || '') + ' ' + (r.to || '') +
        ' ' + (r.snippet || '') + ' ' + (r.labels || '') + ' ' + (r.attachmentNames || '')
      );
      return hay.indexOf(v) >= 0;
    }
    default:
      return _ci_(JSON.stringify(r)).indexOf(v) >= 0;
  }
}

function _docMatchesToken_(d, t) {
  const v = (t.value || '').toLowerCase();
  switch (t.field) {
    case 'pdf': case 'doc': case 'filename': case 'attachment':
      return _ci_(d.filename).indexOf(v) >= 0;
    case 'from':    return _ci_(d.emailFrom).indexOf(v) >= 0;
    case 'subject': return _ci_(d.emailSubject).indexOf(v) >= 0;
    case 'month':   return _ci_(d.month).indexOf(v) >= 0;
    case 'date':    return _ci_(d.month).indexOf(v) >= 0;
    case 'has': {
      if (v === 'zip') return _ci_(d.source).indexOf('zip') >= 0;
      if (v === 'nested' || v === 'nestedemail') return _ci_(d.source).indexOf('nested') >= 0;
      return false;
    }
    case 'freetext':
      return _ci_((d.filename || '') + ' ' + (d.emailSubject || '') + ' ' + (d.emailFrom || '')).indexOf(v) >= 0;
    default:
      return _ci_(JSON.stringify(d)).indexOf(v) >= 0;
  }
}

function _ci_(s) { return (s === null || s === undefined) ? '' : String(s).toLowerCase(); }

function _buildCsv_(h1, r1, h2, r2) {
  const lines = [];
  lines.push(h1.map(_csvCell_).join(','));
  for (let i = 0; i < r1.length; i++) lines.push(r1[i].map(_csvCell_).join(','));
  if (h2 && r2) {
    lines.push('');
    lines.push(h2.map(_csvCell_).join(','));
    for (let i = 0; i < r2.length; i++) lines.push(r2[i].map(_csvCell_).join(','));
  }
  return lines.join('\n');
}

function _csvCell_(v) {
  if (v === null || v === undefined) return '';
  const s = String(v);
  if (s.indexOf('"') >= 0 || s.indexOf(',') >= 0 || s.indexOf('\n') >= 0) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function _buildXlsxFromRows_(headers, rows) {
  const ss = SpreadsheetApp.create('__tmp_export_' + Date.now());
  const sh = ss.getActiveSheet();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  const blob = ss.getBlob();
  try { DriveApp.getFileById(ss.getId()).setTrashed(true); } catch (e) {}
  return Utilities.base64Encode(blob.getBytes());
}
