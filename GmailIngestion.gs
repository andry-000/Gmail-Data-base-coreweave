/**
 * GmailIngestion.gs
 * =====================================================================
 * The actual email-scanning engine. Pulls matching messages from
 * Gmail in pages, extracts EVERY label and EVERY header the API
 * exposes, then recurses into zip archives and nested email
 * attachments. Writes one row per message to the Emails sheet and
 * one row per "billable" document to the Documents sheet.
 *
 * Resumable: state is persisted in PropertiesService so the next
 * invocation can continue from the last cursor.
 * =====================================================================
 */

const GMAIL_API = 'https://gmail.googleapis.com/gmail/v1/users/me';
const STATE_KEY = 'GMAIL_INGEST_STATE_v1';
const OVERFLOW_FOLDER_NAME = 'Received-Emails-Overflow-Blobs';
let _labelNameCache_ = {};

/**
 * Public entry — called by menu, by trigger, or by the web app's
 * "Start ingestion" button. Returns a status object the UI can show.
 */
function startIngestion() {
  const cfg = loadConfigOverridesFromSheet_();
  const state = loadState_();
  const startMs = Date.now();
  let processed = 0;
  let totalInRun = 0;
  let pageToken = state.pageToken || null;

  log_('INFO', 'Ingestion starting', {
    query: cfg.buildGmailQuery(),
    resume: !!state.pageToken,
    runCount: state.runCount || 0
  });

  while (true) {
    if (totalInRun >= cfg.MAX_MESSAGES_PER_RUN) {
      log_('INFO', 'Hit MAX_MESSAGES_PER_RUN, pausing', { totalInRun });
      break;
    }
    if (timeBudgetExceeded_(startMs, cfg)) {
      log_('WARN', 'Time budget exceeded, pausing', {
        elapsedMs: Date.now() - startMs,
        budgetSec: cfg.TIME_BUDGET_SECONDS
      });
      break;
    }

    let page;
    try {
      page = listMessagesPage_(cfg.buildGmailQuery(), pageToken, cfg.PAGE_SIZE);
    } catch (e) {
      log_('ERROR', 'Gmail list failed', e);
      state.lastError = String(e && e.message || e);
      saveState_(state);
      return { ok: false, error: state.lastError, processed };
    }

    if (!page.messages || page.messages.length === 0) {
      log_('INFO', 'No more messages, ingestion complete', { totalInRun });
      state.pageToken = null;
      state.completedAt = new Date().toISOString();
      state.lastError = null;
      saveState_(state);
      try { PropertiesService.getScriptProperties().deleteProperty(PROP_RUNNING); } catch (e) {}
      return {
        ok: true,
        done: true,
        processed,
        runCount: (state.runCount || 0) + 1
      };
    }

    const batchEmailRows = [];
    const batchDocRows = [];
    for (let i = 0; i < page.messages.length; i++) {
      const meta = page.messages[i];
      if (messageAlreadyIngested_(meta.id)) {
        if (cfg.VERBOSE_LOGGING) log_('DEBUG', 'skip already-ingested', { id: meta.id });
        continue;
      }
      try {
        const result = processOneMessage_(meta.id, cfg);
        if (result.emailRow) batchEmailRows.push(result.emailRow);
        if (result.docRows && result.docRows.length) {
          for (let d = 0; d < result.docRows.length; d++) batchDocRows.push(result.docRows[d]);
        }
        processed++;
        totalInRun++;
      } catch (e) {
        log_('ERROR', 'processOneMessage_ failed for ' + meta.id, e);
        batchEmailRows.push(_stubErrorRow_(meta, e));
        processed++;
        totalInRun++;
      }

      if (totalInRun >= cfg.MAX_MESSAGES_PER_RUN) {
        log_('INFO', 'Hit MAX_MESSAGES_PER_RUN inside page, stopping');
        break;
      }
      if (timeBudgetExceeded_(startMs, cfg)) {
        log_('WARN', 'Time budget exceeded mid-page, will resume next page');
        break;
      }
    }

    if (batchEmailRows.length) {
      batchAppendRows_(CFG.EMAILS_SHEET, SHEET_HEADERS.Emails.length, batchEmailRows);
    }
    if (batchDocRows.length) {
      batchAppendRows_(CFG.DOCUMENTS_SHEET, SHEET_HEADERS.Documents.length, batchDocRows);
    }

    pageToken = page.nextPageToken || null;
    state.pageToken = pageToken;
    state.lastRunAt = new Date().toISOString();
    state.lastError = null;
    state.runCount = (state.runCount || 0) + 1;
    saveState_(state);

    // Mid-pagination safety: if the user clicked Reset, bail out.
    if (PropertiesService.getScriptProperties().getProperty(PROP_RUNNING) !== 'true') {
      log_('INFO', 'Running flag cleared mid-run, stopping');
      return { ok: true, done: true, processed, runCount: state.runCount, stopped: true };
    }

    if (!pageToken) {
      log_('INFO', 'Pagination complete', { totalInRun });
      state.completedAt = new Date().toISOString();
      saveState_(state);
      try { PropertiesService.getScriptProperties().deleteProperty(PROP_RUNNING); } catch (e) {}
      return { ok: true, done: true, processed, runCount: state.runCount };
    }
  }

  // Reschedule the next chunk if there's still work.
  if (state.pageToken) {
    _resumeIn_(cfg);
  } else {
    try { PropertiesService.getScriptProperties().deleteProperty(PROP_RUNNING); } catch (e) {}
  }
  return { ok: true, done: false, processed, runCount: state.runCount };
}

function resumeIngestion() { return startIngestion(); }

function _resumeIn_(cfg) {
  try {
    const existing = ScriptApp.getProjectTriggers();
    for (let i = 0; i < existing.length; i++) {
      if (existing[i].getHandlerFunction() === CFG.RESUME_HANDLER) {
        ScriptApp.deleteTrigger(existing[i]);
      }
    }
  } catch (e) { log_('WARN', 'trigger cleanup failed', e); }
  try {
    ScriptApp.newTrigger(CFG.RESUME_HANDLER)
      .timeBased()
      .after(Math.max(1, cfg.RESCHEDULE_MINUTES) * 60 * 1000)
      .create();
    log_('INFO', 'Resumable trigger installed', { minutes: cfg.RESCHEDULE_MINUTES });
  } catch (e) {
    log_('ERROR', 'Failed to install resumable trigger', e);
  }
}

function cancelResumableTrigger() {
  let n = 0;
  try {
    const existing = ScriptApp.getProjectTriggers();
    for (let i = 0; i < existing.length; i++) {
      if (existing[i].getHandlerFunction() === CFG.RESUME_HANDLER) {
        ScriptApp.deleteTrigger(existing[i]);
        n++;
      }
    }
  } catch (e) { log_('WARN', 'cancelResumableTrigger', e); }
  log_('INFO', 'Cancelled resumable triggers', { count: n });
  return { ok: true, cancelled: n };
}

// ----- State persistence -----

function loadState_() {
  const raw = PropertiesService.getScriptProperties().getProperty(STATE_KEY);
  if (!raw) return {};
  try { return JSON.parse(raw); } catch (e) { return {}; }
}

function saveState_(state) {
  try {
    PropertiesService.getScriptProperties().setProperty(STATE_KEY, JSON.stringify(state));
  } catch (e) { log_('WARN', 'saveState_ failed', e); }
}

function resetIngestion() {
  try { PropertiesService.getScriptProperties().deleteProperty(STATE_KEY); } catch (e) {}
  log_('INFO', 'Ingestion state reset');
}

// ----- Gmail API wrappers -----

function listMessagesPage_(q, pageToken, max) {
  const url = GMAIL_API + '/messages?q=' + encodeURIComponent(q) +
    '&maxResults=' + max +
    (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
  const resp = _gmailFetch_(url);
  return {
    messages: resp.messages || [],
    nextPageToken: resp.nextPageToken || null
  };
}

function getMessageFull_(id) {
  return _gmailFetch_(GMAIL_API + '/messages/' + id + '?format=full');
}

function _gmailFetch_(url) {
  let lastErr = null;
  for (let attempt = 0; attempt < 4; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      const code = resp.getResponseCode();
      if (code === 200) return JSON.parse(resp.getContentText());
      if (code === 429 || code === 500 || code === 503) {
        const wait = Math.pow(2, attempt) * 1000;
        log_('WARN', 'Gmail transient ' + code + ', sleeping ' + wait + 'ms');
        Utilities.sleep(wait);
        lastErr = new Error('Gmail HTTP ' + code);
        continue;
      }
      throw new Error('Gmail HTTP ' + code);
    } catch (e) {
      lastErr = e;
      if (attempt === 3) break;
      Utilities.sleep(Math.pow(2, attempt) * 500);
    }
  }
  throw lastErr || new Error('Gmail fetch failed');
}

// ----- Per-message processing -----

function processOneMessage_(messageId, cfg) {
  const full = getMessageFull_(messageId);
  return _processMessageObject_(full, cfg);
}

function _processMessageObject_(full, cfg) {
  const labelIds = full.labelIds || [];
  const labelsHuman = labelIds.map(_labelIdToHuman_).filter(Boolean);

  const headers = {};
  const rawHeaderArr = (full.payload && full.payload.headers) || [];
  for (let i = 0; i < rawHeaderArr.length; i++) {
    const h = rawHeaderArr[i];
    if (!h || !h.name) continue;
    headers[String(h.name).toLowerCase()] = h.value || '';
  }

  const from = headers['from'] || '';
  const to = headers['to'] || '';
  const cc = headers['cc'] || '';
  const bcc = headers['bcc'] || '';
  const replyTo = headers['reply-to'] || headers['reply'] || '';
  const subject = headers['subject'] || '';
  const dateStr = headers['date'] || '';
  const dateMs = full.internalDate ? Number(full.internalDate) : Date.parse(dateStr) || 0;
  const month = fmtMonth_(dateMs);

  const haystack = (from + '\n' + to + '\n' + cc + '\n' + bcc + '\n' + replyTo + '\n' + subject).toLowerCase();
  let firstMatched = '';
  for (let i = 0; i < cfg.TARGET_ADDRESSES.length; i++) {
    const a = cfg.TARGET_ADDRESSES[i].toLowerCase();
    if (haystack.indexOf(a) >= 0) { firstMatched = cfg.TARGET_ADDRESSES[i]; break; }
  }
  if (!firstMatched) firstMatched = cfg.TARGET_ADDRESSES[0];

  const bodies = _walkForBodies_(full.payload || {});
  const attachments = _walkForAttachments_(full.payload || {}, cfg);

  const isUnread = labelIds.indexOf('UNREAD') >= 0;
  const isStarred = labelIds.indexOf('STARRED') >= 0;
  const isImportant = labelIds.indexOf('IMPORTANT') >= 0;
  const isInInbox = labelIds.indexOf('INBOX') >= 0;
  const isInTrash = labelIds.indexOf('TRASH') >= 0;
  const isInSpam = labelIds.indexOf('SPAM') >= 0;
  const isDraft = labelIds.indexOf('DRAFT') >= 0;
  const isSent = labelIds.indexOf('SENT') >= 0;

  const attachmentNames = attachments.direct.map(a => a.filename);
  const attachmentMimes = attachments.direct.map(a => a.mimeType);
  const hasZip = attachments.direct.some(a => isZipExt_(a.filename, cfg));
  const hasNestedEmail = attachments.direct.some(a => isEmailExt_(a.filename, cfg));

  const rawHeaders = rawHeaderArr.map(h => h.name + ': ' + h.value).join('\n');
  const inquiryAccount = _extractAccountAfterAt_(from) || _extractAccountAfterAt_(replyTo) || firstMatched;

  const emailRow = [
    full.id,
    full.threadId || '',
    fmtDate_(dateMs),
    dateMs,
    month,
    from, to, cc, bcc, replyTo,
    subject,
    full.snippet || '',
    labelsHuman.join(' | '),
    labelIds.join(' | '),
    isUnread, isStarred, isImportant, isInInbox,
    isInTrash, isInSpam, isDraft, isSent,
    full.sizeEstimate || 0,
    attachments.direct.length,
    attachmentNames.join(' | '),
    attachmentMimes.join(' | '),
    hasZip, hasNestedEmail,
    bodies.plain,
    bodies.html,
    bodies.plainSize,
    bodies.htmlSize,
    rawHeaders,
    rawHeaderArr.length,
    inquiryAccount,
    firstMatched,
    new Date().toISOString(),
    JSON.stringify({
      payloadPartCount: _countParts_(full.payload || {}),
      internalDate: full.internalDate,
      historyId: full.historyId,
      threadId: full.threadId
    })
  ];

  // ---- Document rows ----
  const docRows = [];
  for (let i = 0; i < attachments.direct.length; i++) {
    const a = attachments.direct[i];
    if (isExcludedMime_(a.mimeType, cfg)) continue;
    if (isZipExt_(a.filename, cfg)) {
      const inner = _expandZipAttachment_(a, full, cfg);
      for (let k = 0; k < inner.length; k++) {
        docRows.push(_docRowFromBlob_(inner[k], _docCtx_(full, subject, from, dateMs, month, a.filename, ''), cfg));
      }
    } else if (isEmailExt_(a.filename, cfg)) {
      const inner = _expandNestedEmailAttachment_(a, full, cfg);
      for (let k = 0; k < inner.length; k++) {
        docRows.push(_docRowFromBlob_(inner[k], _docCtx_(full, subject, from, dateMs, month, '', subject), cfg));
      }
    } else {
      docRows.push(_docRowFromBlob_({
        filename: a.filename,
        mimeType: a.mimeType,
        bytes: '',
        body: '',
        source: 'direct'
      }, _docCtx_(full, subject, from, dateMs, month, '', ''), cfg));
    }
  }

  return { emailRow, docRows };
}

function _docCtx_(full, subject, from, dateMs, month, parentZipName, parentEmailSubject) {
  return {
    source: parentZipName ? 'zip' : (parentEmailSubject ? 'nestedEmail' : 'direct'),
    messageId: full.id,
    threadId: full.threadId,
    emailSubject: subject,
    emailFrom: from,
    emailDate: fmtDate_(dateMs),
    emailDateMs: dateMs,
    month: month,
    parentZipName: parentZipName,
    parentEmailSubject: parentEmailSubject
  };
}

function _stubErrorRow_(meta, err) {
  const out = new Array(SHEET_HEADERS.Emails.length).fill('');
  out[0] = meta.id;
  out[1] = meta.threadId || '';
  out[37] = new Date().toISOString();
  out[38] = JSON.stringify({ error: String(err && err.message || err) });
  return out;
}

function _labelIdToHuman_(id) {
  if (!id) return '';
  if (id.indexOf('Label_') !== 0) return id;
  if (_labelNameCache_[id]) return _labelNameCache_[id];
  try {
    const resp = _gmailFetch_(GMAIL_API + '/labels/' + encodeURIComponent(id));
    _labelNameCache_[id] = resp.name || id;
    return _labelNameCache_[id];
  } catch (e) {
    _labelNameCache_[id] = id;
    return id;
  }
}

function _extractAccountAfterAt_(s) {
  if (!s) return '';
  const m = String(s).match(/[\w.+-]+@[\w-]+(?:\.[\w-]+)+/);
  return m ? m[0] : '';
}

// ----- MIME tree walkers -----

function _walkForBodies_(part) {
  const out = { plain: '', html: '', plainSize: 0, htmlSize: 0 };
  _walkBodiesRec_(part, out);
  if (out.plain) out.plain = truncateForSheet_(out.plain, CFG.SHEET_CELL_CHAR_LIMIT);
  if (out.html) out.html = truncateForSheet_(out.html, CFG.SHEET_CELL_CHAR_LIMIT);
  return out;
}

function _walkBodiesRec_(part, out) {
  if (!part) return;
  if (part.mimeType === 'text/plain' && part.body && part.body.data) {
    const s = _b64UrlDecode_(part.body.data);
    out.plainSize += s.length;
    out.plain += s + '\n';
  } else if (part.mimeType === 'text/html' && part.body && part.body.data) {
    const s = _b64UrlDecode_(part.body.data);
    out.htmlSize += s.length;
    out.html += s + '\n';
  }
  if (part.parts && part.parts.length) {
    for (let i = 0; i < part.parts.length; i++) _walkBodiesRec_(part.parts[i], out);
  }
}

function _walkForAttachments_(part, cfg) {
  const out = { direct: [] };
  _walkAttachRec_(part, out, cfg);
  return out;
}

function _walkAttachRec_(part, out, cfg) {
  if (!part) return;
  if (part.filename && part.body && part.body.attachmentId) {
    out.direct.push({
      filename: part.filename,
      mimeType: part.mimeType || 'application/octet-stream',
      attachmentId: part.body.attachmentId,
      size: part.body.size || 0
    });
  }
  if (part.parts && part.parts.length) {
    for (let i = 0; i < part.parts.length; i++) _walkAttachRec_(part.parts[i], out, cfg);
  }
}

function _countParts_(part) {
  if (!part) return 0;
  let n = 1;
  if (part.parts) for (let i = 0; i < part.parts.length; i++) n += _countParts_(part.parts[i]);
  return n;
}

function _b64UrlDecode_(s) {
  if (!s) return '';
  s = String(s).replace(/-/g, '+').replace(/_/g, '/');
  while (s.length % 4) s += '=';
  try {
    return Utilities.newBlob(Utilities.base64Decode(s)).getDataAsString();
  } catch (e) {
    log_('WARN', 'base64 decode failed', e);
    return '';
  }
}

// ----- Attachment byte-fetch -----

function _fetchAttachmentBytes_(messageId, attachmentId) {
  const url = GMAIL_API + '/messages/' + encodeURIComponent(messageId) +
              '/attachments/' + encodeURIComponent(attachmentId);
  const resp = _gmailFetch_(url);
  return _b64UrlDecode_(resp.data);
}

// ----- Expanders (zip / nested email) -----

function _expandZipAttachment_(att, fullMsg, cfg) {
  const out = [];
  try {
    const raw = _fetchAttachmentBytes_(fullMsg.id, att.attachmentId);
    const blob = Utilities.newBlob(raw, att.mimeType || 'application/zip', att.filename);
    const inner = Utilities.unzip(blob);
    for (let i = 0; i < inner.length; i++) {
      const f = inner[i];
      const fname = f.getName() || ('file_' + i);
      const mime = f.getContentType() || 'application/octet-stream';
      if (isExcludedMime_(mime, cfg)) continue;
      if (isZipExt_(fname, cfg)) {
        try {
          const sub = Utilities.unzip(f);
          for (let k = 0; k < sub.length; k++) {
            const subItem = _blobToDocBlob_(sub[k], 'zipAndNestedEmail');
            if (!isExcludedMime_(subItem.mimeType, cfg)) out.push(subItem);
          }
        } catch (e) { log_('WARN', 'Nested zip extract failed: ' + fname, e); }
        continue;
      }
      out.push(_blobToDocBlob_(f, 'zip'));
    }
  } catch (e) { log_('ERROR', 'Zip expansion failed: ' + att.filename, e); }
  return out;
}

function _expandNestedEmailAttachment_(att, fullMsg, cfg) {
  const out = [];
  try {
    const raw = _fetchAttachmentBytes_(fullMsg.id, att.attachmentId);
    const parsed = _parseRfc822ToJson_(raw, cfg);
    if (parsed && parsed.payload) {
      const innerAtts = _walkForAttachments_(parsed.payload, cfg);
      for (let i = 0; i < innerAtts.direct.length; i++) {
        const a = innerAtts.direct[i];
        if (isExcludedMime_(a.mimeType, cfg)) continue;
        if (isZipExt_(a.filename, cfg)) {
          const z = _expandZipAttachment_(a, parsed, cfg);
          for (let k = 0; k < z.length; k++) {
            if (!isExcludedMime_(z[k].mimeType, cfg)) out.push(z[k]);
          }
        } else if (isEmailExt_(a.filename, cfg)) {
          try {
            const r2 = _fetchAttachmentBytes_(parsed.id, a.attachmentId);
            const p2 = _parseRfc822ToJson_(r2, cfg);
            if (p2 && p2.payload) {
              const inner2 = _walkForAttachments_(p2.payload, cfg);
              for (let j = 0; j < inner2.direct.length; j++) {
                const a2 = inner2.direct[j];
                if (!isExcludedMime_(a2.mimeType, cfg) && !isEmailExt_(a2.filename, cfg) && !isZipExt_(a2.filename, cfg)) {
                  const bytes2 = _fetchAttachmentBytes_(p2.id, a2.attachmentId);
                  const blob2 = Utilities.newBlob(bytes2, a2.mimeType, a2.filename);
                  if (!isExcludedMime_(blob2.getContentType(), cfg)) out.push(_blobToDocBlob_(blob2, 'nestedEmail'));
                }
              }
            }
          } catch (e) { log_('WARN', '2-deep nested email parse failed', e); }
        } else {
          const bytes = _fetchAttachmentBytes_(parsed.id, a.attachmentId);
          const blob = Utilities.newBlob(bytes, a.mimeType, a.filename);
          if (!isExcludedMime_(blob.getContentType(), cfg)) out.push(_blobToDocBlob_(blob, 'nestedEmail'));
        }
      }
    }
  } catch (e) { log_('ERROR', 'Nested email expansion failed: ' + att.filename, e); }
  return out;
}

/** Parse a raw RFC822 byte string into a Gmail-like "message" object. */
function _parseRfc822ToJson_(rawBytes, cfg) {
  try {
    const rawStr = Utilities.newBlob(rawBytes).getDataAsString();
    const parsed = _miniRfc822Parse_(rawStr);
    if (!parsed) return null;
    const headers = [];
    Object.keys(parsed.headers).forEach(function (k) {
      headers.push({ name: k, value: parsed.headers[k] });
    });
    return {
      id: 'nested_' + sha256_(rawStr).slice(0, 16),
      threadId: 'nested',
      internalDate: parsed.dateMs || 0,
      labelIds: [],
      snippet: (parsed.bodyPlain || '').slice(0, 200),
      sizeEstimate: rawBytes.length,
      payload: parsed.payload
    };
  } catch (e) {
    log_('ERROR', '_parseRfc822ToJson_ failed', e);
    return null;
  }
}

function _miniRfc822Parse_(raw) {
  const headerEnd = _findHeaderEnd_(raw);
  if (headerEnd < 0) return null;
  const headerBlock = raw.slice(0, headerEnd);
  const body = raw.slice(headerEnd);

  const headers = _parseHeaders_(headerBlock);
  const ctype = (headers['content-type'] || 'text/plain').toLowerCase();
  const dateMs = Date.parse(headers['date'] || '') || 0;
  const payload = _parseMimeBody_(body, ctype, headers);
  return {
    headers: headers,
    dateMs: dateMs,
    payload: payload,
    bodyPlain: _flattenText_(payload)
  };
}

function _findHeaderEnd_(raw) {
  const i = raw.indexOf('\r\n\r\n');
  if (i >= 0) return i + 4;
  const j = raw.indexOf('\n\n');
  if (j >= 0) return j + 2;
  return -1;
}

function _parseHeaders_(block) {
  const out = {};
  const lines = block.split(/\r?\n/);
  let cur = '';
  for (let i = 0; i < lines.length; i++) {
    const ln = lines[i];
    if (!ln.trim()) break;
    if (/^[ \t]/.test(ln)) {
      cur += ' ' + ln.trim();
    } else {
      if (cur) {
        const idx = cur.indexOf(':');
        if (idx > 0) {
          const k = cur.slice(0, idx).trim().toLowerCase();
          const v = cur.slice(idx + 1).trim();
          if (!(k in out)) out[k] = v; else out[k] += ' ' + v;
        }
      }
      cur = ln;
    }
  }
  if (cur) {
    const idx = cur.indexOf(':');
    if (idx > 0) {
      const k = cur.slice(0, idx).trim().toLowerCase();
      const v = cur.slice(idx + 1).trim();
      if (!(k in out)) out[k] = v; else out[k] += ' ' + v;
    }
  }
  return out;
}

function _parseMimeBody_(body, contentType, parentHeaders) {
  const mainType = contentType.split(';')[0].trim().toLowerCase();
  if (mainType.indexOf('multipart/') === 0) {
    const boundary = _extractParam_(contentType, 'boundary');
    if (!boundary) return { mimeType: mainType, filename: '', body: { data: '' } };
    const parts = _splitByBoundary_(body, boundary);
    const sub = [];
    for (let i = 0; i < parts.length; i++) {
      if (!parts[i].trim()) continue;
      const subHdrEnd = _findHeaderEnd_(parts[i]);
      if (subHdrEnd < 0) continue;
      const subHdrBlock = parts[i].slice(0, subHdrEnd);
      const subBody = parts[i].slice(subHdrEnd);
      const subHeaders = _parseHeaders_(subHdrBlock);
      const subCtype = (subHeaders['content-type'] || 'text/plain').toLowerCase();
      const subDisp = (subHeaders['content-disposition'] || '').toLowerCase();
      const filename = _extractParam_(subCtype, 'name') || _extractParam_(subDisp, 'filename') || '';
      const subPart = _parseMimeBody_(subBody, subCtype, subHeaders);
      if (filename) {
        subPart.filename = filename;
        subPart.mimeType = subCtype.split(';')[0].trim();
        if (subPart.body && subPart.body.data) {
          subPart.body.attachmentId = '__parsed__' + sha256_(subPart.body.data).slice(0, 16);
        }
      }
      sub.push(subPart);
    }
    return { mimeType: mainType, parts: sub, filename: '' };
  }

  const transferEnc = (parentHeaders['content-transfer-encoding'] || '7bit').toLowerCase();
  const disposition = (parentHeaders['content-disposition'] || '').toLowerCase();
  const filename = _extractParam_(contentType, 'name') || _extractParam_(disposition, 'filename') || '';
  let data = body;
  if (transferEnc === 'base64') {
    data = data.replace(/\s+/g, '');
    try {
      data = Utilities.base64Encode(Utilities.base64Decode(data));
    } catch (e) {
      log_('WARN', 'part base64 decode failed', e);
      data = '';
    }
  } else if (transferEnc === 'quoted-printable') {
    data = _qpDecode_(data);
  }
  return {
    mimeType: mainType,
    filename: filename,
    body: { data: data, size: (data || '').length }
  };
}

function _flattenText_(payload) {
  if (!payload) return '';
  if (payload.mimeType === 'text/plain' && payload.body && payload.body.data) {
    return payload.body.data;
  }
  if (payload.parts) {
    let out = '';
    for (let i = 0; i < payload.parts.length; i++) {
      const t = _flattenText_(payload.parts[i]);
      if (t) out += t + '\n';
    }
    return out;
  }
  return '';
}

function _splitByBoundary_(body, boundary) {
  const delim = '--' + boundary;
  const lines = body.split(/\r?\n/);
  const chunks = [];
  let cur = [];
  for (let i = 0; i < lines.length; i++) {
    if (lines[i] === delim || lines[i] === delim + '--') {
      chunks.push(cur.join('\n'));
      cur = [];
    } else {
      cur.push(lines[i]);
    }
  }
  if (cur.length) chunks.push(cur.join('\n'));
  return chunks.slice(1);
}

function _extractParam_(s, name) {
  if (!s) return '';
  const re = new RegExp('(?:^|;)\\s*' + name + '\\s*=\\s*"([^"]*)"', 'i');
  const m = s.match(re);
  if (m) return m[1];
  const re2 = new RegExp('(?:^|;)\\s*' + name + '\\s*=\\s*([^;\\s]+)', 'i');
  const m2 = s.match(re2);
  return m2 ? m2[1].replace(/^"|"$/g, '') : '';
}

function _qpDecode_(s) {
  s = s.replace(/=\r?\n/g, '');
  return s.replace(/=([A-Fa-f0-9]{2})/g, function (_, h) {
    return String.fromCharCode(parseInt(h, 16));
  });
}

// ----- Helpers shared by expanders -----

function _blobToDocBlob_(blob, source) {
  return {
    filename: blob.getName() || ('file_' + Date.now()),
    mimeType: blob.getContentType() || 'application/octet-stream',
    bytes: Utilities.base64Encode(blob.getBytes()),
    body: '',
    source: source
  };
}

function _docRowFromBlob_(item, ctx, cfg) {
  let bytes = 0;
  let sha = '';
  let driveFileId = '';
  let driveFileUrl = '';
  try {
    const raw = item.bytes
      ? Utilities.newBlob(Utilities.base64Decode(item.bytes), item.mimeType, item.filename).getBytes()
      : [];
    bytes = raw.length;
    if (bytes > 0 && bytes <= 4 * 1024 * 1024) {
      sha = sha256_(Utilities.newBlob(raw).getDataAsString()).slice(0, 64);
    } else if (bytes > 4 * 1024 * 1024) {
      sha = sha256_(String(bytes) + ':' + item.filename).slice(0, 64);
    }
    if (bytes >= 1024 * 1024) {
      try {
        const blob = Utilities.newBlob(raw, item.mimeType, item.filename);
        const folder = _getOrCreateOverflowFolder_();
        const f = folder.createFile(blob);
        driveFileId = f.getId();
        driveFileUrl = f.getUrl();
      } catch (e) { log_('WARN', 'overflow upload failed for ' + item.filename, e); }
    }
  } catch (e) {
    log_('WARN', '_docRowFromBlob_ hashing failed', e);
  }

  const path = (ctx.parentZipName ? ctx.parentZipName + '!/' : '') + (item.filename || '');

  return [
    sha256_(path + '|' + (ctx.messageId || '') + '|' + bytes).slice(0, 32),
    '',
    item.source || ctx.source || 'direct',
    ctx.messageId || '',
    ctx.threadId || '',
    ctx.emailSubject || '',
    ctx.emailFrom || '',
    ctx.emailDate || '',
    ctx.emailDateMs || 0,
    ctx.month || '',
    item.filename || '',
    fileExt_(item.filename || ''),
    item.mimeType || '',
    bytes,
    sha,
    driveFileId,
    driveFileUrl,
    ctx.parentZipName || '',
    ctx.parentEmailSubject || '',
    path,
    new Date().toISOString()
  ];
}

function _getOrCreateOverflowFolder_() {
  const it = DriveApp.getFoldersByName(OVERFLOW_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(OVERFLOW_FOLDER_NAME);
}
