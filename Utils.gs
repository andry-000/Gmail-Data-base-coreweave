/**
 * Utils.gs
 * =====================================================================
 * Generic helpers: SAFE logging, string handling, hashing, MIME
 * classification, time-budget tracking.
 *
 * !!  Logging note  !!
 * Apps Script's V8 `console.log` does printf-style format expansion
 * (%s / %d / %j). Any stray `%` followed by an invalid specifier
 * (e.g. in an error message, a header value, a regex pattern) makes
 * the engine throw "Invalid regular expression" / "Stack overflow",
 * which then masks the real bug. To avoid that, `log_` here escapes
 * every `%` to `%%` before writing, and uses `Logger.log` as a
 * secondary sink that does NOT do format expansion.
 * =====================================================================
 */

/** Escape a string for safe printing via console.log/console.error.
 *  Replaces every `%` with `%%` so V8's util.format doesn't interpret
 *  it as the start of a format specifier. */
function _pctEscape_(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/%/g, '%%');
}

/** Centralized logger. Always writes to Logger + console. All messages
 *  are pct-escaped so a stray `%` in a Gmail header or error message
 *  cannot crash the engine. */
function log_(level, msg, extra) {
  const parts = ['[' + level + '] ' + (msg == null ? '' : msg)];
  if (extra !== undefined) {
    try { parts.push(' | ' + safeStringify_(extra)); }
    catch (e) { parts.push(' | [unstringifiable: ' + (e && e.message || e) + ']'); }
  }
  const line = parts.join('');
  // Logger.log is safe — it does NOT do printf expansion.
  try { Logger.log(line); } catch (e) { /* ignore quota */ }
  // For Stackdriver / Executions UI, use console with escaped % chars.
  const safe = _pctEscape_(line);
  try {
    if (level === 'ERROR') console.error(safe);
    else console.log(safe);
  } catch (e) { /* last-resort: swallow if even escaped output fails */ }
}

function safeStringify_(v) {
  try {
    if (v === null || v === undefined) return String(v);
    if (typeof v === 'string') return v;
    if (v instanceof Error) {
      const st = v.stack || '';
      // cap stack to keep logs readable
      const shortStack = st.length > 4000 ? st.slice(0, 4000) + '\n...[truncated]' : st;
      return v.name + ': ' + v.message + '\n' + shortStack;
    }
    if (v instanceof Date) return v.toISOString();
    if (v instanceof Blob) return '[Blob ' + v.getBytes().length + 'B ' + v.getContentType() + ']';
    // Range / Sheet / anything else — coerce via JSON with a safe replacer
    return JSON.stringify(v, function (k, val) {
      if (val === undefined) return '[undefined]';
      if (typeof val === 'function') return '[function]';
      if (val instanceof Date) return val.toISOString();
      if (val instanceof Blob) return '[Blob]';
      // Avoid serializing huge Spreadsheet objects
      if (typeof val === 'object' && val !== null) {
        const ctor = val.constructor && val.constructor.name;
        if (ctor === 'Sheet' || ctor === 'Spreadsheet' || ctor === 'Range') return '[' + ctor + ']';
      }
      return val;
    });
  } catch (e) {
    try { return '[safeStringify failed: ' + e.message + ']'; }
    catch (e2) { return '[unstringifiable]'; }
  }
}

/** SHA-256 hash, returns hex string. */
function sha256_(s) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    s,
    Utilities.Charset.UTF_8
  );
  return raw.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

/** Truncate to N chars and append a marker so the cell is never silently clipped. */
function truncateForSheet_(s, limit) {
  if (s === null || s === undefined) return '';
  s = String(s);
  if (s.length <= limit) return s;
  const marker = '\n...[TRUNCATED ' + (s.length - limit) + ' chars]';
  return s.slice(0, Math.max(0, limit - marker.length)) + marker;
}

/** Returns true if the mime is in the excluded bucket (image / video / audio). */
function isExcludedMime_(mime, cfg) {
  if (!mime) return false;
  const m = String(mime).toLowerCase();
  for (let i = 0; i < cfg.EXCLUDED_MIME_PREFIXES.length; i++) {
    if (m.indexOf(cfg.EXCLUDED_MIME_PREFIXES[i]) === 0) return true;
  }
  return false;
}

/** Returns lowercase extension without the dot, or '' if none. */
function fileExt_(name) {
  if (!name) return '';
  const i = String(name).lastIndexOf('.');
  if (i < 0 || i === name.length - 1) return '';
  return String(name).slice(i + 1).toLowerCase();
}

/** True if filename has a zip-like extension. */
function isZipExt_(name, cfg) {
  const e = fileExt_(name);
  return cfg.ZIP_EXTENSIONS.indexOf(e) >= 0;
}

/** True if filename is an email message we should recurse into. */
function isEmailExt_(name, cfg) {
  const e = fileExt_(name);
  return cfg.EMAIL_EXTENSIONS.indexOf(e) >= 0;
}

/** Wraps a Date.now() check so each run stops early if we're burning the budget. */
function timeBudgetExceeded_(startMs, cfg) {
  return (Date.now() - startMs) >= cfg.TIME_BUDGET_SECONDS * 1000;
}

/** Format an ISO date / Gmail internal date as a Sheet-friendly string. */
function fmtDate_(ms) {
  if (!ms) return '';
  const d = new Date(Number(ms));
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone() || 'America/Chicago',
    'yyyy-MM-dd HH:mm:ss'
  );
}

/** Format YYYY-MM for grouping by month. */
function fmtMonth_(ms) {
  if (!ms) return '';
  const d = new Date(Number(ms));
  if (isNaN(d.getTime())) return '';
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}
