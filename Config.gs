/**
 * Config.gs
 * =====================================================================
 * All configurable settings. Hard-coded defaults here; the optional
 * "Settings" sheet in the auto-created spreadsheet overrides them at
 * runtime. Overrides are loaded by `loadConfigOverridesFromSheet_()`
 * which is called from the ingestion loop and the API endpoints
 * (NOT from `getOrCreateSpreadsheet_` — see comment below).
 *
 * !!  Architecture note  !!
 * The earlier version of this file had a fatal recursion:
 *   loadConfigFromSheet_ → getOrCreateSpreadsheet_ → ensureSettingsSheet_ → loadConfigFromSheet_ → ...
 * which blew the stack after a few hundred calls and produced the
 * 10-minute "Invalid regular expression" hang. The new design is
 *   - `getOrCreateSpreadsheet_` uses HARDCODED `CFG` only (no Settings read)
 *   - `loadConfigOverridesFromSheet_` is a separate function called
 *     once at the start of each ingestion run / API call, and it
 *     ALSO uses a hardcoded guard so a missing/empty sheet just
 *     returns `CFG` without recursion.
 * =====================================================================
 */

const CFG = {
  // ----- Search / ingestion -----
  TARGET_ADDRESSES: [
    'ap-invoices@coreweave.com',
    'apinvoices@coreweave.com'
  ],
  START_DATE: '2026/02/01',
  PAGE_SIZE: 100,
  MAX_MESSAGES_PER_RUN: 400,

  // ----- Time-budget / resumable ingestion -----
  TIME_BUDGET_SECONDS: 300, // 5 min — leave 1 min buffer
  RESCHEDULE_MINUTES: 1,

  // ----- Sheets -----
  SPREADSHEET_NAME: 'Received Emails Dashboard — Database',
  EMAILS_SHEET: 'Emails',
  DOCUMENTS_SHEET: 'Documents',
  SETTINGS_SHEET: 'Settings',
  FOLDER_NAME: '',

  // ----- Filtering rules -----
  EXCLUDED_MIME_PREFIXES: ['image/', 'video/', 'audio/'],
  ZIP_EXTENSIONS: ['zip', 'rar', '7z', 'tar', 'gz', 'tgz', 'bz2'],
  EMAIL_EXTENSIONS: ['eml', 'msg'],

  // ----- Char-limit fallbacks -----
  SHEET_CELL_CHAR_LIMIT: 49000,
  OVERFLOW_BLOB_THRESHOLD: 200000,

  // ----- Trigger orchestration -----
  RESUME_HANDLER: 'resumeIngestion',

  // ----- Logging -----
  VERBOSE_LOGGING: false,

  // ----- Search query builder -----
  buildGmailQuery: function () {
    const addrs = this.TARGET_ADDRESSES.map(a => '"' + a + '"').join(' OR ');
    const parts = this.START_DATE.split('/');
    const d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    d.setDate(d.getDate() - 1);
    const cutoff =
      d.getFullYear() + '/' +
      String(d.getMonth() + 1).padStart(2, '0') + '/' +
      String(d.getDate()).padStart(2, '0');
    return '(' + addrs + ') after:' + cutoff;
  }
};

/**
 * Read the Settings sheet and return a copy of CFG with any overrides
 * applied. Returns CFG unchanged on ANY error (missing sheet, blank
 * sheet, parse error). NEVER calls back into getOrCreateSpreadsheet_.
 */
function loadConfigOverridesFromSheet_() {
  // Hard guard: no recursive calls to anything that creates the sheet.
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SPREADSHEET_ID');
  if (!ssId) return CFG;
  let ss;
  try { ss = SpreadsheetApp.openById(ssId); } catch (e) { return CFG; }
  let sheet;
  try { sheet = ss.getSheetByName(CFG.SETTINGS_SHEET); } catch (e) { return CFG; }
  if (!sheet) return CFG;
  let data;
  try { data = sheet.getDataRange().getValues(); } catch (e) { return CFG; }
  if (!data || data.length < 2) return CFG;

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const k = (data[i][0] == null ? '' : String(data[i][0])).trim();
    const v = (data[i][1] == null ? '' : String(data[i][1])).trim();
    if (k) map[k] = v;
  }

  const out = {};
  // Copy all primitive props from CFG; skip functions and modules.
  for (const k in CFG) {
    if (typeof CFG[k] !== 'function') out[k] = CFG[k];
  }
  out.buildGmailQuery = CFG.buildGmailQuery;

  if (map.START_DATE) out.START_DATE = map.START_DATE;
  if (map.TARGET_ADDRESSES) {
    out.TARGET_ADDRESSES = map.TARGET_ADDRESSES
      .split(/[\n,;]+/)
      .map(s => s.trim())
      .filter(Boolean);
  }
  if (map.PAGE_SIZE && !isNaN(Number(map.PAGE_SIZE))) out.PAGE_SIZE = Number(map.PAGE_SIZE);
  if (map.MAX_MESSAGES_PER_RUN && !isNaN(Number(map.MAX_MESSAGES_PER_RUN))) {
    out.MAX_MESSAGES_PER_RUN = Number(map.MAX_MESSAGES_PER_RUN);
  }
  if (map.TIME_BUDGET_SECONDS && !isNaN(Number(map.TIME_BUDGET_SECONDS))) {
    out.TIME_BUDGET_SECONDS = Number(map.TIME_BUDGET_SECONDS);
  }
  if (map.RESCHEDULE_MINUTES && !isNaN(Number(map.RESCHEDULE_MINUTES))) {
    out.RESCHEDULE_MINUTES = Number(map.RESCHEDULE_MINUTES);
  }
  if (map.SPREADSHEET_NAME) out.SPREADSHEET_NAME = map.SPREADSHEET_NAME;
  if (typeof map.VERBOSE_LOGGING === 'string') {
    out.VERBOSE_LOGGING = map.VERBOSE_LOGGING === 'true';
  }
  // Rebind query helper so it sees the overrides.
  out.buildGmailQuery = CFG.buildGmailQuery.bind(out);
  return out;
}
