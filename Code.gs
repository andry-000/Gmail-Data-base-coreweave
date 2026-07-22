/**
 * Code.gs
 * =====================================================================
 * The web app entry point + a few menu helpers for the script editor.
 * Run `install()` once after deploy to register the menu and (optionally)
 * schedule a recurring ingestion trigger.
 * =====================================================================
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('AP Invoices Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AP Invoices Dashboard')
    .addItem('Start ingestion', 'startIngestion')
    .addItem('Resume ingestion', 'resumeIngestion')
    .addItem('Reset cursor', 'resetIngestion')
    .addItem('Open dashboard', 'openDashboard')
    .addToUi();
}

function openDashboard() {
  const url = ScriptApp.getService().getUrl();
  const html = HtmlService.createHtmlOutput(
    '<script>window.open("' + url + '");google.script.host.close();</script>'
  ).setWidth(300).setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening dashboard…');
}

/** One-time setup: create spreadsheet, ensure sheets. */
function install() {
  getOrCreateSpreadsheet_();
  log_('INFO', 'install() complete');
  return { ok: true };
}
