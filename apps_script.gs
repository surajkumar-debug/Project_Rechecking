// Control Google Sheet me: Extensions > Apps Script > ye poora code paste karo
// TOKEN ko apna koi lamba random text bana do (portal me wahi daalna hai)
const TOKEN = 'CHANGE_ME_LONG_RANDOM_TEXT';

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.token !== TOKEN) return out({ error: 'bad token' });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (p.action === 'tabs') return out({ tabs: ss.getSheets().map(s => s.getName()) });
  if (p.action === 'tab') {
    const sh = ss.getSheetByName(p.name);
    if (!sh) return out({ error: 'tab not found: ' + p.name });
    if (sh.getLastRow() === 0) return out({ values: [] });
    return out({ values: sh.getDataRange().getValues() });
  }
  return out({ error: 'unknown action' });
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
