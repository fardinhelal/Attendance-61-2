// Google Apps Script: shared leaderboard + attendance store for Business Quest
// Sheet row 1 headers: Date | Game | Name | ID | Ms | Mistakes | Avatar
const PW = 'OnlyBossIsAllowedToDoIt';
const sh = () => SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
function doGet() {
  const v = sh().getDataRange().getValues().slice(1).filter(r => r[1]);
  const out = v.map(r => ({date: String(r[0]), game: r[1], name: r[2], sid: String(r[3]), ms: r[4], miss: r[5], av: r[6]}));
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
  const d = JSON.parse(e.postData.contents), s = sh();
  if (d.action === 'reset') {
    if (d.pw === PW && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 7).clearContent();
  } else {
    s.appendRow([d.date, d.game, d.name, "'" + d.sid, d.ms, d.miss, d.av]);
  }
  return ContentService.createTextOutput('ok');
}
