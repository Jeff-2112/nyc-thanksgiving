// Google Apps Script: paste into a Google Sheet via Extensions > Apps Script.
function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Status') || ss.insertSheet('Status');
  if (sh.getLastRow() == 0) sh.appendRow(['Name','Level','Stamps','Discoveries','TriviaDone','TriviaFirstTry','Updated','TriviaByArea']);
  var rows = sh.getDataRange().getValues(), r = -1;
  for (var i = 1; i < rows.length; i++) if (rows[i][0] == d.who) { r = i + 1; break; }
  var row = [d.who, d.depth, d.stamps, d.disc, d.trDone, d.trFirst, new Date(), JSON.stringify(d.trArea)];
  if (r < 0) sh.appendRow(row); else sh.getRange(r, 1, 1, row.length).setValues([row]);
  return ContentService.createTextOutput('ok');
}
function doGet() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Status'), out = [];
  if (sh) {
    var v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) out.push({who:v[i][0],depth:v[i][1],stamps:v[i][2],disc:v[i][3],trDone:v[i][4],trFirst:v[i][5],updated:v[i][6]});
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
