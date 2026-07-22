/**  BINGO BACKEND — Google Apps Script
 *  Paste into: your RSVP responses Google Sheet -> Extensions -> Apps Script
 *  Then Deploy -> New deployment -> Web app -> Execute as Me -> Access: Anyone
 *  Copy the Web App URL into bingo.html (const API_URL = "...")
 *
 *  Expected columns on your responses tab: Name | Email | Fun Fact
 *  (adjust COL_* below if your columns are in a different order)
 */
var SHEET_RESPONSES = 0;   // first tab
var SHEET_GUESSES   = "Guesses";
var COL_NAME = 1, COL_EMAIL = 2, COL_FACT = 3;   // 1-based column numbers

function getData_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[SHEET_RESPONSES];
  var rows = sh.getDataRange().getValues();
  rows.shift(); // drop header
  var facts = [], names = [], key = {};
  rows.forEach(function (r) {
    var n = String(r[COL_NAME - 1]).trim();
    var f = String(r[COL_FACT - 1]).trim();
    if (n) names.push(n);
    if (f) { facts.push(f); key[f] = n; }
  });
  return { facts: facts, names: names, key: key };
}

function guessSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_GUESSES);
  if (!sh) { sh = ss.insertSheet(SHEET_GUESSES); sh.appendRow(["Player", "Fact", "Guess", "Updated"]); }
  return sh;
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "facts";
  var d = getData_();
  if (action === "facts") {
    return json_({ facts: d.facts, names: d.names });
  }
  if (action === "leaderboard") {
    var rows = guessSheet_().getDataRange().getValues(); rows.shift();
    var score = {};
    rows.forEach(function (r) {
      var player = r[0], fact = r[1], guess = String(r[2]).trim().toLowerCase();
      if (!player) return;
      if (score[player] === undefined) score[player] = 0;
      var truth = String(d.key[fact] || "").trim().toLowerCase();
      if (truth && guess && truth === guess) score[player]++;
    });
    var lb = Object.keys(score).map(function (p) { return { player: p, correct: score[p] }; })
      .sort(function (a, b) { return b.correct - a.correct; });
    return json_({ leaderboard: lb });
  }
  return json_({ ok: true });
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  if (body.action === "guess") {
    var sh = guessSheet_();
    var rows = sh.getDataRange().getValues();
    // remove this player's old guesses
    for (var i = rows.length - 1; i >= 1; i--) {
      if (rows[i][0] === body.player) sh.deleteRow(i + 1);
    }
    var now = new Date();
    Object.keys(body.guesses || {}).forEach(function (fact) {
      sh.appendRow([body.player, fact, body.guesses[fact], now]);
    });
  }
  return json_({ ok: true });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
