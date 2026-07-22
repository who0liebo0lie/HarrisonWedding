/**  MAIL MERGE — email your invite to a spreadsheet of contacts
 *  1. Google Sheet with columns:  Name | Email
 *  2. Extensions -> Apps Script -> paste this -> Save
 *  3. Edit SUBJECT / BODY / SITE_URL below
 *  4. Run "sendInvites" -> approve permissions
 *  Gmail sends ~500/day on a personal account.
 */
var SITE_URL = "https://harrisonsetsail2027.com/";
var SUBJECT  = "You're invited — Wesley & Julia are getting married at sea!";

function bodyFor(name) {
  return "Hi " + name + ",\n\n" +
    "We're getting married aboard Royal Caribbean's Harmony of the Seas, and we'd love for you to be there.\n\n" +
    "Wesley Harrison & Julia Porrino\n" +
    "Ceremony: Tuesday, January 26, 2027 — Nassau, Bahamas\n" +
    "Sailing: January 23–28, 2027 from Port Canaveral, FL\n\n" +
    "Everything you need — schedule, staterooms, attire, and RSVP — is here:\n" +
    SITE_URL + "\n\n" +
    "The site is just for our invited guests. To get in, enter your FIRST AND LAST NAME\n" +
    "exactly as it appears on this invitation.\n\n" +
    "Please RSVP and book your room soon — final payment is due at the end of October 2026.\n\n" +
    "Can't wait to celebrate with you!\n" +
    "Wesley & Julia\n" +
    "#HarrisonsSetSail";
}

function sendInvites() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var rows = sh.getDataRange().getValues();
  rows.shift(); // drop header row
  var sent = 0;
  rows.forEach(function (r) {
    var name = String(r[0]).trim();
    var email = String(r[1]).trim();
    if (!email || email.indexOf("@") < 0) return;
    MailApp.sendEmail(email, SUBJECT, bodyFor(name || "there"));
    sent++;
    Utilities.sleep(500);
  });
  SpreadsheetApp.getUi().alert("Sent " + sent + " invitations.");
}
