/**
 * Cognitronics — Investor deck access logging
 *
 * Deploy as a separate Google Apps Script Web App (not the thesis-validation script):
 * 1. Open the same Google Sheet as thesis validation, or create a new one.
 * 2. Extensions → Apps Script → New project → paste this file.
 * 3. Set SHEET_ID, SHEET_NAME, EMAIL_TO below.
 * 4. Run setupSheet() once (authorize when prompted).
 * 5. Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy the /exec URL into deck/index.html → GOOGLE_APPS_SCRIPT_URL
 *
 * Frontend sends JSON via POST (Content-Type: text/plain) to avoid CORS preflight.
 */

const SHEET_ID = "PASTE_YOUR_SHEET_ID";
const SHEET_NAME = "Deck access";
const EMAIL_TO = "cognitronics@proton.me";

function setupSheet() {
  const sheet = getSheet_();
  const headers = [
    "timestamp",
    "email",
    "session_id",
    "source",
    "user_agent",
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}

function doPost(e) {
  try {
    const raw =
      (e && e.postData && e.postData.contents) ||
      (e && e.parameter && e.parameter.payload) ||
      "";

    if (!raw) {
      return jsonResponse_(false, "Empty request body.");
    }

    const data = JSON.parse(raw);
    appendRow_(data);
    sendNotificationEmail_(data);

    return jsonResponse_(true, "Deck access recorded.");
  } catch (err) {
    console.error(err);
    return jsonResponse_(false, String(err));
  }
}

function doGet() {
  return jsonResponse_(true, "Deck access endpoint is live.");
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

function appendRow_(data) {
  const sheet = getSheet_();
  if (sheet.getLastRow() === 0) {
    setupSheet();
  }

  const row = [
    data.timestamp || new Date().toISOString(),
    data.email || "",
    data.session_id || "",
    data.source || "",
    data.user_agent || "",
  ];

  sheet.appendRow(row);
}

function sendNotificationEmail_(data) {
  const lines = [
    "New investor deck access.",
    "",
    "Email: " + (data.email || "(not provided)"),
    "Session ID: " + (data.session_id || "—"),
    "Timestamp: " + (data.timestamp || new Date().toISOString()),
    "Source: " + (data.source || "—"),
    "User agent: " + (data.user_agent || "—"),
  ];

  MailApp.sendEmail({
    to: EMAIL_TO,
    subject: "deck access",
    body: lines.join("\n"),
  });
}

function jsonResponse_(ok, message) {
  const output = ContentService.createTextOutput(
    JSON.stringify({ ok: ok, message: message })
  );
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
