/**
 * Cognitronics — Thesis Validation backend
 *
 * Deploy as a Google Apps Script Web App:
 * 1. Create a Google Sheet; copy its ID into SHEET_ID below.
 * 2. Paste this file into the Apps Script editor (Extensions → Apps Script).
 * 3. Run setupSheet() once (authorize when prompted).
 * 4. Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web App URL into thesis-validation.html → GOOGLE_APPS_SCRIPT_URL
 *
 * Frontend sends JSON via POST (Content-Type: text/plain) to avoid CORS preflight.
 */

const SHEET_ID = "1NgvC4bT03PfvqS6qjVoX_m-ltNZ94tOHkw31_YNJRbU";
const SHEET_NAME = "Cognitronics survey";
const EMAIL_TO = "cognitronics@proton.me";

/** One-time header row setup */
function setupSheet() {
  const sheet = getSheet_();
  const headers = [
    "timestamp",
    "source",
    "missing_training_signals",
    "embodiment_specific_bottleneck",
    "semantic_traces_improve_learning",
    "capture_time_better_than_posthoc",
    "semantic_simulation_value",
    "worth_solving_now",
    "product_choice",
    "product_other_text",
    "strongest_reason",
    "comments",
    "name",
    "company",
    "email",
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

    return jsonResponse_(true, "Feedback recorded.");
  } catch (err) {
    console.error(err);
    return jsonResponse_(false, String(err));
  }
}

/** Optional health check in browser */
function doGet() {
  return jsonResponse_(true, "Thesis validation endpoint is live.");
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

function formatProductChoices_(product) {
  if (!product) return "";
  if (Array.isArray(product.choices) && product.choices.length) {
    return product.choices.join(", ");
  }
  return product.choice || "";
}

function appendRow_(data) {
  const sheet = getSheet_();
  if (sheet.getLastRow() === 0) {
    setupSheet();
  }

  const respondent = data.respondent || {};
  const statements = data.statements || {};
  const product = data.product_signal || {};

  const row = [
    data.timestamp || new Date().toISOString(),
    data.source || "",
    statements.missing_training_signals || "",
    statements.embodiment_specific_bottleneck || "",
    statements.semantic_traces_improve_learning || "",
    statements.capture_time_better_than_posthoc || "",
    statements.semantic_simulation_value || "",
    statements.worth_solving_now || "",
    formatProductChoices_(product),
    product.other_text || "",
    data.strongest_reason || "",
    data.comments || "",
    respondent.name || "",
    respondent.company || "",
    respondent.email || "",
  ];

  sheet.appendRow(row);
}

function sendNotificationEmail_(data) {
  const respondent = data.respondent || {};
  const statements = data.statements || {};
  const product = data.product_signal || {};

  const lines = [
    "New thesis-validation feedback received.",
    "",
    "Validation Statements",
    "-------------------",
    "1. Missing training signals: " + (statements.missing_training_signals || "—"),
    "2. Embodiment-specific bottleneck: " + (statements.embodiment_specific_bottleneck || "—"),
    "3. Semantic traces improve learning: " + (statements.semantic_traces_improve_learning || "—"),
    "4. Capture-time better than post-hoc: " + (statements.capture_time_better_than_posthoc || "—"),
    "5. Semantic simulation value: " + (statements.semantic_simulation_value || "—"),
    "6. Worth solving now: " + (statements.worth_solving_now || "—"),
    "",
    "Product Signal",
    "--------------",
    "Choice: " + (formatProductChoices_(product) || "—"),
    product.other_text ? "Other: " + product.other_text : "",
    "",
    "Critical Feedback",
    "-----------------",
    data.strongest_reason || "(none)",
    "",
    "Additional Comments",
    "-------------------",
    data.comments || "(none)",
    "",
    "Respondent",
    "--------",
    "Name: " + (respondent.name || "(not provided)"),
    "Company: " + (respondent.company || "(not provided)"),
    "Email: " + (respondent.email || "(not provided)"),
    "",
    "Timestamp: " + (data.timestamp || new Date().toISOString()),
    "Source: " + (data.source || ""),
  ].filter(Boolean);

  MailApp.sendEmail({
    to: EMAIL_TO,
    subject: "New Thesis-Validation Feedback",
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
