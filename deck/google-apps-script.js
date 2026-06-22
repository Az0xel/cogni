const SHEET_ID = "1NgvC4bT03PfvqS6qjVoX_m-ltNZ94tOHkw31_YNJRbU";
const SHEET_NAME = "Deck access";
const EMAIL_TO = "cognitronics@proton.me";
const SLIDE_COUNT = 15;

function setupSheet() {
  var sheet = getSheet_();
  var headers = buildHeaders_();

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}

function buildHeaders_() {
  var headers = [
    "session_id",
    "email",
    "session_start",
    "session_end",
    "last_sync_at",
    "session_status",
    "total_duration_sec",
    "slides_revisited",
    "link_clicks_count",
  ];

  for (var i = 1; i <= SLIDE_COUNT; i++) {
    headers.push("slide_" + i + "_sec");
    headers.push("slide_" + i + "_visits");
    headers.push("slide_" + i + "_revisited");
  }

  headers.push("link_clicks_json");
  headers.push("event_log_json");
  headers.push("source");
  headers.push("user_agent");

  return headers;
}

function doPost(e) {
  try {
    var raw =
      (e && e.postData && e.postData.contents) ||
      (e && e.parameter && e.parameter.payload) ||
      "";

    if (!raw) {
      return jsonResponse_(false, "Empty request body.");
    }

    var data = JSON.parse(raw);
    var type = data.type || "";

    if (
      type !== "session_start" &&
      type !== "session_sync" &&
      type !== "session_complete"
    ) {
      return jsonResponse_(false, "Unsupported payload type.");
    }

    if (!isValidEmail_(data.email)) {
      return jsonResponse_(false, "Invalid email address.");
    }

    if (!data.session_id) {
      return jsonResponse_(false, "Missing session_id.");
    }

    var wasComplete = upsertSessionRow_(data);

    if (type === "session_start") {
      try {
        sendStartNotificationEmail_(data);
      } catch (mailErr) {
        console.error("Start email failed:", mailErr);
      }
    }

    if (type === "session_complete" && !wasComplete) {
      try {
        sendCompleteNotificationEmail_(data);
      } catch (mailErr) {
        console.error("Complete email failed:", mailErr);
      }
    }

    return jsonResponse_(true, "Session " + type.replace("session_", "") + ".");
  } catch (err) {
    console.error(err);
    return jsonResponse_(false, String(err));
  }
}

function doGet() {
  return jsonResponse_(true, "Deck session endpoint is live.");
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

function isValidEmail_(email) {
  if (!email || typeof email !== "string") return false;
  email = email.trim();
  if (email.length > 254) return false;
  return /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+$/.test(
    email
  );
}

function findSessionRowIndex_(sessionId) {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  var ids = sheet.getRange("A2:A" + lastRow).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(sessionId)) {
      return i + 2;
    }
  }
  return -1;
}

function getSlideMetrics_(data, slideNum) {
  var slides = data.slides || {};
  var slide = slides[String(slideNum)] || slides[slideNum] || {};
  return {
    timeSec: Number(slide.time_sec) || 0,
    visits: Number(slide.visits) || 0,
    revisited: slide.revisited === true || Number(slide.visits) > 1,
  };
}

function buildSlidesRevisited_(data) {
  var slidesRevisited = data.slides_revisited || [];
  if (!slidesRevisited.length && data.slides) {
    slidesRevisited = Object.keys(data.slides).filter(function (key) {
      var slide = data.slides[key];
      return slide && (slide.revisited === true || Number(slide.visits) > 1);
    });
  }
  return slidesRevisited;
}

function buildRowFromData_(data) {
  var slidesRevisited = buildSlidesRevisited_(data);
  var linkClicks = data.link_clicks || [];
  var row = [
    data.session_id || "",
    String(data.email || "").trim(),
    data.session_start || "",
    data.session_end || "",
    data.last_sync_at || new Date().toISOString(),
    data.session_status || "active",
    Number(data.total_duration_sec) || 0,
    slidesRevisited.join(", "),
    linkClicks.length,
  ];

  for (var i = 1; i <= SLIDE_COUNT; i++) {
    var metrics = getSlideMetrics_(data, i);
    row.push(metrics.timeSec);
    row.push(metrics.visits);
    row.push(metrics.revisited ? "Y" : "N");
  }

  row.push(JSON.stringify(linkClicks));
  row.push(JSON.stringify(data.event_log || []));
  row.push(data.source || "");
  row.push(data.user_agent || "");

  return row;
}

function upsertSessionRow_(data) {
  var sheet = getSheet_();
  if (sheet.getLastRow() === 0) {
    setupSheet();
  }

  var row = buildRowFromData_(data);
  var rowIndex = findSessionRowIndex_(data.session_id);
  var wasComplete = false;

  if (rowIndex > 0) {
    wasComplete = sheet.getRange(rowIndex, 6).getValue() === "complete";
    sheet.getRange(rowIndex, 1, rowIndex, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  return wasComplete;
}

function sendStartNotificationEmail_(data) {
  var lines = [
    "New investor deck session started.",
    "",
    "Session ID: " + (data.session_id || "—"),
    "Email: " + (data.email || "—"),
    "Started: " + (data.session_start || data.last_sync_at || "—"),
    "Source: " + (data.source || "—"),
    "",
    "Live row updates in Google Sheet tab \"" + SHEET_NAME + "\".",
    "A summary email follows when the session ends.",
  ];

  MailApp.sendEmail({
    to: EMAIL_TO,
    subject: "deck access · started · " + (data.session_id || "session"),
    body: lines.join("\n"),
  });
}

function sendCompleteNotificationEmail_(data) {
  var slidesRevisited = buildSlidesRevisited_(data);
  var linkClicks = data.link_clicks || [];
  var lines = [
    "Investor deck session completed.",
    "",
    "Session ID: " + (data.session_id || "—"),
    "Email: " + (data.email || "—"),
    "Started: " + (data.session_start || "—"),
    "Ended: " + (data.session_end || "—"),
    "Total duration (sec): " + (Number(data.total_duration_sec) || 0),
    "Slides revisited: " + (slidesRevisited.length ? slidesRevisited.join(", ") : "none"),
    "Link clicks: " + linkClicks.length,
    "",
    "See full row in Google Sheet tab \"" + SHEET_NAME + "\".",
  ];

  if (linkClicks.length) {
    lines.push("", "Link clicks:");
    linkClicks.forEach(function (click, index) {
      lines.push(
        (index + 1) +
          ". slide " +
          (click.slide || "?") +
          " · " +
          (click.label || "link") +
          " · @" +
          (click.slide_sec != null ? click.slide_sec + "s" : "?") +
          " on slide"
      );
    });
  }

  MailApp.sendEmail({
    to: EMAIL_TO,
    subject: "deck access · " + (data.session_id || "session"),
    body: lines.join("\n"),
  });
}

function jsonResponse_(ok, message) {
  var output = ContentService.createTextOutput(
    JSON.stringify({ ok: ok, message: message })
  );
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
