const SHEET_ID = "1NgvC4bT03PfvqS6qjVoX_m-ltNZ94tOHkw31_YNJRbU";
const SHEET_NAME = "Data room access";
const EMAIL_TO = "cognitronics@proton.me";

const FOLDER_IDS = ["root", "business", "corporate", "fundraise", "financials"];

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
    "folders_revisited",
    "items_opened_count",
    "items_revisited",
    "current_view",
  ];

  FOLDER_IDS.forEach(function (folderId) {
    headers.push("folder_" + folderId + "_sec");
    headers.push("folder_" + folderId + "_visits");
    headers.push("folder_" + folderId + "_revisited");
  });

  headers.push("folders_json");
  headers.push("items_json");
  headers.push("item_opens_json");
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
  return jsonResponse_(true, "Data room session endpoint is live.");
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

function getFolderMetrics_(data, folderId) {
  var folders = data.folders || {};
  var folder = folders[folderId] || {};
  return {
    timeSec: Number(folder.time_sec) || 0,
    visits: Number(folder.visits) || 0,
    revisited: folder.revisited === true || Number(folder.visits) > 1,
  };
}

function buildFoldersRevisited_(data) {
  var foldersRevisited = data.folders_revisited || [];
  if (!foldersRevisited.length && data.folders) {
    foldersRevisited = Object.keys(data.folders).filter(function (key) {
      if (key === "root") return false;
      var folder = data.folders[key];
      return folder && (folder.revisited === true || Number(folder.visits) > 1);
    });
  }
  return foldersRevisited;
}

function buildItemsRevisited_(data) {
  var itemsRevisited = data.items_revisited || [];
  if (!itemsRevisited.length && data.items) {
    itemsRevisited = Object.keys(data.items).filter(function (key) {
      var item = data.items[key];
      return item && (item.revisited === true || Number(item.visits) > 1);
    });
  }
  return itemsRevisited;
}

function buildRowFromData_(data) {
  var foldersRevisited = buildFoldersRevisited_(data);
  var itemsRevisited = buildItemsRevisited_(data);
  var itemOpens = data.item_opens || [];

  var row = [
    data.session_id || "",
    String(data.email || "").trim(),
    data.session_start || "",
    data.session_end || "",
    data.last_sync_at || new Date().toISOString(),
    data.session_status || "active",
    Number(data.total_duration_sec) || 0,
    foldersRevisited.join(", "),
    itemOpens.length,
    itemsRevisited.join(", "),
    data.current_view || "",
  ];

  FOLDER_IDS.forEach(function (folderId) {
    var metrics = getFolderMetrics_(data, folderId);
    row.push(metrics.timeSec);
    row.push(metrics.visits);
    row.push(metrics.revisited ? "Y" : "N");
  });

  row.push(JSON.stringify(data.folders || {}));
  row.push(JSON.stringify(data.items || {}));
  row.push(JSON.stringify(itemOpens));
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
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  return wasComplete;
}

function sendStartNotificationEmail_(data) {
  var lines = [
    "New data room session started.",
    "",
    "Session ID: " + (data.session_id || "—"),
    "Email: " + (data.email || "—"),
    "Started: " + (data.session_start || data.last_sync_at || "—"),
    "Source: " + (data.source || "—"),
    "",
    'Live row updates in Google Sheet tab "' + SHEET_NAME + '".',
    "A summary email follows when the session ends.",
  ];

  MailApp.sendEmail({
    to: EMAIL_TO,
    subject: "data room access · started · " + (data.session_id || "session"),
    body: lines.join("\n"),
  });
}

function sendCompleteNotificationEmail_(data) {
  var foldersRevisited = buildFoldersRevisited_(data);
  var itemsRevisited = buildItemsRevisited_(data);
  var itemOpens = data.item_opens || [];
  var lines = [
    "Data room session completed.",
    "",
    "Session ID: " + (data.session_id || "—"),
    "Email: " + (data.email || "—"),
    "Started: " + (data.session_start || "—"),
    "Ended: " + (data.session_end || "—"),
    "Total duration (sec): " + (Number(data.total_duration_sec) || 0),
    "Folders revisited: " + (foldersRevisited.length ? foldersRevisited.join(", ") : "none"),
    "Items opened: " + itemOpens.length,
    "Items revisited: " + (itemsRevisited.length ? itemsRevisited.join(", ") : "none"),
    "",
    'See full row in Google Sheet tab "' + SHEET_NAME + '".',
  ];

  if (itemOpens.length) {
    lines.push("", "Item / link opens:");
    itemOpens.forEach(function (open, index) {
      lines.push(
        (index + 1) +
          ". " +
          (open.folder || "?") +
          " · " +
          (open.name || open.item || "item") +
          " · " +
          (open.kind || "?") +
          (open.revisit ? " · revisit" : "")
      );
    });
  }

  MailApp.sendEmail({
    to: EMAIL_TO,
    subject: "data room access · " + (data.session_id || "session"),
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
