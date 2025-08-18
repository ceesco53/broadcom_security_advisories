function updateBroadcomCVEs() {
  const SHEET_NAME = "CVEs";
  const SEGMENTS = ["VT"];
  const PAGE_SIZE = 50;
  const API_URL = "https://support.broadcom.com/web/ecx/security-advisory/-/securityadvisory/getSecurityAdvisoryList";

  // 👇 Replace with your real spreadsheet ID
  const ss = SpreadsheetApp.openById("spreadsheet-id");

  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  const headers = ["CVE ID", "RATING", "COMMENTS", "Link", "Pub date"];

  // --- Ensure headers exist ---
  const firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = firstRow.join("") !== "";
  if (!hasHeaders) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // --- Fetch advisories ---
  let collected = [];
  SEGMENTS.forEach(segment => {
    collected = collected.concat(fetchAllRows_(API_URL, PAGE_SIZE, segment));
  });

  const rows = [];
  const now = new Date();

  for (const item of collected) {
    const updatedRaw = String(item?.updated || "").trim();
    const updatedDate = parseUpdatedToDate_(updatedRaw);
    if (!updatedDate) continue;
    if (!withinLast14Days_(updatedDate, now)) continue;

    rows.push([
      getFullNotificationId_(item),
      String(item?.severity || "").trim(),
      String(item?.title || "").trim(),
      String(item?.notificationUrl || "").trim(),
      formatPubDate_(updatedDate)
    ]);
  }

  // --- Append data ---
  if (rows.length > 0) {
    const lastRow = sh.getLastRow();
    sh.getRange(lastRow + 1, 1, rows.length, headers.length).setValues(rows);
  }

  Logger.log("Appended " + rows.length + " rows to '" + SHEET_NAME + "'.");
}

/* ----------------------- Helpers ----------------------- */

function fetchAllRows_(apiUrl, pageSize, segment) {
  let page = 0, results = [];
  while (true) {
    const data = fetchPage_(apiUrl, page, pageSize, segment);
    const batch = (data && data.list) ? data.list : [];
    results = results.concat(batch);

    const total = (data && data.pageInfo && typeof data.pageInfo.totalCount === "number")
      ? data.pageInfo.totalCount
      : results.length;

    if (results.length >= total || batch.length === 0) break;
    page += 1;
    Utilities.sleep(150);
  }
  return results;
}

function fetchPage_(apiUrl, pageNumber, pageSize, segment) {
  const payload = {
    pageNumber: pageNumber,
    pageSize: pageSize,
    searchVal: "",
    segment: segment,
    sortInfo: { column: "", order: "" }
  };

  const res = UrlFetchApp.fetch(apiUrl, {
    method: "post",
    contentType: "application/json; charset=UTF-8",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error("HTTP " + code + ": " + res.getContentText());

  const json = JSON.parse(res.getContentText());
  if (!json || json.success !== true) throw new Error("API returned success=false: " + res.getContentText());
  return json.data;
}

function getFullNotificationId_(item) {
  const candidates = [
    "notificationCode",
    "notificationNo",
    "notificationNumber",
    "notificationIdentifier",
    "notificationIdStr",
    "notification_id"
  ];
  for (var i = 0; i < candidates.length; i++) {
    var key = candidates[i];
    var val = item && item[key];
    if (typeof val === "string" && /[A-Z]{3,}-\d{4}-\d+/.test(val)) {
      return val.trim();
    }
  }
  const url = String(item?.notificationUrl || "").trim();
  const m = url.match(/\/([A-Z]{3,}-\d{4}-\d+)(?:[/?#]|$)/);
  if (m) return m[1];
  return String(item?.notificationId ?? "").trim();
}

function parseUpdatedToDate_(ts) {
  if (!ts) return null;
  var t = String(ts).trim().replace(" ", "T");
  var d = new Date(t);
  if (isNaN(d.getTime())) {
    d = new Date(t.substring(0, 10));
    if (isNaN(d.getTime())) return null;
  }
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function withinLast14Days_(updatedDate, now) {
  const MS_PER_DAY = 24 * 60 * 60 * 1000;
  const diffDays = (now.getTime() - updatedDate.getTime()) / MS_PER_DAY;
  return diffDays <= 14;
}

function formatPubDate_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "d MMMM yyyy");
}