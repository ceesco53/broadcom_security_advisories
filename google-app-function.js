/**
 * Broadcom Security Advisories → Google Sheet exporter (Spreadsheet ID based)
 *
 * Spreadsheet: opened via SPREADSHEET_ID
 * Tab: "TasCVE"
 *
 * Required headers in row 1 (exact text):
 *   "CVE ID", "RATING", "COMMENTS", "Link", "Pub Date", "RR Date"
 *
 * Mapping:
 *   CVE ID  <- notificationId
 *   RATING  <- severity
 *   Link    <- notificationUrl
 *   Pub Date<- published (e.g. "28 December 2025")
 *
 * Behavior:
 * - Default: last 7 days (inclusive), based on "published"
 * - Clears ONLY rows in that date window, then re-populates
 * - One row per advisory (no CVE splitting)
 * - COMMENTS and RR Date are left blank
 */

const TAS_CVE_CONFIG = {
  // >>> IMPORTANT: set this <<<
  SPREADSHEET_ID: "1TtB7bee5KbIJZzrzcSEZboViufs2SrDuEmbFs_5blTw",

  TAB_NAME: "TasCVE",

  ENDPOINT_URL:
    "https://support.broadcom.com/web/ecx/security-advisory/-/securityadvisory/getSecurityAdvisoryList",

  SEGMENT: "VT",
  PAGE_SIZE: 200,
  MAX_PAGES: 15,

  HEADERS: {
    accept: "application/json",
    "content-type": "application/json",
  },
};

/* ================= MENU =================
 * Note: onOpen() only adds a menu to the spreadsheet UI when this script is
 * bound to that spreadsheet. If this is a standalone script, the menu won't
 * appear—but the export functions still work.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("TasCVE")
    .addItem("Refresh last 7 days", "menuTasCVE_last7")
    .addItem("Refresh last 14 days", "menuTasCVE_last14")
    .addItem("Refresh last 30 days", "menuTasCVE_last30")
    .addSeparator()
    .addItem("Refresh custom date range…", "menuTasCVE_customRange")
    .addSeparator()
    .addItem("Debug: fetch first page (Logs)", "debugTasCVE_fetchFirstPage")
    .addToUi();
}

function menuTasCVE_last7() { exportTasCVE(); }
function menuTasCVE_last14() { exportTasCVE({ daysBack: 14 }); }
function menuTasCVE_last30() { exportTasCVE({ daysBack: 30 }); }

function menuTasCVE_customRange() {
  const ui = SpreadsheetApp.getUi();

  const startResp = ui.prompt(
    "TasCVE – Custom Range",
    'Enter START date (YYYY-MM-DD), e.g. "2025-12-20":',
    ui.ButtonSet.OK_CANCEL
  );
  if (startResp.getSelectedButton() !== ui.Button.OK) return;

  const endResp = ui.prompt(
    "TasCVE – Custom Range",
    'Enter END date (YYYY-MM-DD), e.g. "2025-12-31":',
    ui.ButtonSet.OK_CANCEL
  );
  if (endResp.getSelectedButton() !== ui.Button.OK) return;

  const startDate = startResp.getResponseText().trim();
  const endDate = endResp.getResponseText().trim();

  if (!/^\d{4}-\d{2}-\d{2}$/.test(startDate) || !/^\d{4}-\d{2}-\d{2}$/.test(endDate)) {
    ui.alert("Invalid format. Please use YYYY-MM-DD.");
    return;
  }

  exportTasCVE({ startDate, endDate });
  ui.alert(`TasCVE refresh complete for ${startDate} → ${endDate}.`);
}

/* ================= ENTRY POINT ================= */

function exportTasCVE(options) {
  options = options || {};
  const range = resolveDateRange_(options);

  Logger.log(`TasCVE window: ${range.startDate.toISOString()} → ${range.endDate.toISOString()} (tz=${range.tz})`);

  const { ss, sheet } = getSpreadsheetAndTab_();
  ensureHeadersExist_(sheet);

  clearTasCVERowsInRange_(sheet, range);

  const advisories = fetchAdvisoriesWindowed_(range);
  Logger.log(`Fetched advisories total (all pages): ${advisories.length}`);

  const rows = advisoriesToRows_(advisories, range);
  Logger.log(`Rows to write (in window): ${rows.length}`);

  writeRows_(sheet, rows);

  Logger.log("TasCVE export complete.");
}

/* ================= DATE RANGE ================= */

function resolveDateRange_(options) {
  const tz = Session.getScriptTimeZone();
  const now = new Date();

  let startDate, endDate;

  if (options.startDate && options.endDate) {
    startDate = new Date(options.startDate);
    endDate = new Date(options.endDate);
  } else {
    const daysBack = Number.isFinite(options.daysBack) ? options.daysBack : 7;
    endDate = now;
    startDate = new Date(now.getTime() - daysBack * 86400000);
  }

  startDate = new Date(Utilities.formatDate(startDate, tz, "yyyy/MM/dd") + " 00:00:00");
  endDate = new Date(Utilities.formatDate(endDate, tz, "yyyy/MM/dd") + " 23:59:59");

  return { startDate, endDate, tz };
}

/* ================= SHEET ACCESS ================= */

function getSpreadsheetAndTab_() {
  const ss = SpreadsheetApp.openById(TAS_CVE_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TAS_CVE_CONFIG.TAB_NAME);
  if (!sheet) throw new Error(`Tab "${TAS_CVE_CONFIG.TAB_NAME}" not found in spreadsheet ${TAS_CVE_CONFIG.SPREADSHEET_ID}.`);
  return { ss, sheet };
}

/* ================= VALIDATE HEADERS ================= */

function ensureHeadersExist_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = mapHeaders_(headers);

  const required = ["CVE ID", "RATING", "COMMENTS", "Link", "Pub Date", "RR Date"];
  const missing = required.filter(h => !col[h]);
  if (missing.length) {
    throw new Error(`TasCVE is missing required header(s): ${missing.join(", ")}. Headers must match exactly.`);
  }
}

/* ================= CLEAR EXISTING ROWS ================= */

function clearTasCVERowsInRange_(sheet, range) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = mapHeaders_(headers);
  const pubCol = col["Pub Date"];

  const values = sheet.getRange(2, pubCol, lastRow - 1, 1).getValues();

  const rows = [];
  values.forEach((v, i) => {
    const d = coerceDate_(v[0]);
    if (d && d >= range.startDate && d <= range.endDate) rows.push(i + 2);
  });

  if (!rows.length) return;

  rows.sort((a, b) => a - b);

  let start = rows[0];
  let prev = rows[0];
  const blocks = [];

  for (let i = 1; i < rows.length; i++) {
    if (rows[i] === prev + 1) {
      prev = rows[i];
    } else {
      blocks.push({ start, count: prev - start + 1 });
      start = rows[i];
      prev = rows[i];
    }
  }
  blocks.push({ start, count: prev - start + 1 });

  for (let i = blocks.length - 1; i >= 0; i--) {
    sheet.deleteRows(blocks[i].start, blocks[i].count);
  }

  Logger.log(`Cleared rows in window: ${rows.length}`);
}

/* ================= FETCH (paged POST) ================= */

function fetchAdvisoriesWindowed_(range) {
  const results = [];
  let page = 0;

  for (let i = 0; i < TAS_CVE_CONFIG.MAX_PAGES; i++) {
    Logger.log(`Fetching page ${page} (pageSize=${TAS_CVE_CONFIG.PAGE_SIZE})...`);

    const payload = {
      pageNumber: page,
      pageSize: TAS_CVE_CONFIG.PAGE_SIZE,
      searchVal: "",
      segment: TAS_CVE_CONFIG.SEGMENT,

      // Force sort so we can stop early.
      // If the API ignores this, it still works; it just may need more pages.
      sortInfo: { column: "published", order: "DESC" },
    };

    const resp = UrlFetchApp.fetch(TAS_CVE_CONFIG.ENDPOINT_URL, {
      method: "post",
      headers: TAS_CVE_CONFIG.HEADERS,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      followRedirects: true,
    });

    const code = resp.getResponseCode();
    const body = resp.getContentText();
    if (code < 200 || code >= 300) {
      throw new Error(`HTTP ${code} from Broadcom endpoint. Body: ${body.substring(0, 500)}`);
    }

    const json = JSON.parse(body);
    const list = Array.isArray(json?.data?.list) ? json.data.list : [];
    Logger.log(`Page ${page} returned ${list.length} items`);

    // Track oldest publish date on this page (for early stop)
    let oldestOnPage = null;

    for (const item of list) {
      const pub = parseBroadcomDate_(item.published);
      if (!pub) continue;

      if (!oldestOnPage || pub < oldestOnPage) oldestOnPage = pub;

      // Keep only window items
      if (pub >= range.startDate && pub <= range.endDate) {
        results.push(item);
      }
    }

    // Early stop: if sorted DESC, once the oldest date on the page is older than startDate,
    // subsequent pages will be even older.
    if (oldestOnPage && oldestOnPage < range.startDate) {
      Logger.log(`Early stop: oldestOnPage ${oldestOnPage.toISOString()} < startDate ${range.startDate.toISOString()}`);
      break;
    }

    const nextPage = json?.data?.pageInfo?.nextPage;
    if (nextPage == null) break;
    page = nextPage;
  }

  Logger.log(`Fetched windowed advisories total kept=${results.length}`);
  return results;
}

/* ================= TRANSFORM ================= */

function advisoriesToRows_(advisories, range) {
  const rows = [];
  const seen = new Set();

  advisories.forEach(a => {
    const pub = parseBroadcomDate_(a.published);
    if (!pub) return;
    if (pub < range.startDate || pub > range.endDate) return;

    if (a.notificationId == null) return;

    const advisoryId = String(a.notificationId).trim();
    const severity = (a.severity || "").trim();
    const link = (a.notificationUrl || "").trim();
    if (!advisoryId || !link) return;

    const key = `${advisoryId}|${Utilities.formatDate(pub, range.tz, "yyyy-MM-dd")}|${link}`;
    if (seen.has(key)) return;
    seen.add(key);

    rows.push({
      advisoryId,              // notificationId
      advisoryUrl: link,       // notificationUrl
      severity,
      pub,
      comments: normalizeAndTruncateTitle_(a.title, 60)
    });
  });

  rows.sort((x, y) => y.pub - x.pub || x.advisoryId.localeCompare(y.advisoryId));
  return rows;
}

function parseBroadcomDate_(s) {
  if (!s) return null;
  const m = String(s).trim().match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (!m) return null;

  const months = {
    january:0,february:1,march:2,april:3,may:4,june:5,
    july:6,august:7,september:8,october:9,november:10,december:11
  };

  const month = months[m[2].toLowerCase()];
  if (month == null) return null;

  return new Date(+m[3], month, +m[1], 12, 0, 0); // noon avoids DST edges
}

/* ================= WRITE ================= */

function writeRows_(sheet, rows) {
  if (!rows.length) {
    Logger.log("No rows to write (0 advisories found in the selected window).");
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = mapHeaders_(headers);

  const out = rows.map(r => {
    const line = new Array(headers.length).fill("");
    // CVE ID as hyperlink
    line[col["CVE ID"] - 1] =
      `=HYPERLINK("${r.advisoryUrl}", "${r.advisoryId}")`;

    line[col["RATING"] - 1] = r.severity;

    // COMMENTS = truncated title
    line[col["COMMENTS"] - 1] = r.comments;

    line[col["Link"] - 1] = r.advisoryUrl;
    line[col["Pub Date"] - 1] = r.pub;
    return line;
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);

  sheet.getRange(2, col["Pub Date"], Math.max(sheet.getLastRow() - 1, 1), 1)
    .setNumberFormat("yyyy-mm-dd");
}

/* ================= HELPERS ================= */

function mapHeaders_(headers) {
  const m = {};
  headers.forEach((h, i) => {
    if (!h) return;
    m[String(h).trim()] = i + 1;
  });
  return m;
}

function coerceDate_(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

/* ================= DEBUG ================= */

function debugTasCVE_fetchFirstPage() {
  const payload = {
    pageNumber: 0,
    pageSize: 20,
    searchVal: "",
    segment: TAS_CVE_CONFIG.SEGMENT,
    sortInfo: { column: "", order: "" },
  };

  const resp = UrlFetchApp.fetch(TAS_CVE_CONFIG.ENDPOINT_URL, {
    method: "post",
    headers: TAS_CVE_CONFIG.HEADERS,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    followRedirects: true,
  });

  const body = resp.getContentText();
  const json = JSON.parse(body);

  Logger.log(`HTTP ${resp.getResponseCode()}`);
  Logger.log(`success=${json?.success === true}, listLen=${json?.data?.list?.length || 0}`);

  if (json?.data?.list?.length) {
    Logger.log("First item: " + JSON.stringify(json.data.list[0], null, 2));
  }
}

function tasCVE_debugFetchFirstPage() {
  const payload = {
    pageNumber: 0,
    pageSize: 20,
    searchVal: "",
    segment: TAS_CVE_CONFIG.SEGMENT, // "VT"
    sortInfo: { column: "", order: "" },
  };

  const resp = UrlFetchApp.fetch(TAS_CVE_CONFIG.ENDPOINT_URL, {
    method: "post",
    headers: TAS_CVE_CONFIG.HEADERS,
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText();
  console.log(`HTTP ${code}`);
  if (code < 200 || code >= 300) throw new Error(text);

  const json = JSON.parse(text);
  console.log(`success=${json.success}, listLen=${json?.data?.list?.length}`);

  const first = json?.data?.list?.[0];
  console.log("First item:", JSON.stringify(first, null, 2));

  return json?.data?.list?.length || 0;
}

function normalizeAndTruncateTitle_(title, maxLen) {
  if (!title) return "";

  let s = String(title).trim();

  // Remove leading boilerplate if present
  const PREFIX = "Product Release Advisory - ";
  if (s.startsWith(PREFIX)) {
    s = s.substring(PREFIX.length).trim();
  }

  if (s.length <= maxLen) return s;
  return s.substring(0, maxLen) + "...";
}
