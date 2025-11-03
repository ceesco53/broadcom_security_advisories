/**
 * Google Docs: Paste Broadcom Tanzu advisories via cURL CSV (last 30 days via GET fromDate/toDate)
 * - Adds "Broadcom CVEs → Paste CSV from cURL" to the Doc menu
 * - Dialog shows a prebuilt curl using GET:
 *     https://www.broadcom.com/support/security/advisories/json?segment=VT&fromDate=YYYY-MM-DD&toDate=YYYY-MM-DD&pageSize=500
 * - You paste CSV output; script overwrites the current Doc with a table:
 *   Notification Id | Release Date | Products | Level | Severity
 */

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Broadcom CVEs')
    .addItem('Paste CSV from cURL', 'showPasteDialog')
    .addToUi();
}

function showPasteDialog() {
  var curl = buildCurlCommand_(); // last-30-days GET with fromDate/toDate
  var tmpl = HtmlService.createTemplateFromFile('paste_dialog');
  tmpl.curl = curl;
  var html = tmpl.evaluate().setWidth(700).setHeight(560);
  DocumentApp.getUi().showModalDialog(html, 'Broadcom CVEs (segment=VT)');
}

/**
 * Build a ready-to-run curl + jq command that calls the documented GET endpoint
 * with ISO8601 fromDate/toDate for the last 30 days, then outputs CSV.
 * Doc: https://knowledge.broadcom.com/external/article/408302/json-api-for-product-security-advisories.html
 */
function buildCurlCommand_() {
  var tz = 'UTC';
  var now = new Date();
  var from = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);

  // API expects date-only ISO8601 (YYYY-MM-DD)
  var FROM = Utilities.formatDate(from, tz, 'yyyy-MM-dd');
  var TO   = Utilities.formatDate(now,  tz, 'yyyy-MM-dd');

  var endpoint = 'https://www.broadcom.com/support/security/advisories/json';
  var qs = `?segment=VT&fromDate=${FROM}&toDate=${TO}&pageSize=500`;

  // CSV columns: Notification Id, Release Date, Products, Level, Severity
  // Products joined with '; ' to avoid commas inside CSV fields.
  // Handle both array root and {items: [...]} shapes.
  var cmd =
`# === Broadcom Tanzu advisories (segment=VT) – last 30 days ===
# Uses GET with fromDate/toDate per Broadcom KB.
# 1) Copy the curl below, run it (requires jq). 2) Paste the CSV output here. 3) Click Insert.

curl -s '${endpoint}${qs}' \\
  -H 'accept: application/json' --compressed \\
| jq -r '
  def arr(x): (x // []) | if type=="array" then join("; ") else tostring end;
  (["Notification Id","Release Date","Products","Level","Severity"]),
  (
    (if type=="array" then . else .items end)
    | .[]
    | [
        (.advisoryId // .id),
        (.issueDate  // .published // .date),
        (arr(.impactedProducts // .products)),
        (.advisorySeverity // .severity // ""),
        (.cvssMaxSeverity  // .cvssV3Range // "")
      ]
  )
  | @csv'`;

  return cmd;
}

/**
 * Server-side entry from the dialog: parse pasted CSV and overwrite the current Doc.
 * Expects header row to match the five requested columns (adds it if missing).
 */
function insertCsvIntoDoc(csvText) {
  if (!csvText || !csvText.trim()) {
    throw new Error('No CSV content received.');
  }

  var rows = Utilities.parseCsv(csvText);
  if (!rows || !rows.length) throw new Error('CSV parsing produced no rows.');

  var expectedHeader = ['Notification Id','Release Date','Products','Level','Severity'];
  var header = rows[0].map(function (s) { return String(s || '').trim(); });
  var headerMatches = expectedHeader.join('|').toLowerCase() === header.join('|').toLowerCase();
  if (!headerMatches) rows.unshift(expectedHeader);

  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  clearBody_(body);

  var title = 'Broadcom Tanzu Security Advisories – last 30 days (' +
              Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + ')';
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Source: broadcom.com (GET json with fromDate/toDate) | Segment=VT | Last 30 days')
      .setItalic(true);

  var table = body.appendTable(rows);

  // Style header row
  var headerRow = table.getRow(0);
  for (var i = 0; i < headerRow.getNumCells(); i++) {
    headerRow.getCell(i).editAsText().setBold(true);
    headerRow.getCell(i).setBackgroundColor('#eeeeee');
  }
  table.setColumnWidth(0, 160); // Notification Id
  table.setColumnWidth(1, 160); // Release Date

  body.appendParagraph('Total advisories: ' + (rows.length - 1)).setBold(true);

  return 'Inserted ' + (rows.length - 1) + ' advisories into the document.';
}

/* ---------------- helpers ---------------- */
function clearBody_(body) {
  try {
    body.clear();
  } catch (e) {
    for (var i = body.getNumChildren() - 1; i >= 0; i--) {
      body.removeChild(body.getChild(i));
    }
  }
}

// For HtmlService includes (not strictly needed here, but handy if you expand)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
