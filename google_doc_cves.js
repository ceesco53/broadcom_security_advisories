/**
 * Google Docs: Paste Broadcom Tanzu advisories via cURL CSV (last 30 days)
 * - Adds "Broadcom CVEs → Paste CSV from cURL" to the Doc menu
 * - Dialog shows a prebuilt curl (segment=VT, last 30 days via jq filter)
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
  var curl = buildCurlCommand_(); // prebuilt with last-30-days cutoff
  var tmpl = HtmlService.createTemplateFromFile('paste_dialog');
  tmpl.curl = curl;
  var html = tmpl.evaluate().setWidth(700).setHeight(540);
  DocumentApp.getUi().showModalDialog(html, 'Broadcom CVEs (segment=VT)');
}

/**
 * Build a ready-to-run curl + jq command that outputs CSV of the last 30 days.
 * We compute the cutoff as a fixed ISO timestamp so the command is portable across shells.
 */
function buildCurlCommand_() {
  var tz = 'UTC';
  var now = new Date();
  var from = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  var cutoffISO = Utilities.formatDate(from, tz, "yyyy-MM-dd'T'00:00:00'Z'");
  var endpoint = "https://support.broadcom.com/web/ecx/security-advisory/-/securityadvisory/getSecurityAdvisoryList";

  // CSV columns: Notification Id, Release Date, Products, Level, Severity
  // Products joined with '; ' to avoid commas inside CSV fields.
  var cmd =
`# === Broadcom Tanzu advisories (segment=VT) – last 30 days ===
# 1) Copy everything between the "curl" and the final single quote.
# 2) Run it in your terminal (requires jq).
# 3) Paste the CSV output into the box below and click "Insert".

CUTOFF="${cutoffISO}"

curl -s '${endpoint}' \\
  -H 'accept: application/json' \\
  -H 'content-type: application/json' \\
  -H 'origin: https://support.broadcom.com' \\
  -H 'referer: https://support.broadcom.com/web/ecx/security-advisory?segment=VT' \\
  -H 'user-agent: Mozilla/5.0' \\
  --data-raw '{"pageNumber":0,"pageSize":200,"searchVal":"","segment":"VT","sortInfo":{"column":"","order":""}}' \\
| jq -r --arg cutoff "$CUTOFF" '
  def arr(x): (x // []) | if type=="array" then join("; ") else tostring end;
  (["Notification Id","Release Date","Products","Level","Severity"]),
  (
    .data.list
    | map(select(((.issueDate // .date) | fromdateiso8601) >= ($cutoff | fromdateiso8601)))
    | .[]
    | [
        (.advisoryId // .id),
        (.issueDate  // .date),
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
  body.appendParagraph('Source: support.broadcom.com (via cURL) | Segment=VT | Last 30 days')
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

// For HtmlService includes
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
