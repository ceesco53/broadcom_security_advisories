/**
 * Google Docs: Paste Broadcom Tanzu advisories via cURL CSV
 * - Adds "Broadcom CVEs → Paste CSV from cURL" to the Doc menu
 * - Dialog shows a prebuilt curl (last 30 days, segment=VT)
 * - You paste CSV output, and the script overwrites the Doc with a table
 *
 * Table columns: Notification Id | Release Date | Products | Level | Severity
 */

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Broadcom CVEs')
    .addItem('Paste CSV from cURL', 'showPasteDialog')
    .addToUi();
}

function showPasteDialog() {
  var curl = buildCurlCommand_(); // prefilled with dates
  var html = HtmlService.createHtmlOutputFromFile('paste_dialog')
    .setWidth(700)
    .setHeight(540);
  // Pass the curl string to the template
  html = HtmlService.createTemplateFromFile('paste_dialog');
  html.curl = curl;
  DocumentApp.getUi().showModalDialog(html.evaluate().setWidth(700).setHeight(540), 'Broadcom CVEs (segment=VT)');
}

/** Build a ready-to-run curl + jq that outputs CSV for the last 30 days (segment=VT). */
function buildCurlCommand_() {
  const tz = 'UTC';
  const now = new Date();
  const from = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  const cutoffISO = Utilities.formatDate(from, tz, "yyyy-MM-dd'T'00:00:00'Z'");
  const endpoint =
    "https://support.broadcom.com/web/ecx/security-advisory/-/securityadvisory/getSecurityAdvisoryList";

  const cmd = 
`# === Broadcom Tanzu advisories (segment=VT) – last 30 days ===
# Run this in your terminal (requires jq), then paste the CSV output below.

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

/** Server-side: parse pasted CSV and overwrite the Doc with a table. */
function insertCsvIntoDoc(csvText) {
  if (!csvText || !csvText.trim()) {
    throw new Error('No CSV content received.');
  }

  // Parse CSV into 2D array (rows x cols)
  var rows = Utilities.parseCsv(csvText);
  if (!rows || !rows.length) throw new Error('CSV parsing produced no rows.');

  // Ensure our header is correct; if not present, prepend it
  var expectedHeader = ['Notification Id','Release Date','Products','Level','Severity'];
  var header = rows[0].map(function(s){ return String(s || '').trim(); });
  var headerMatches = expectedHeader.join('|').toLowerCase() === header.join('|').toLowerCase();
  if (!headerMatches) {
    rows.unshift(expectedHeader);
  }

  // Overwrite current doc
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  clearBody_(body);

  var title = 'Broadcom Tanzu Security Advisories – last 30 days (' +
              Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + ')';
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Source: support.broadcom.com (via cURL) | Segment=VT')
      .setItalic(true);

  var table = body.appendTable(rows);

  // Style header row
  var headerRow = table.getRow(0);
  for (var i = 0; i < headerRow.getNumCells(); i++) {
    headerRow.getCell(i).editAsText().setBold(true);
    headerRow.getCell(i).setBackgroundColor('#eeeeee');
  }
  table.setColumnWidth(0, 160);
  table.setColumnWidth(1, 160);

  body.appendParagraph('Total advisories: ' + (rows.length - 1)).setBold(true);

  return 'Inserted ' + (rows.length - 1) + ' advisories into the document.';
}

/** Helpers */
function clearBody_(body) {
  try { body.clear(); }
  catch (e) {
    for (var i = body.getNumChildren() - 1; i >= 0; i--) body.removeChild(body.getChild(i));
  }
}
