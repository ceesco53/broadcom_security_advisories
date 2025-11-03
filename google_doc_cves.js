/**
 * Overwrite an existing Google Doc with Broadcom Tanzu (segment=VT) advisories from the last 30 days.
 * Table columns: Notification Id | Release Date | Products | Level | Severity
 *
 * 1) Set DOC_ID to your target Google Doc ID.
 * 2) Run buildBroadcomTanzuDocIntoExisting()
 */

const DOC_ID = 'PASTE_YOUR_GOOGLE_DOC_ID_HERE';   // <-- 👈 paste your Doc ID

function buildBroadcomTanzuDocIntoExisting() {
  if (!DOC_ID || DOC_ID === 'PASTE_YOUR_GOOGLE_DOC_ID_HERE') {
    throw new Error('Please set DOC_ID to the target Google Doc ID.');
  }

  const tz = 'UTC';
  const now = new Date();
  const from = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  const FROM = Utilities.formatDate(from, tz, 'yyyy-MM-dd'); // ISO8601 date-only
  const TO   = Utilities.formatDate(now,  tz, 'yyyy-MM-dd');

  const BASE = 'https://www.broadcom.com/support/security/advisories/json';
  const SEGMENT = 'VT';  // Tanzu

  const url = `${BASE}?segment=${SEGMENT}&fromDate=${FROM}&toDate=${TO}&pageSize=500`;
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error(`Fetch failed (${resp.getResponseCode()}): ${resp.getContentText()}`);
  }

  const data = JSON.parse(resp.getContentText() || '[]');
  const items = Array.isArray(data) ? data : (Array.isArray(data.items) ? data.items : []);

  const header = ['Notification Id', 'Release Date', 'Products', 'Level', 'Severity'];
  const rows = [header];

  const seen = new Set();
  for (const adv of items) {
    const id = safe(adv, ['advisoryId']) || safe(adv, ['id']);
    if (!id || seen.has(id)) continue;
    seen.add(id);

    const issueDate = safe(adv, ['issueDate']) || safe(adv, ['date']) || '';
    const productsA = safe(adv, ['impactedProducts']) || safe(adv, ['products']) || [];
    const products  = Array.isArray(productsA) ? productsA.join(', ') : String(productsA || '');

    const level = safe(adv, ['advisorySeverity']) || safe(adv, ['severity']) || '';
    const cvssMax   = safe(adv, ['cvssMaxSeverity']) || '';
    const cvssRange = safe(adv, ['cvssV3Range']) || '';
    const severity  = cvssMax || cvssTextFromRange(cvssRange);

    rows.push([id, issueDate, products, level, severity]);
  }

  // Open the existing Doc and wipe its content
  const doc = DocumentApp.openById(DOC_ID);
  const body = doc.getBody();
  clearBody(body);

  // (Re)build the document contents
  const title = `Broadcom Tanzu Security Advisories – last 30 days (${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd')})`;
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Source: Broadcom Support (Security Advisories) | Segment=VT | Window: ${FROM} to ${TO} (UTC)`)
      .setItalic(true);

  const table = body.appendTable(rows);

  // Style header
  const headerRow = table.getRow(0);
  for (let i = 0; i < headerRow.getNumCells(); i++) {
    headerRow.getCell(i).editAsText().setBold(true);
    headerRow.getCell(i).setBackgroundColor('#eeeeee');
  }
  table.setColumnWidth(0, 150);
  table.setColumnWidth(1, 150);

  body.appendParagraph(`Total advisories: ${rows.length - 1}`).setBold(true);

  // Optionally set the Doc title (the file name in Drive) to match the heading
  doc.setName(title);

  Logger.log(`✅ Updated Doc: ${doc.getUrl()}`);
}

/* ---------- helpers ---------- */
function safe(obj, path) {
  try { return path.reduce((o, k) => (o && k in o ? o[k] : undefined), obj); } catch { return undefined; }
}

function cvssTextFromRange(rangeStr) {
  if (!rangeStr) return '';
  const m = String(rangeStr).match(/([\d.]+)\s*[-–]\s*([\d.]+)/);
  const max = m ? parseFloat(m[2]) : parseFloat(rangeStr);
  if (isNaN(max)) return '';
  if (max >= 9.0) return 'CRITICAL';
  if (max >= 7.0) return 'HIGH';
  if (max >= 4.0) return 'MEDIUM';
  if (max >= 0.1) return 'LOW';
  return '';
}

function clearBody(body) {
  // Robustly clear the doc body across Apps Script versions
  try {
    body.clear(); // Works in modern Apps Script
  } catch (e) {
    // Fallback: remove all child elements
    const num = body.getNumChildren();
    for (let i = num - 1; i >= 0; i--) {
      body.removeChild(body.getChild(i));
    }
  }
}
