/**
 * Build a Google Doc with last 30 days of Broadcom security advisories
 * Columns: Notification Id | Release Date | Products | Level | Severity
 *
 * Data source: Broadcom advisories JSON documented here:
 * https://www.broadcom.com/support/vmware-security-advisories  (landing)
 * https://knowledge.broadcom.com/external/article/408302/json-api-for-product-security-advisories.html (API doc)
 *
 * Notes:
 * - We query several "segments" (divisions) and de-duplicate by advisoryId.
 * - "Level" is the advisory's overall severity (e.g., Moderate/High).
 * - "Severity" is the max CVSS severity text if available, else blank.
 */
function buildBroadcomAdvisoriesDoc() {
  const now = new Date();
  const from = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);

  const FROM = Utilities.formatDate(from, 'UTC', 'yyyy-MM-dd');
  const TO   = Utilities.formatDate(now,  'UTC', 'yyyy-MM-dd');

  // Tweak this list as needed; see Broadcom API docs for segment codes.
  const SEGMENTS = ['VC', 'TNZ', 'ANS']; // VMware Cloud, Tanzu, App Net & Sec (examples)

  // Public JSON endpoint documented by Broadcom
  const BASE = 'https://www.broadcom.com/support/security/advisories/json';

  const all = [];
  SEGMENTS.forEach(seg => {
    const url = `${BASE}?segment=${encodeURIComponent(seg)}&fromDate=${FROM}&toDate=${TO}&pageSize=500`;
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      // Non-fatal: continue other segments; log for debugging.
      console.warn(`Segment ${seg} fetch failed: ${resp.getResponseCode()} ${resp.getContentText()}`);
      return;
    }
    const data = JSON.parse(resp.getContentText() || '{}');
    // Expecting an array; if the shape changes, adjust mapping below.
    if (Array.isArray(data)) all.push(...data);
    else if (Array.isArray(data.items)) all.push(...data.items);
  });

  // De-duplicate by advisoryId
  const byId = new Map();
  all.forEach(item => {
    const id = safe(item, ['advisoryId']) || safe(item, ['id']) || '';
    if (!id) return;
    if (!byId.has(id)) byId.set(id, item);
  });

  // Build rows
  const header = ['Notification Id', 'Release Date', 'Products', 'Level', 'Severity'];
  const rows = [header];

  byId.forEach(advisory => {
    // Try multiple field names to be resilient to minor API changes.
    const id        = safe(advisory, ['advisoryId']) || safe(advisory, ['id']) || '';
    const issueDate = safe(advisory, ['issueDate']) || safe(advisory, ['date']) || '';
    const productsA = safe(advisory, ['impactedProducts']) || safe(advisory, ['products']) || [];
    const products  = Array.isArray(productsA) ? productsA.join(', ') : String(productsA || '');

    // Advisory-level severity (e.g., "High", "Moderate")
    const level =
      safe(advisory, ['advisorySeverity']) ||
      safe(advisory, ['severity']) || '';

    // CVSS max severity text (if available)
    const cvssRange = safe(advisory, ['cvssV3Range']) || '';      // e.g., "4.4–8.6"
    const cvssMax   = safe(advisory, ['cvssMaxSeverity']) || '';  // e.g., "CRITICAL/HIGH"
    const severity  = cvssMax || cvssTextFromRange(cvssRange);

    rows.push([id, issueDate, products, level, severity]);
  });

  // Create Doc
  const title = `Broadcom Security Advisories – last 30 days (${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd')})`;
  const doc = DocumentApp.create(title);
  const body = doc.getBody();

  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Source: Broadcom Support Security Advisories; Window: ${FROM} to ${TO} (UTC)`)
      .setItalic(true);

  const table = body.appendTable(rows);
  // Style header
  const headerRow = table.getRow(0);
  for (let i = 0; i < headerRow.getNumCells(); i++) {
    headerRow.getCell(i).editAsText().setBold(true);
    headerRow.getCell(i).setBackgroundColor('#eeeeee');
  }
  // Light column sizing (Docs has limited control)
  table.setColumnWidth(0, 140);
  table.setColumnWidth(1, 140);

  body.appendParagraph(`Total advisories: ${rows.length - 1}`).setBold(true);
  Logger.log(`Created Doc: ${doc.getUrl()}`);
}

/** Helpers **/

function safe(obj, pathArr) {
  try {
    return pathArr.reduce((o, k) => (o && k in o ? o[k] : undefined), obj);
  } catch (e) { return undefined; }
}

function cvssTextFromRange(rangeStr) {
  if (!rangeStr) return '';
  // If we get something like "4.4 - 8.6", map the max to a severity label.
  const m = String(rangeStr).match(/([\d.]+)\s*[-–]\s*([\d.]+)/);
  const max = m ? parseFloat(m[2]) : parseFloat(rangeStr);
  if (isNaN(max)) return '';
  if (max >= 9.0) return 'CRITICAL';
  if (max >= 7.0) return 'HIGH';
  if (max >= 4.0) return 'MEDIUM';
  if (max >= 0.1) return 'LOW';
  return '';
}
