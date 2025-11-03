/**
 * Creates a Google Doc and populates it with the last 30 days of CVEs
 * in a table with columns:
 * Notification Id | Release Date | Products | Level | Severity
 *
 * Notes:
 * - "Notification Id" = CVE ID (e.g., CVE-2025-12345)
 * - "Release Date" = NVD "published" field (UTC)
 * - "Level" = CVSS Base Score (v3.1 -> v3.0 -> v2 fallback)
 * - "Severity" = Base Severity text (CRITICAL/HIGH/MEDIUM/LOW — same fallback)
 * - "Products" = deduped list of affected product names parsed from CPEs
 *
 * Optional: Add an NVD API key in Project Settings > Script properties:
 *   Name: NVD_API_KEY   Value: <your key>
 */

function buildCveDocLast30Days() {
  const now = new Date();
  const start = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);

  const pubStartDate = toIsoNoMillis(start); // e.g., 2025-10-04T00:00:00.000Z
  const pubEndDate   = toIsoNoMillis(now);

  const title = `CVEs – last 30 days (${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd')})`;
  const doc = DocumentApp.create(title);
  const body = doc.getBody();

  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(
    `Window: ${pubStartDate} to ${pubEndDate} (UTC)`
  ).setItalic(true);

  // Fetch CVEs (handles pagination)
  const cves = fetchAllCves(pubStartDate, pubEndDate);

  // Build table rows
  const header = ['Notification Id', 'Release Date', 'Products', 'Level', 'Severity'];
  const rows = [header];

  cves.forEach(vuln => {
    const cve = vuln.cve || vuln; // tolerate slight variations
    const id = cve.id || '';
    const published = cve.published || '';
    const { score, severity } = getCvss(cve);
    const products = extractProducts(cve);

    rows.push([
      id,
      published,
      products,
      score != null ? String(score) : '',
      severity || ''
    ]);
  });

  // Create the table
  const table = body.appendTable(rows);

  // Style header row
  const headerRow = table.getRow(0);
  for (let i = 0; i < headerRow.getNumCells(); i++) {
    headerRow.getCell(i).editAsText().setBold(true);
    headerRow.getCell(i).setBackgroundColor('#eeeeee');
  }

  // Fit columns a bit (Docs is limited; this mainly ensures text wrap)
  table.setColumnWidth(0, 130); // Notification Id
  table.setColumnWidth(1, 160); // Release Date

  body.appendParagraph(`Total CVEs: ${cves.length}`).setBold(true);

  Logger.log(`Created Doc: ${doc.getUrl()}`);
}

/** --- Helpers --- **/

function toIsoNoMillis(d) {
  // NVD accepts full ISO; keeping millis is fine, but we’ll retain them for clarity.
  // If you prefer no millis: return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
  return d.toISOString();
}

function fetchAllCves(pubStartDate, pubEndDate) {
  const baseUrl = 'https://services.nvd.nist.gov/rest/json/cves/2.0';
  const pageSize = 2000; // NVD max per page is typically 2000
  let startIndex = 0;
  let total = null;
  const all = [];

  const apiKey = PropertiesService.getScriptProperties().getProperty('NVD_API_KEY');

  // Keep paging until we collect totalResults (or hit a safety cap)
  for (let page = 0; page < 20; page++) { // safety: max ~40k entries
    const params = {
      pubStartDate: pubStartDate,
      pubEndDate: pubEndDate,
      startIndex: startIndex,
      resultsPerPage: pageSize,
      // You can add keywordSearch=, cvssV3Severity=, etc. if you want filters
    };

    const url = buildUrl(baseUrl, params);
    const options = {
      method: 'get',
      muteHttpExceptions: true,
      headers: apiKey ? { 'apiKey': apiKey } : {}
    };

    const resp = UrlFetchApp.fetch(url, options);
    if (resp.getResponseCode() !== 200) {
      throw new Error(`NVD API error ${resp.getResponseCode()}: ${resp.getContentText()}`);
    }

    const data = JSON.parse(resp.getContentText());
    const batch = (data.vulnerabilities || []).map(v => v); // already in v2.0 shape
    all.push(...batch);

    if (total == null) total = data.totalResults || batch.length;

    startIndex += batch.length;
    if (startIndex >= total || batch.length === 0) break;

    // Be polite with rate limits (esp. without API key)
    Utilities.sleep(apiKey ? 200 : 1200);
  }

  return all;
}

function buildUrl(base, params) {
  const esc = encodeURIComponent;
  const q = Object.keys(params)
    .filter(k => params[k] !== undefined && params[k] !== null)
    .map(k => `${esc(k)}=${esc(params[k])}`)
    .join('&');
  return `${base}?${q}`;
}

function getCvss(cve) {
  // Try CVSS v3.1, then v3.0, then v2
  try {
    const m31 = cve.metrics && cve.metrics.cvssMetricV31 && cve.metrics.cvssMetricV31[0];
    if (m31 && m31.cvssData) {
      return {
        score: m31.cvssData.baseScore,
        severity: m31.cvssData.baseSeverity
      };
    }
  } catch (e) {}

  try {
    const m30 = cve.metrics && cve.metrics.cvssMetricV30 && cve.metrics.cvssMetricV30[0];
    if (m30 && m30.cvssData) {
      return {
        score: m30.cvssData.baseScore,
        severity: m30.cvssData.baseSeverity
      };
    }
  } catch (e) {}

  try {
    const m20 = cve.metrics && cve.metrics.cvssMetricV2 && cve.metrics.cvssMetricV2[0];
    if (m20 && m20.cvssData) {
      // v2 stores severity at top level sometimes (baseSeverity may be undefined)
      return {
        score: m20.cvssData.baseScore,
        severity: m20.baseSeverity || (m20.cvssData.baseScore != null ? v2SeverityFromScore(m20.cvssData.baseScore) : '')
      };
    }
  } catch (e) {}

  return { score: null, severity: '' };
}

function v2SeverityFromScore(score) {
  // Rough mapping for v2 when baseSeverity not present
  if (score >= 7.0) return 'HIGH';
  if (score >= 4.0) return 'MEDIUM';
  if (score >= 0.0) return 'LOW';
  return '';
}

function extractProducts(cve) {
  // Pull product names from CPE criteria strings in configurations.nodes[].cpeMatch[].criteria
  // Dedup and join with commas. Fall back to blank if none.
  const products = new Set();

  const configs = (cve.configurations || []);
  configs.forEach(cfg => {
    const nodes = (cfg.nodes || []);
    nodes.forEach(node => {
      const matches = (node.cpeMatch || node.cpeMatches || []);
      matches.forEach(m => {
        if (m.vulnerable && m.criteria) {
          const product = parseProductFromCpe(m.criteria);
          if (product) products.add(product);
        }
      });
      // Some schemas use "children" for nested nodes
      const children = (node.children || []);
      children.forEach(child => {
        const cm = (child.cpeMatch || child.cpeMatches || []);
        cm.forEach(m => {
          if (m.vulnerable && m.criteria) {
            const product = parseProductFromCpe(m.criteria);
            if (product) products.add(product);
          }
        });
      });
    });
  });

  // If nothing found in configurations, try weaknesses/reference products (rarely present)
  // Keep it simple: we only use configurations for product list
  const list = Array.from(products);
  // Keep it readable—limit to 6, with a +N suffix if more
  if (list.length > 6) {
    const shown = list.slice(0, 6).join(', ');
    return `${shown} +${list.length - 6} more`;
  }
  return list.join(', ');
}

function parseProductFromCpe(cpeStr) {
  // CPE 2.3 format: cpe:2.3:a:vendor:product:version:update:...
  // We’ll return "vendor product" (product with hyphens/underscores normalized)
  try {
    const parts = cpeStr.split(':');
    // parts[2] is part (a/h/o), [3]=vendor, [4]=product
    const vendor = parts[3] || '';
    const product = parts[4] || '';
    if (!product) return '';
    const niceVendor = vendor.replace(/[_-]+/g, ' ');
    const niceProduct = product.replace(/[_-]+/g, ' ');
    return `${niceVendor} ${niceProduct}`.trim();
  } catch (e) {
    return '';
  }
}
