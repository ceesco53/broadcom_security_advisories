function onOpen() {
  DocumentApp.getUi()
    .createMenu('Broadcom CVEs')
    .addItem('Fetch advisories (JSON)', 'showFetcher')
    .addToUi();
}

function showFetcher() {
  const defaults = getDefaults_();
  const tmpl = HtmlService.createTemplateFromFile('fetcher_sidebar');
  tmpl.defaults = defaults;
  DocumentApp.getUi().showSidebar(tmpl.evaluate().setTitle('Broadcom Advisories'));
}

function getDefaults_() {
  const tz = 'UTC';
  const now = new Date();
  const from = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  return {
    fromDate: Utilities.formatDate(from, tz, 'yyyy-MM-dd'),
    toDate:   Utilities.formatDate(now,  tz, 'yyyy-MM-dd'),
    segment:  'VT',
    pageSize: 10000
  };
}

/** === Public: used by "Fetch & Insert" button === */
function runFetchAndInsert(params) {
  const { fromDate, toDate, segment, pageSize } = params || {};
  const items = fetchAdvisories_(fromDate, toDate, segment, pageSize);

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  clearBody_(body);

  const titleText = `Broadcom Security Advisories – ${segment} (${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')})`;
  body.appendParagraph(titleText).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Source: support.broadcom.com | Window: ${fromDate} to ${toDate} (UTC) | Segment=${segment}`)
      .setItalic(true);

  // New columns: Id | Release Date | Title | Level
  const header = ['Id', 'Release Date', 'Title', 'Level'];
  const values = items.map(o => [o.id, o.issueDate, o.title, o.level]);
  const table = body.appendTable([header, ...values]);

  // Style header
  const headerRow = table.getRow(0);
  for (let i = 0; i < headerRow.getNumCells(); i++) {
    headerRow.getCell(i).editAsText().setBold(true);
    headerRow.getCell(i).setBackgroundColor('#eeeeee');
  }

  // Adjust column widths
  table.setColumnWidth(0, 100);  // Id (10-ish characters)
  table.setColumnWidth(1, 190);  // Release Date
  table.setColumnWidth(3, 120);  // Release Date
  // Let Title auto-size naturally (best fit)
  // Level column stays minimal width

  // Make Id clickable using notificationUrl
  for (let r = 1; r < table.getNumRows(); r++) {
    const item = items[r - 1];
    if (item && item.url) {
      const cell = table.getRow(r).getCell(0);
      const text = cell.editAsText();
      text.setText(item.id);
      text.setLinkUrl(item.url);
    }
  }

  body.appendParagraph(`Total advisories: ${items.length}`).setBold(true);
  return { count: items.length, segment, fromDate, toDate };
}

/** === Public: used by new "Test fetch" button (no document changes) === */
function testFetch(params) {
  const payload = {
    fromDate: params.fromDate,
    toDate: params.toDate,
    pageNumber: 0,
    pageSize: Number(params.pageSize || 10000),
    searchVal: '',
    segment: params.segment || 'VT',
    sortInfo: { column: '', order: '' }
  };

  const endpoint = 'https://support.broadcom.com/web/ecx/security-advisory/-/securityadvisory/getSecurityAdvisoryList';
  const options = {
    method: 'post',
    contentType: 'application/json;charset=UTF-8',
    muteHttpExceptions: true,
    headers: {
      accept: 'application/json, text/plain, */*',
      origin: 'https://support.broadcom.com',
      referer: 'https://support.broadcom.com/web/ecx/security-advisory?segment=' + payload.segment,
      'User-Agent': 'Mozilla/5.0 (AppsScript)'
    },
    payload: JSON.stringify(payload)
  };

  const resp = UrlFetchApp.fetch(endpoint, options);
  const status = resp.getResponseCode();
  const text = resp.getContentText() || '';
  let json, list = [];
  let firstKeys = [], firstItemSample = '';
  try {
    json = JSON.parse(text);
    list = (((json || {}).data || {}).list) || [];
    if (list.length) {
      firstKeys = Object.keys(list[0]).sort();
      firstItemSample = JSON.stringify(list[0], null, 2).slice(0, 1500);
    }
  } catch (e) {
    // leave as-is
  }

  return {
    httpStatus: status,
    payloadUsed: payload,
    rawLength: text.length,
    listLength: list.length,
    sampleIds: list.slice(0, 5).map(x =>
      x.notificationId || x.notificationID || x.notificationNo ||
      x.advisoryId || x.id || x.vmsaId || x.vmsa || null
    ),
    firstItemKeys: firstKeys,
    firstItemSample: firstItemSample,
    rawSnippet: list.length ? '' : text.slice(0, 1500)
  };
}

/** Internal fetcher used by runFetchAndInsert */
function fetchAdvisories_(fromDate, toDate, segment, pageSize) {
  if (!fromDate || !toDate) throw new Error('Please provide fromDate and toDate (YYYY-MM-DD).');
  segment = segment || 'VT';
  pageSize = Number(pageSize || 10000);

  const endpoint = 'https://support.broadcom.com/web/ecx/security-advisory/-/securityadvisory/getSecurityAdvisoryList';
  let pageNumber = 0, all = [];

  while (true) {
    const payload = {
      fromDate, toDate, pageNumber, pageSize,
      searchVal: '', segment, sortInfo: { column: '', order: '' }
    };

    const options = {
      method: 'post',
      contentType: 'application/json;charset=UTF-8',
      muteHttpExceptions: true,
      headers: {
        accept: 'application/json, text/plain, */*',
        origin: 'https://support.broadcom.com',
        referer: 'https://support.broadcom.com/web/ecx/security-advisory?segment=' + segment,
        'User-Agent': 'Mozilla/5.0 (AppsScript)'
      },
      payload: JSON.stringify(payload)
    };

    const resp = UrlFetchApp.fetch(endpoint, options);
    const code = resp.getResponseCode();
    if (code !== 200) throw new Error(`Fetch failed (${code}): ${resp.getContentText().slice(0, 500)}`);

    const json = JSON.parse(resp.getContentText() || '{}');
    const list = (((json || {}).data || {}).list) || [];
    all.push(...list);

    if (list.length < pageSize) break; // end of pages
    pageNumber += 1;
    Utilities.sleep(200);
  }

  // ---------- helpers ----------
  const pick = (obj, keys, fallback='') => {
    for (let k of keys) {
      if (obj != null && obj[k] != null && String(obj[k]).trim() !== '') return obj[k];
    }
    return fallback;
  };

  // Very-tolerant date parser for common Broadcom shapes
  const toDateObj = (val) => {
    if (val === null || val === undefined || val === '') return null;

    // numeric (epoch seconds or ms)
    if (typeof val === 'number') {
      // treat <1e12 as seconds
      const ms = val < 1e12 ? val * 1000 : val;
      const d = new Date(ms);
      return isNaN(d) ? null : d;
    }

    let s = String(val).trim();
    if (!s) return null;

    // plain yyyy-mm-dd -> UTC midnight
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return new Date(s + 'T00:00:00Z');

    // yyyy/mm/dd
    if (/^\d{4}\/\d{2}\/\d{2}$/.test(s)) {
      const [Y, M, D] = s.split('/');
      const d = new Date(Date.UTC(Number(Y), Number(M) - 1, Number(D)));
      return isNaN(d) ? null : d;
    }

    // mm/dd/yyyy
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
      const [m, d, y] = s.split('/').map(Number);
      const dt = new Date(Date.UTC(y, m - 1, d));
      return isNaN(dt) ? null : dt;
    }

    // "Oct 29, 2025" or "29 Oct 2025"
    const monthMap = {
      jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11
    };
    let m = s.match(/^(?:([A-Za-z]{3,})\s+(\d{1,2}),?\s+(\d{4}))|(?:(\d{1,2})\s+([A-Za-z]{3,}),?\s+(\d{4}))/);
    if (m) {
      if (m[1]) { // "Oct 29, 2025"
        const mon = monthMap[m[1].slice(0,3).toLowerCase()];
        const day = Number(m[2]), year = Number(m[3]);
        const dt = new Date(Date.UTC(year, mon, day));
        return isNaN(dt) ? null : dt;
      } else {    // "29 Oct 2025"
        const day = Number(m[4]);
        const mon = monthMap[m[5].slice(0,3).toLowerCase()];
        const year = Number(m[6]);
        const dt = new Date(Date.UTC(year, mon, day));
        return isNaN(dt) ? null : dt;
      }
    }

    // ISO-ish: "2025-10-29 13:22:10", "2025-10-29T13:22:10", with/without Z/offset
    s = s.replace(' ', 'T');
    if (/^\d{4}-\d{2}-\d{2}T/.test(s) && !/[Zz+\-]\d{2}:?\d{2}$/.test(s)) {
      // add Z if it looks like a full time without zone
      s += 'Z';
    }
    const d2 = new Date(s);
    if (!isNaN(d2)) return d2;

    // last resort: extract yyyy-mm-dd anywhere in the string
    const m2 = s.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (m2) {
      const dt = new Date(Date.UTC(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3])));
      return isNaN(dt) ? null : dt;
    }

    return null;
  };

  // inclusive UTC bounds
  const fromBound = toDateObj(fromDate + 'T00:00:00Z');
  const toBound   = toDateObj(toDate   + 'T23:59:59Z');

  const rows = [];
  const seen = new Set();

  all.forEach(item => {
    const id = pick(item, [
      'notificationId','notificationID','notificationNo',
      'advisoryId','id','vmsaId','vmsa'
    ]);
    if (!id) return;
    const idKey = String(id);
    if (seen.has(idKey)) return;
    seen.add(idKey);

    // include lots of candidate date keys Broadcom might use
    const dateRaw = pick(item, [
      'issueDate','releaseDate','published','publishDate',
      'date','releasedOn','publicationDate','lastUpdated',
      'postedDate','advisoryDate','createDate','createdDate','modifiedDate'
    ]);
    const d = toDateObj(dateRaw);
    if (!d) return;

    // filter to window (inclusive)
    if ((fromBound && d < fromBound) || (toBound && d > toBound)) return;

    const issueDate = Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');

    // Title cleanup + trim
    let title = pick(item, ['title','summary','subject','headline','description'], '');
    title = title.replace(/^Product Release Advisory\s*-\s*/i, '');
    if (title.length > 200) title = title.substring(0, 200) + '…';

    const level = pick(item, ['advisorySeverity','severity','level','riskLevel','threatLevel']);
    const url = pick(item, ['notificationUrl','advisoryUrl','detailUrl','url']);

    rows.push({
      id: idKey,
      issueDate,
      title,
      level: String(level || ''),
      url: String(url || '')
    });
  });

  return rows;
}

function clearBody_(body) {
  try { body.clear(); }
  catch (e) {
    for (let i = body.getNumChildren() - 1; i >= 0; i--) {
      body.removeChild(body.getChild(i));
    }
  }
}
