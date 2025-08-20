/**
 * Taskabana → Export Google Tasks to a Sheet tab "TaskSync"
 * Setup (one time):
 * 1) In Apps Script editor: Services (puzzle icon) → Enable "Tasks API".
 * 2) It will prompt you to enable Tasks API in the cloud project as well.
 * 3) Adjust LIST_NAME below (or leave null to pick the user's default list).
 */

const LIST_NAME = null; // e.g. "My Tasks". If null, uses the first (primary) list.
const SHEET_NAME = 'TaskSync';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Taskabana')
    .addItem('Export Tasks to "TaskSync"', 'exportTasksToTaskSync')
    .addToUi();
}

function exportTasksToTaskSync() {
  const list = resolveTaskList_(LIST_NAME);
  if (!list) {
    SpreadsheetApp.getUi().alert('Could not find a Google Tasks list. Check LIST_NAME or your Tasks account.');
    return;
  }

  const tasks = fetchAllTasks_(list.id);
  const ordered = orderTasksHierarchically_(tasks);

  const sheet = getOrCreateSheet_(SHEET_NAME);
  sheet.clear();

  const headers = [
    'Level',
    'Title',
    'Notes',
    'Status',
    'Due',
    'Completed',
    'Tags',
    'Task ID',
    'Parent ID',
    'Position',
    'Updated',
    'List'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  // Build rows (plain values first for speed)
  const rows = [];
  const titleRich = [];
  const notesRich = [];
  const tagsRegex = /\B#[\p{L}\p{N}_-]+/ug; // simple #tag detector

  ordered.forEach(item => {
    const t = item.task;
    const level = item.level;
    const tags = extractTags_(t.notes || '');
    rows.push([
      level,
      stripMarkdown_(t.title || ''),
      stripMarkdown_(t.notes || ''),
      t.status || 'needsAction',
      t.due ? new Date(t.due) : '',
      t.completed ? new Date(t.completed) : '',
      tags.join(', '),
      t.id || '',
      t.parent || '',
      t.position || '',
      t.updated ? new Date(t.updated) : '',
      list.title || ''
    ]);
    titleRich.push(buildRichTextFromMarkdown_(t.title || ''));
    notesRich.push(buildRichTextFromMarkdown_(t.notes || ''));
  });

  if (rows.length === 0) {
    sheet.getRange(2, 1).setValue('No tasks found.');
    return;
  }

  // Write data
  const range = sheet.getRange(2, 1, rows.length, headers.length);
  range.setValues(rows);

  // Dates formatting
  const rowStart = 2;
  const dueCol = headers.indexOf('Due') + 1;
  const completedCol = headers.indexOf('Completed') + 1;
  const updatedCol = headers.indexOf('Updated') + 1;
  sheet.getRange(rowStart, dueCol, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(rowStart, completedCol, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange(rowStart, updatedCol, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');

  // Apply rich text for Title and Notes columns
  const titleCol = headers.indexOf('Title') + 1;
  const notesCol = headers.indexOf('Notes') + 1;
  applyRichTextColumn_(sheet, rowStart, titleCol, titleRich);
  applyRichTextColumn_(sheet, rowStart, notesCol, notesRich);

  // Autofit
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
  // Make Notes a bit wider
  sheet.setColumnWidth(notesCol, 420);
  sheet.setColumnWidth(titleCol, 280);
}

/** ---------- Helpers ---------- **/

function resolveTaskList_(nameOrNull) {
  let pageToken;
  do {
    const resp = Tasks.Tasklists.list({ maxResults: 100, pageToken });
    const items = resp.items || [];
    if (nameOrNull) {
      const found = items.find(x => (x.title || '').trim() === nameOrNull);
      if (found) return found;
    } else {
      // default to the first list
      if (items.length > 0) return items[0];
    }
    pageToken = resp.nextPageToken;
  } while (pageToken);
  return null;
}

function fetchAllTasks_(taskListId) {
  const out = [];
  let pageToken;
  do {
    const resp = Tasks.Tasks.list(taskListId, {
      showDeleted: false,
      showHidden: true,
      maxResults: 100,
      pageToken
    });
    (resp.items || []).forEach(t => out.push(t));
    pageToken = resp.nextPageToken;
  } while (pageToken);
  return out;
}

/**
 * Order tasks in a flattened, hierarchical sequence:
 * - roots first (no parent), lexicographically by position
 * - then DFS their children (also by position)
 */
function orderTasksHierarchically_(tasks) {
  const byId = new Map();
  const children = new Map();
  tasks.forEach(t => {
    byId.set(t.id, t);
    if (t.parent) {
      const arr = children.get(t.parent) || [];
      arr.push(t);
      children.set(t.parent, arr);
    }
  });
  // sort children by position
  children.forEach(arr => arr.sort((a, b) => (a.position || '').localeCompare(b.position || '')));

  const roots = tasks.filter(t => !t.parent)
    .sort((a, b) => (a.position || '').localeCompare(b.position || ''));

  const result = [];
  function dfs(task, level) {
    result.push({ task, level });
    const kids = children.get(task.id) || [];
    kids.forEach(k => dfs(k, level + 1));
  }
  roots.forEach(r => dfs(r, 0));
  return result;
}

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/**
 * Apply RichTextValue per-row to a given column.
 * richArray: array of RichTextValue objects matching number of rows.
 */
function applyRichTextColumn_(sheet, rowStart, col, richArray) {
  for (let i = 0; i < richArray.length; i++) {
    const rtv = richArray[i];
    if (rtv) {
      sheet.getRange(rowStart + i, col).setRichTextValue(rtv);
    }
  }
}

/**
 * Basic Markdown → Sheets Rich Text:
 * Supports:
 * - **bold** and *italic*
 * - `code` (monospace + light bg)
 * - [label](url) (hyperlinks)
 * - Headings (#, ##, ###) → bold entire line
 * - Bullet lines (- item or * item) → prefix with •
 * Other MD is left as plain text (safely).
 */
function buildRichTextFromMarkdown_(md) {
  md = md || '';
  // Normalize line endings
  md = md.replace(/\r\n?/g, '\n');

  // Convert headings to bold lines
  md = md.replace(/^(#{1,6})\s*(.+)$/gm, (_, hashes, text) => {
    return '**' + text.trim() + '**';
  });

  // Convert bullet lines to bullets
  md = md.replace(/^\s*[-*]\s+/gm, '• ');

  // Handle links first: [label](url)
  const linkRuns = [];
  let text = md.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, (m, label, url) => {
    // we place the label in text and record a link run
    linkRuns.push({ label, url });
    return '\u0000' + label + '\u0001'; // markers for later mapping
  });

  // Now process bold and italic and code into markers we can style by indices.
  // Use simple non-greedy runs; nested combos are not fully supported.
  const boldRanges = [];
  text = text.replace(/\*\*(.+?)\*\*/g, (m, inner) => {
    boldRanges.push(inner);
    return '\u0002' + inner + '\u0003';
  });

  const italicRanges = [];
  text = text.replace(/\*(.+?)\*/g, (m, inner) => {
    italicRanges.push(inner);
    return '\u0004' + inner + '\u0005';
  });

  const codeRanges = [];
  text = text.replace(/`([^`]+)`/g, (m, inner) => {
    codeRanges.push(inner);
    return '\u0006' + inner + '\u0007';
  });

  // Build final plain string while recording absolute index ranges
  let final = '';
  const runs = []; // {start,end,type,extra}
  function consumeMarkerSegments(src, openMarker, closeMarker, type) {
    let idx = 0;
    for (;;) {
      const open = src.indexOf(openMarker, idx);
      if (open === -1) {
        final += src.slice(idx);
        break;
      }
      final += src.slice(idx, open);
      const afterOpen = open + openMarker.length;
      const close = src.indexOf(closeMarker, afterOpen);
      if (close === -1) {
        // no close; treat marker as text
        final += src.slice(open, afterOpen);
        idx = afterOpen;
        continue;
      }
      const inner = src.slice(afterOpen, close);
      const start = final.length;
      final += inner;
      const end = final.length;
      runs.push({ start, end, type });
      idx = close + closeMarker.length;
    }
  }

  // First apply link placeholders, then bold/italic/code placeholders
  // Step 1: Expand link markers, recording link runs
  // Replace each \u0000label\u0001 sequentially
  let linkIdx = 0;
  let tmp = '';
  for (let i = 0; i < text.length; ) {
    const c = text[i];
    if (c === '\u0000') {
      // find closing marker
      const close = text.indexOf('\u0001', i + 1);
      if (close === -1) {
        tmp += c;
        i++;
        continue;
      }
      const label = text.slice(i + 1, close);
      const start = (tmp + '').length;
      tmp += label;
      const end = tmp.length;
      const url = linkRuns[linkIdx] && linkRuns[linkIdx].url;
      if (url) runs.push({ start, end, type: 'link', url });
      linkIdx++;
      i = close + 1;
    } else {
      tmp += c;
      i++;
    }
  }
  text = tmp;

  // Step 2: bold, italic, code markers to ranges
  function expandStyle(src, open, close, type) {
    let out = '';
    for (let i = 0; i < src.length; ) {
      const o = src.indexOf(open, i);
      if (o === -1) {
        out += src.slice(i);
        break;
      }
      out += src.slice(i, o);
      const c = src.indexOf(close, o + open.length);
      if (c === -1) {
        out += src.slice(o, o + open.length);
        i = o + open.length;
        continue;
      }
      const inner = src.slice(o + open.length, c);
      const start = (final + out).length;
      out += inner;
      const end = (final + out).length;
      runs.push({ start, end, type });
      i = c + close.length;
    }
    return out;
  }

  text = expandStyle(text, '\u0002', '\u0003', 'bold');
  text = expandStyle(text, '\u0004', '\u0005', 'italic');
  text = expandStyle(text, '\u0006', '\u0007', 'code');

  final += text;

  // Build RichTextValue with all styles
  const rtv = SpreadsheetApp.newRichTextValue().setText(final);
  // Merge overlapping runs naïvely; Sheets can handle multiple setTextStyle calls.
  runs.forEach(run => {
    const style = SpreadsheetApp.newTextStyle()
      .setBold(run.type === 'bold' || undefined)
      .setItalic(run.type === 'italic' || undefined)
      .setFontFamily(run.type === 'code' ? 'Courier New' : undefined)
      .setForegroundColor(run.type === 'code' ? '#503' : undefined)
      .build();

    if (run.type === 'link') {
      rtv.setLinkUrl(run.start, run.end, run.url);
    } else {
      rtv.setTextStyle(run.start, run.end, style);
    }
  });

  return rtv.build();
}

/** Strip MD to plain text for non-rich columns */
function stripMarkdown_(s) {
  if (!s) return '';
  s = s.replace(/\r\n?/g, '\n');
  s = s.replace(/^(#{1,6})\s*(.+)$/gm, '$2'); // headings
  s = s.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, '$1 ($2)'); // links
  s = s.replace(/\*\*(.+?)\*\*/g, '$1'); // bold
  s = s.replace(/\*(.+?)\*/g, '$1');    // italic
  s = s.replace(/`([^`]+)`/g, '$1');    // code
  s = s.replace(/^\s*[-*]\s+/gm, '• '); // bullets
  return s;
}

/** Extract inline #tags (e.g., #icebucket) */
function extractTags_(notes) {
  if (!notes) return [];
  const tags = [];
  const re = /\B#([\p{L}\p{N}_-]+)/ug;
  let m;
  while ((m = re.exec(notes)) !== null) {
    tags.push(m[1]);
  }
  // de-dupe
  return Array.from(new Set(tags));
}
