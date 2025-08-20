/**
 * Taskabana → Export a SPECIFIC Google Tasks list to a Sheet tab "TaskSync"
 * 1) In Apps Script editor: Services (puzzle icon) → Enable "Tasks API".
 * 2) Also enable the Tasks API in the linked Google Cloud project when prompted.
 * 3) Run menu: Taskabana → Log Task List IDs, copy the ID you want, and set LIST_ID.
 */

const LIST_ID = 'PUT_YOUR_TASK_LIST_ID_HERE'; // <-- replace after running "Log Task List IDs"
const SHEET_NAME = 'TaskSync';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Taskabana')
    .addItem('Export Tasks to "TaskSync"', 'exportTasksToTaskSync')
    .addSeparator()
    .addItem('Log Task List IDs', 'logTaskListIds')
    .addToUi();
}

/** Lists all your Task Lists (title + id) in the Logs so you can set LIST_ID. */
function logTaskListIds() {
  // Single request — no pagination
  const resp = Tasks.Tasklists.list({
    maxResults: 100,
    fields: 'items(id,title)'
  });
  const items = (resp && resp.items) || [];

  // Log to execution log
  if (!items.length) {
    Logger.log('No task lists found for this account.');
  } else {
    items.forEach(l => Logger.log(`List: "${l.title}"\tID: ${l.id}`));
  }

  // Also write to a sheet for convenience
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('TaskLists') || ss.insertSheet('TaskLists');
  sh.clear();
  sh.getRange(1, 1, 1, 2).setValues([['Title', 'ID']]).setFontWeight('bold');

  if (items.length) {
    const rows = items.map(l => [l.title || '', l.id || '']);
    sh.getRange(2, 1, rows.length, 2).setValues(rows);
    sh.autoResizeColumns(1, 2);
    ss.toast(`Listed ${rows.length} task list(s) in the "TaskLists" tab.`, 'Taskabana', 5);
  } else {
    sh.getRange(2, 1).setValue('No task lists found.');
    ss.toast('No task lists found.', 'Taskabana', 5);
  }
}

function exportTasksToTaskSync() {
  // Validate LIST_ID
  try {
    const test = Tasks.Tasklists.get(LIST_ID);
    if (!test) {
      SpreadsheetApp.getUi().alert('LIST_ID was not found. Use "Log Task List IDs" and set LIST_ID.');
      return;
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Could not fetch that LIST_ID. Use "Log Task List IDs" and set LIST_ID.');
    console.error(e);
    return;
  }

  // Fetch all tasks (including completed + hidden), with pagination
  const tasks = fetchAllTasks_(LIST_ID);
  console.log(`Fetched ${tasks.length} tasks from list ${LIST_ID}`);

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
    'List ID'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  if (tasks.length === 0) {
    sheet.getRange(2, 1).setValue('No tasks found in this list.');
    sheet.setFrozenRows(1);
    return;
  }

  // Flattened hierarchical order (roots → children by position)
  const ordered = orderTasksHierarchically_(tasks);

  const rows = [];
  const titleRich = [];
  const notesRich = [];

  ordered.forEach(item => {
    const t = item.task;
    const level = item.level;

    rows.push([
      level,
      stripMarkdown_(t.title || ''),
      stripMarkdown_(t.notes || ''),
      t.status || 'needsAction',
      t.due ? new Date(t.due) : '',
      t.completed ? new Date(t.completed) : '',
      extractTags_(t.notes || '').join(', '),
      t.id || '',
      t.parent || '',
      t.position || '',
      t.updated ? new Date(t.updated) : '',
      LIST_ID
    ]);

    titleRich.push(buildRichTextFromMarkdown_(t.title || ''));
    notesRich.push(buildRichTextFromMarkdown_(t.notes || ''));
  });

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

  // Rich text for Title & Notes
  const titleCol = headers.indexOf('Title') + 1;
  const notesCol = headers.indexOf('Notes') + 1;
  applyRichTextColumn_(sheet, rowStart, titleCol, titleRich);
  applyRichTextColumn_(sheet, rowStart, notesCol, notesRich);

  // Nice layout
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(notesCol, 420);
  sheet.setColumnWidth(titleCol, 280);

  SpreadsheetApp.getUi().alert(`Exported ${rows.length} row(s) to "${SHEET_NAME}".`);
}

/** ----------- Helpers ----------- **/

function fetchAllTasks_(taskListId) {
  const out = [];
  let pageToken;
  do {
    const resp = Tasks.Tasks.list(taskListId, {
      showCompleted: true,
      showHidden: true,
      maxResults: 100,
      pageToken
    });
    (resp.items || []).forEach(t => out.push(t));
    pageToken = resp.nextPageToken;
  } while (pageToken);
  return out;
}

/** Flatten in hierarchical order using position, with level info for indentation. */
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

  const roots = tasks
    .filter(t => !t.parent)
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

/** Apply RichTextValue per-row to a given column. */
function applyRichTextColumn_(sheet, rowStart, col, richArray) {
  for (let i = 0; i < richArray.length; i++) {
    const rtv = richArray[i];
    if (rtv) sheet.getRange(rowStart + i, col).setRichTextValue(rtv);
  }
}

/** Simple Markdown → Sheets rich text (bold/italic/code/links + bullets). */
function buildRichTextFromMarkdown_(md) {
  md = md || '';
  md = md.replace(/\r\n?/g, '\n');

  // Headings → bold (rendered as **text**)
  md = md.replace(/^(#{1,6})\s*(.+)$/gm, (_, __, text) => `**${text.trim()}**`);

  // Bullet lines -> •
  md = md.replace(/^\s*[-*]\s+/gm, '• ');

  // Links: [label](url) → hyperlink markers
  const linkRuns = [];
  let text = md.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, (m, label, url) => {
    linkRuns.push({ label, url });
    return '\u0000' + label + '\u0001'; // sentinel
  });

  // Collect style runs for bold/italic/code
  const runs = [];

  function expandStyle(src, open, close, type) {
    let out = '';
    for (let i = 0; i < src.length; ) {
      const o = src.indexOf(open, i);
      if (o === -1) { out += src.slice(i); break; }
      out += src.slice(i, o);
      const c = src.indexOf(close, o + open.length);
      if (c === -1) { out += src.slice(o, o + open.length); i = o + open.length; continue; }
      const inner = src.slice(o + open.length, c);
      const start = out.length;
      out += inner;
      const end = out.length;
      runs.push({ start, end, type });
      i = c + close.length;
    }
    return out;
  }

  // Expand link markers first (record absolute ranges)
  let tmp = '';
  let linkIdx = 0;
  for (let i = 0; i < text.length; ) {
    if (text[i] === '\u0000') {
      const close = text.indexOf('\u0001', i + 1);
      if (close === -1) { tmp += text[i++]; continue; }
      const label = text.slice(i + 1, close);
      const start = tmp.length;
      tmp += label;
      const end = tmp.length;
      const url = linkRuns[linkIdx] && linkRuns[linkIdx].url;
      if (url) runs.push({ start, end, type: 'link', url });
      linkIdx++;
      i = close + 1;
    } else {
      tmp += text[i++];
    }
  }
  text = tmp;

  // Bold / italic / code
  text = expandStyle(text, '**', '**', 'bold');
  text = expandStyle(text, '*', '*', 'italic');
  text = expandStyle(text, '`', '`', 'code');

  const final = text;

  // Build RichText
  const rtv = SpreadsheetApp.newRichTextValue().setText(final);
  runs.forEach(run => {
    if (run.type === 'link') {
      rtv.setLinkUrl(run.start, run.end, run.url);
    } else {
      const builder = SpreadsheetApp.newTextStyle();
      if (run.type === 'bold') builder.setBold(true);
      if (run.type === 'italic') builder.setItalic(true);
      if (run.type === 'code') {
        builder.setFontFamily('Courier New');
        builder.setForegroundColor('#503');
      }
      rtv.setTextStyle(run.start, run.end, builder.build());
    }
  });

  return rtv.build();
}

/** Plain-text strip of minimal MD for non-rich columns. */
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

/** Extract #tags */
function extractTags_(notes) {
  if (!notes) return [];
  const tags = [];
  const re = /\B#([\p{L}\p{N}_-]+)/ug;
  let m;
  while ((m = re.exec(notes)) !== null) tags.push(m[1]);
  return Array.from(new Set(tags));
}