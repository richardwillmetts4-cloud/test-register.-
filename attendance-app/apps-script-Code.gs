// Paste into Extensions > Apps Script in the Google Sheet. Then Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).
const TEACHER_PASSWORD = 'teacher123';   // shared app password for all teachers
const ADMIN_PASSWORD   = 'admin123';     // special summary login
const FIRST_ROW = 5, LAST_ROW = 19, NOTES_ROW = 21, DATE_ROW = 4;

const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const roleOf = pw => pw === ADMIN_PASSWORD ? 'admin' : pw === TEACHER_PASSWORD ? 'teacher' : null;
const findSheet = n => SpreadsheetApp.getActive().getSheets().find(s => s.getName().trim().toLowerCase() === String(n||'').trim().toLowerCase());

function readGroup(sh) {
  const cols = Math.max(sh.getLastColumn(), 2);
  const v = sh.getRange(1, 1, NOTES_ROW, cols).getDisplayValues();
  const dates = [];
  for (let c = 1; c < cols; c++) if (v[DATE_ROW-1][c]) dates.push({ col: c, label: v[DATE_ROW-1][c] });
  const pupils = [], marks = [];
  for (let r = FIRST_ROW-1; r < LAST_ROW; r++) if (v[r][0]) { pupils.push(v[r][0]); marks.push(dates.map(d => v[r][d.col])); }
  return { name: sh.getName(), time: v[1][1], dates: dates.map(d => d.label), pupils, marks, notes: dates.map(d => v[NOTES_ROW-1][d.col]) };
}

function doGet(e) {
  const p = e.parameter, role = roleOf(p.pw);
  if (!role) return out({ error: 'Wrong password' });
  if (p.action === 'login') return out({ role });
  if (p.action === 'group') {
    const sh = findSheet(p.group);
    return sh ? out(readGroup(sh)) : out({ error: 'Group not found' });
  }
  if (p.action === 'all') {
    if (role !== 'admin') return out({ error: 'Admin only' });
    return out({ groups: SpreadsheetApp.getActive().getSheets().map(readGroup) });
  }
  return out({ error: 'Unknown action' });
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (!roleOf(b.pw)) return out({ error: 'Wrong password' });
  const sh = findSheet(b.group);
  if (!sh) return out({ error: 'Group not found' });
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const g = readGroup(sh), di = g.dates.indexOf(b.date);
    if (di < 0) return out({ error: 'Date not found' });
    const col = sh.getRange(DATE_ROW, 1, 1, sh.getLastColumn()).getDisplayValues()[0].indexOf(b.date) + 1;
    sh.getRange(FIRST_ROW, col, b.marks.length, 1).setValues(b.marks.map(m => [m]));
    sh.getRange(NOTES_ROW, col).setValue(b.note || '');
    return out({ ok: true });
  } finally { lock.release(); }
}
