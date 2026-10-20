/**
 * IELTS Ladder → Google Sheets
 * ------------------------------------------------------------
 * Paste this whole file into Extensions → Apps Script of an empty
 * Google Sheet, then Deploy → New deployment → Web app
 * (Execute as: Me · Who has access: Anyone). See SETUP.md.
 *
 * Tabs it keeps up to date:
 *   Students  – one row per access code: levels, progress, last activity
 *   Results   – every scored section and drill attempt
 *   Essays    – every writing task, with the full text
 *   Speaking  – every speaking task
 *   _state    – hidden copy of each learner's full progress (for restore on a new device)
 *
 * Nothing here asks for names or emails. A row is identified by the access
 * code and the label you gave that code when you created it.
 */

const TABS = {
  Students: ['Code', 'Label', 'Last active', 'Last synced', 'Placement', 'Reading', 'Listening',
             'Writing', 'Speaking', 'Vocab & grammar', 'Sections done', 'Average %', 'Essays',
             'Speaking tasks', 'Last event (ms)'],
  Results:  ['When', 'Code', 'Label', 'Unit', 'Section', 'Drill', 'Tier', 'Score %', 'Mode', 'Level'],
  Essays:   ['When', 'Code', 'Label', 'Unit', 'Task', 'Tier', 'Words', 'Overall', 'Grammar', 'Spelling',
             'Prompt', 'Essay'],
  Speaking: ['When', 'Code', 'Label', 'Unit', 'Tier', 'Level', 'Words spoken', 'WPM', 'Fillers',
             'Questions answered', 'Target words used', 'Self-check'],
  _state:   ['Code', 'Updated (ms)', 'Data']
};
const CHUNK = 45000;           // a cell holds max 50 000 characters
const SKILLS = ['reading', 'listening', 'writing', 'speaking', 'language'];

/* ---------- web app endpoints ---------- */

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = JSON.parse(e.postData.contents);
    if (!p || !p.code || !p.state) return reply_({ ok: false, error: 'bad request' });
    const code = String(p.code).slice(0, 40);
    const label = String(p.label || '').slice(0, 120);
    const st = p.state;
    const units = p.units || {};
    const sections = p.sections || {};
    const ss = SpreadsheetApp.getActive();
    ensureTabs_(ss);

    // keep only the newest copy for restore
    const saved = readState_(ss, code);
    if (!saved || (Number(st.at) || 0) >= (Number(saved.at) || 0)) writeState_(ss, code, st);

    // append events newer than the last one we logged for this code
    const students = ss.getSheetByName('Students');
    const row = findRow_(students, code);
    const lastEvt = row ? Number(students.getRange(row, 15).getValue()) || 0 : 0;
    let newest = lastEvt;
    const when = t => new Date(Number(t));
    const unitName = id => units[id] ? id + ' · ' + units[id] : id;

    const res = [];
    Object.keys(st.results || {}).forEach(key => {
      const [uid, sid, did] = key.split(':');
      (st.results[key] || []).forEach(r => {
        if (!r || !(r.date > lastEvt)) return;
        res.push([when(r.date), code, label, unitName(uid), sections[uid + ':' + sid] || sid,
                  did || '', r.tier || '', r.pct, r.mode || '', r.lvl || '']);
        newest = Math.max(newest, r.date);
      });
    });
    res.sort((a, b) => a[0] - b[0]);
    append_(ss.getSheetByName('Results'), res);

    const ess = (st.essays || []).filter(x => x && x.date > lastEvt).map(x => {
      newest = Math.max(newest, x.date);
      return [when(x.date), code, label, unitName(x.uid), 'Task ' + (x.taskNo || ''), x.tier || '',
              x.words || '', flat_(x.overall), x.lt ? x.lt.grammar : '', x.lt ? x.lt.spelling : '',
              cell_(flat_(x.task)), cell_(x.text)];
    });
    append_(ss.getSheetByName('Essays'), ess);

    const spk = (st.speaking || []).filter(x => x && x.date > lastEvt).map(x => {
      newest = Math.max(newest, x.date);
      const m = x.metrics || {};
      return [when(x.date), code, label, unitName(x.uid), x.tier || '', flat_(x.level),
              m.words || '', m.wpm || '', m.fillers || '', x.answered || '', x.words || '', cell_(flat_(x.self))];
    });
    append_(ss.getSheetByName('Speaking'), spk);

    // one summary row per code
    const prof = st.profile || {};
    let done = 0, sum = 0, n = 0;
    Object.keys(st.results || {}).forEach(key => {
      const a = st.results[key] || [];
      if (!a.length) return;
      if (key.split(':').length === 2) done++;
      const last = a[a.length - 1];
      if (typeof last.pct === 'number') { sum += last.pct; n++; }
    });
    const summary = [code, label, st.at ? when(st.at) : '', new Date(),
      st.placement ? st.placement.level || '' : '',
      ...SKILLS.map(k => prof[k] || ''),
      done, n ? Math.round(sum / n) : '', (st.essays || []).length, (st.speaking || []).length, newest];
    if (row) students.getRange(row, 1, 1, summary.length).setValues([summary]);
    else students.appendRow(summary);

    return reply_({ ok: true, at: Number(st.at) || 0 });
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const code = e && e.parameter && e.parameter.code;
  if (!code) return reply_({ ok: true, service: 'IELTS Ladder sync' });
  const st = readState_(SpreadsheetApp.getActive(), String(code));
  return reply_({ ok: true, state: st || null });
}

/* ---------- helpers ---------- */

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ensureTabs_(ss) {
  Object.keys(TABS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
      if (name === '_state') sh.hideSheet();
      if (name === 'Students') sh.hideColumns(15);
    }
  });
  const blank = ss.getSheetByName('Sheet1') || ss.getSheetByName('Лист1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
}

function findRow_(sh, code) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const codes = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < codes.length; i++) if (String(codes[i][0]) === code) return i + 2;
  return 0;
}

function append_(sh, rows) {
  if (!rows.length) return;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function flat_(v) {
  if (v === null || v === undefined) return '';
  return typeof v === 'object' ? JSON.stringify(v) : v;
}

function cell_(v) {
  v = String(v || '');
  return v.length > CHUNK ? v.slice(0, CHUNK) + ' …' : v;
}

function readState_(ss, code) {
  const sh = ss.getSheetByName('_state');
  if (!sh) return null;
  const row = findRow_(sh, code);
  if (!row) return null;
  const width = sh.getLastColumn();
  const vals = sh.getRange(row, 3, 1, Math.max(1, width - 2)).getValues()[0];
  const json = vals.map(String).join('');
  try { return json ? JSON.parse(json) : null; } catch (err) { return null; }
}

function writeState_(ss, code, st) {
  const sh = ss.getSheetByName('_state');
  const json = JSON.stringify(st);
  const parts = [];
  for (let i = 0; i < json.length; i += CHUNK) parts.push(json.slice(i, i + CHUNK));
  let row = findRow_(sh, code);
  if (!row) row = sh.getLastRow() + 1;
  const oldWidth = Math.max(0, sh.getLastColumn() - 2);
  while (parts.length < oldWidth) parts.push('');   // clear leftovers from a longer copy
  if (sh.getMaxColumns() < parts.length + 2) sh.insertColumnsAfter(sh.getMaxColumns(), parts.length + 2 - sh.getMaxColumns());
  sh.getRange(row, 1, 1, 2).setValues([[code, Number(st.at) || 0]]);
  sh.getRange(row, 3, 1, parts.length).setNumberFormat('@').setValues([parts]);
}
