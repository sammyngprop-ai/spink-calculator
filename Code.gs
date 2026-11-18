// Spink Calculator – Google Sheet backend (Google Apps Script)
// Paste into Extensions > Apps Script of your Google Sheet, then Deploy > Web app.
//
// Records and Payments are read and written by HEADER NAME, not column position, so you can
// insert, move or add your own columns in those tabs without breaking the calculator.
// Keep the header names below. The script only ever adds a missing header at the end; it never
// renames or reorders your columns.

const REC_HEAD = ['ID', 'Created', 'Updated', 'Mommy Name', 'Room No', 'Package', 'Items',
                  'Discount', 'Total', 'Deposit Paid', 'Paid', 'Balance', 'Status', 'Data',
                  'Upgrade Room', 'Preferred Room', 'Sign Up Date', 'No. of Babies', 'Contact No'];
const PAY_HEAD = ['Payment Date', 'Record ID', 'Mommy Name', 'Room No', 'Amount', 'Type', 'Collected By', 'Logged At'];

// Mom Service Tracker tabs (all plain text so dates and numbers are stored exactly as sent)
const MOM_HEAD = ['Mom ID', 'Updated', 'Name', 'Room', 'Package', 'Check-in', 'Usage', 'Data'];
const CL_HEAD = ['Class ID', 'Title', 'Date', 'Time', 'End', 'Repeat', 'Until', 'Note', 'Updated', 'Deleted'];
const IC_HEAD = ['Infant ID', 'Name', 'Babies', 'Check-in', 'Nights', 'Room', 'Note', 'Updated', 'Deleted'];
const TR_HEAD = ['Treatment ID', 'Mom ID', 'Mommy Name', 'Service', 'Option', 'Date', 'Time', 'Note', 'Updated', 'Deleted'];

function textSheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
  }
  return sh;
}

// insert new rows, or replace an existing row only when the incoming one is newer (column uCol = updated ms)
function upsert_(sh, rows, uCol) {
  const last = sh.getLastRow();
  const width = rows.length ? rows[0].length : sh.getLastColumn();
  const existing = last > 1 ? sh.getRange(2, 1, last - 1, width).getValues().map(r => r.map(String)) : [];
  const pos = {};
  existing.forEach((r, i) => { pos[r[0]] = i; });
  let changed = false;
  rows.forEach(r => {
    const i = pos[r[0]];
    if (i === undefined) { pos[r[0]] = existing.length; existing.push(r); changed = true; }
    else if (Number(r[uCol]) > Number(existing[i][uCol])) { existing[i] = r; changed = true; }
  });
  if (changed && existing.length) sh.getRange(2, 1, existing.length, width).setValues(existing);
  return existing;
}

// Opens (or creates) a tab and finds every column in `head` by its header text.
// Returns { sh, idx, width }: idx[header] is the 0-based column, width is the number of header cells.
// A missing header is added after the last used header; existing headers are left alone.
function table_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
  }
  let hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(h => String(h).trim());
  const missing = head.filter(h => hdr.indexOf(h) < 0);
  if (missing.length) {
    let end = hdr.length;
    while (end > 0 && hdr[end - 1] === '') end--;
    if (sh.getMaxColumns() < end + missing.length) sh.insertColumnsAfter(sh.getMaxColumns(), end + missing.length - sh.getMaxColumns());
    sh.getRange(1, end + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
    hdr = hdr.slice(0, end).concat(missing);
  }
  const idx = {};
  head.forEach(h => { idx[h] = hdr.indexOf(h); });
  return { sh, idx, width: hdr.length };
}

// a new row as an array as wide as the header, with each value placed under its header
function rowFrom_(t, obj) {
  const row = new Array(t.width).fill('');
  Object.keys(obj).forEach(h => { row[t.idx[h]] = obj[h]; });
  return row;
}

function trackerSync_(b) {
  const str = v => String(v == null ? '' : v);
  const moms = upsert_(textSheet_('Mommies', MOM_HEAD),
    (b.moms || []).map(m => [m.id, m.u, m.name, m.room, m.pkg, m.start, m.usage, m.data].map(str)), 1);
  const trs = upsert_(textSheet_('Treatments', TR_HEAD),
    (b.treatments || []).map(t => [t.id, t.momId, t.name, t.service, t.option, t.date, t.time, t.note, t.u, t.deleted ? '1' : ''].map(str)), 8);

  const cls = upsert_(textSheet_('Classes', CL_HEAD),
    (b.classes || []).map(x => [x.id, x.title, x.date, x.time, x.end, x.repeat, x.until, x.note, x.u, x.deleted ? '1' : ''].map(str)), 8);

  const inf = upsert_(textSheet_('InfantCare', IC_HEAD),
    (b.infants || []).map(x => [x.id, x.name, x.babies, x.start, x.nights, x.room, x.note, x.u, x.deleted ? '1' : ''].map(str)), 7);

  const tz = Session.getScriptTimeZone();
  const t = table_('Records', REC_HEAD);
  const records = t.sh.getDataRange().getValues().slice(1).map(r => ({
    id: r[t.idx['ID']],
    created: r[t.idx['Created']] instanceof Date ? Utilities.formatDate(r[t.idx['Created']], tz, 'yyyy-MM-dd') : String(r[t.idx['Created']]),
    mommy: r[t.idx['Mommy Name']], room: r[t.idx['Room No']], pkg: r[t.idx['Package']], data: r[t.idx['Data']],
    upgrade: r[t.idx['Upgrade Room']] || '', pref: r[t.idx['Preferred Room']] || '',
  }));
  return {
    ok: true, records,
    infants: inf.map(r => ({ id: r[0], name: r[1], babies: Number(r[2]), start: r[3], nights: Number(r[4]), room: r[5], note: r[6], u: Number(r[7]), deleted: r[8] === '1' })),
    classes: cls.map(r => ({ id: r[0], title: r[1], date: r[2], time: r[3], end: r[4], repeat: r[5], until: r[6], note: r[7], u: Number(r[8]), deleted: r[9] === '1' })),
    moms: moms.map(r => ({ id: r[0], u: Number(r[1]), data: r[7] })),
    treatments: trs.map(r => ({ id: r[0], momId: r[1], service: r[3], option: r[4], date: r[5], time: r[6], note: r[7], u: Number(r[8]), deleted: r[9] === '1' })),
  };
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'contacts') return out_(contacts_(b));
    if (b.action === 'tracker_sync') return out_(trackerSync_(b));
    if (b.action !== 'save') return out_({ ok: false, error: 'Unknown action' });
    return out_(save_(b));
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Payments between two dates (yyyy-mm-dd, inclusive) for the Incoming Sales daily report: ?action=payments&from=...&to=...
function payments_(p) {
  const tz = Session.getScriptTimeZone();
  const from = String(p.from || ''), to = String(p.to || p.from || '');
  const day = v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v).slice(0, 10);
  const t = table_('Payments', PAY_HEAD), c = t.idx;
  const payments = t.sh.getDataRange().getValues().slice(1)
    .filter(r => Number(r[c['Amount']]) > 0)
    .map(r => ({ date: day(r[c['Payment Date']]), id: r[c['Record ID']], mommy: r[c['Mommy Name']], room: r[c['Room No']],
                 amount: Number(r[c['Amount']]), type: r[c['Type']], by: r[c['Collected By']] }))
    .filter(x => (!from || x.date >= from) && (!to || x.date <= to));
  return { ok: true, payments };
}

// Confirmed mommies from the Mom Service Tracker (tab "Mommies") for the Incoming Sales Rooms view: ?action=mommies
function mommies_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mommies');
  if (!sh || sh.getLastRow() < 2) return { ok: true, mommies: [] };
  const mommies = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, MOM_HEAD.length).getValues().forEach(r => {
    let d = {};
    try { d = JSON.parse(r[7] || '{}'); } catch (err) {}
    if (d.deleted) return;
    const start = String(r[5] || d.start || '').slice(0, 10), days = Number(d.days) || 0, name = String(r[2] || d.name || '');
    if (!name || !/^\d{4}-\d\d-\d\d$/.test(start) || !days) return;
    mommies.push({ id: String(r[0]), name, room: String(r[3] || d.room || ''), pkg: String(r[4] || d.pkgName || ''), start, days, babies: Number(d.babies) || 0 });
  });
  return { ok: true, mommies };
}

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'mommies') return out_(mommies_());
    if (e && e.parameter && e.parameter.action === 'payments') return out_(payments_(e.parameter));
    const t = table_('Records', REC_HEAD), c = t.idx;
    const rows = t.sh.getDataRange().getValues().slice(1);
    // treatment usage written by the Mom Service Tracker (tab "Mommies"), e.g. "Massage 2/5 · Facial 0/1"
    const usage = {};
    const ms = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mommies');
    if (ms && ms.getLastRow() > 1) {
      ms.getRange(2, 1, ms.getLastRow() - 1, MOM_HEAD.length).getValues().forEach(r => { usage[r[0]] = r[6]; });
    }
    const records = rows.reverse().slice(0, 300).map(r => ({
      id: r[c['ID']], mommy: r[c['Mommy Name']], room: r[c['Room No']], pkg: r[c['Package']],
      total: r[c['Total']], deposit: r[c['Deposit Paid']], paid: r[c['Paid']], balance: r[c['Balance']],
      status: r[c['Status']], data: r[c['Data']], usage: usage[r[c['ID']]] || '',
      contact: r[c['Contact No']] ? String(r[c['Contact No']]) : '',
    }));
    return out_({ ok: true, records });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function save_(b) {
  const pt = table_('Payments', PAY_HEAD), pc = pt.idx;
  const rt = table_('Records', REC_HEAD), rc = rt.idx;

  // each payment is typed 'Deposit' or 'Payment' by the calculator
  (b.payments || []).forEach(p => {
    if (p.amount > 0) {
      pt.sh.appendRow(rowFrom_(pt, {
        'Payment Date': String(p.date), 'Record ID': b.id, 'Mommy Name': b.mommy, 'Room No': b.room,
        'Amount': Number(p.amount), 'Type': p.type, 'Collected By': p.by, 'Logged At': new Date(),
      }));
    }
  });

  // deposit = the Deposit payment; paid = sum of all of this record's payments
  let paid = 0, deposit = 0;
  pt.sh.getDataRange().getValues().slice(1).forEach(r => {
    if (r[pc['Record ID']] !== b.id) return;
    paid += Number(r[pc['Amount']]) || 0;
    if (r[pc['Type']] === 'Deposit') deposit += Number(r[pc['Amount']]) || 0;
  });

  const total = Number(b.total) || 0;
  const balance = Math.round((total - paid) * 100) / 100;
  const status = balance <= 0 ? 'Paid' : paid > 0 ? 'Partial' : 'Confirmed';
  const now = new Date();

  // contact numbers are plain text so a leading 0 is never dropped
  rt.sh.getRange(2, rc['Contact No'] + 1, Math.max(rt.sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');

  const fields = {
    'Updated': now, 'Mommy Name': b.mommy, 'Room No': b.room, 'Package': b.pkg, 'Items': b.items,
    'Discount': Number(b.discount) || 0, 'Total': total, 'Deposit Paid': deposit, 'Paid': paid, 'Balance': balance,
    'Status': status, 'Data': b.data, 'Upgrade Room': b.upgrade || '', 'Preferred Room': b.pref || '',
    'Sign Up Date': b.signup || '', 'No. of Babies': b.babies || '',
  };
  if (b.contact) fields['Contact No'] = String(b.contact); // an empty contact never wipes a number already in the sheet

  const last = rt.sh.getLastRow();
  const ids = last > 1 ? rt.sh.getRange(2, rc['ID'] + 1, last - 1, 1).getValues().map(r => r[0]) : [];
  const at = ids.indexOf(b.id);
  if (at >= 0) {
    const row = at + 2;
    const rng = rt.sh.getRange(row, 1, 1, rt.width);
    const vals = rng.getValues()[0], formulas = rng.getFormulas()[0];
    const out = vals.map((v, i) => formulas[i] ? formulas[i] : v); // keep any formulas in your own columns
    Object.keys(fields).forEach(h => { out[rc[h]] = fields[h]; });
    rng.setValues([out]);
  } else {
    rt.sh.appendRow(rowFrom_(rt, Object.assign({ 'ID': b.id, 'Created': now, 'Contact No': String(b.contact || '') }, fields)));
  }
  return { ok: true, id: b.id, deposit, paid, balance, status };
}

// Bulk-add contact numbers: b.items = [{name, contact}]. Matches Records by Mommy Name (ignores case and extra spaces).
function contacts_(b) {
  const t = table_('Records', REC_HEAD), c = t.idx;
  const last = t.sh.getLastRow();
  const norm = x => String(x || '').toLowerCase().replace(/\s+/g, ' ').trim();
  const names = last > 1 ? t.sh.getRange(2, c['Mommy Name'] + 1, last - 1, 1).getValues().map(r => norm(r[0])) : [];
  const rng = last > 1 ? t.sh.getRange(2, c['Contact No'] + 1, last - 1, 1) : null;
  const col = rng ? rng.getValues() : [];
  let updated = 0;
  const notFound = [];
  (b.items || []).forEach(it => {
    const key = norm(it.name);
    let hit = false;
    names.forEach((n, i) => { if (n === key && it.contact) { col[i][0] = String(it.contact); hit = true; updated++; } });
    if (!hit) notFound.push(it.name);
  });
  if (rng) {
    rng.setNumberFormat('@');
    rng.setValues(col);
  }
  return { ok: true, updated, notFound };
}
