// Spink Calculator – Google Sheet backend (Google Apps Script)
// Paste into Extensions > Apps Script of your Google Sheet, then Deploy > Web app.

const REC_HEAD = ['ID', 'Created', 'Updated', 'Mommy Name', 'Room No', 'Package', 'Items',
                  'Discount', 'Total', 'Deposit Paid', 'Paid', 'Balance', 'Status', 'Data',
                  'Upgrade Room', 'Preferred Room', 'Sign Up Date', 'No. of Babies']; // kept after Data so older sheets keep their column positions
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
  const records = sheet_('Records', REC_HEAD).getDataRange().getValues().slice(1).map(r => ({
    id: r[0],
    created: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
    mommy: r[3], room: r[4], pkg: r[5], data: r[13], upgrade: r[14] || '', pref: r[15] || '',
  }));
  return {
    ok: true, records,
    infants: inf.map(r => ({ id: r[0], name: r[1], babies: Number(r[2]), start: r[3], nights: Number(r[4]), room: r[5], note: r[6], u: Number(r[7]), deleted: r[8] === '1' })),
    classes: cls.map(r => ({ id: r[0], title: r[1], date: r[2], time: r[3], end: r[4], repeat: r[5], until: r[6], note: r[7], u: Number(r[8]), deleted: r[9] === '1' })),
    moms: moms.map(r => ({ id: r[0], u: Number(r[1]), data: r[7] })),
    treatments: trs.map(r => ({ id: r[0], momId: r[1], service: r[3], option: r[4], date: r[5], time: r[6], note: r[7], u: Number(r[8]), deleted: r[9] === '1' })),
  };
}

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
  } else if (sh.getMaxColumns() < head.length || sh.getRange(1, 1, 1, head.length).getValues()[0].join() !== head.join()) {
    // older sheet: add any newly introduced columns and headers
    if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'tracker_sync') return out_(trackerSync_(b));
    if (b.action !== 'save') return out_({ ok: false, error: 'Unknown action' });
    return out_(save_(b));
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    const rows = sheet_('Records', REC_HEAD).getDataRange().getValues().slice(1);
    // treatment usage written by the Mom Service Tracker (tab "Mommies"), e.g. "Massage 2/5 · Facial 0/1"
    const usage = {};
    const ms = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mommies');
    if (ms && ms.getLastRow() > 1) {
      ms.getRange(2, 1, ms.getLastRow() - 1, MOM_HEAD.length).getValues().forEach(r => { usage[r[0]] = r[6]; });
    }
    const records = rows.reverse().slice(0, 300).map(r => ({
      id: r[0], mommy: r[3], room: r[4], pkg: r[5], total: r[8], deposit: r[9], paid: r[10],
      balance: r[11], status: r[12], data: r[13], usage: usage[r[0]] || '',
    }));
    return out_({ ok: true, records });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function save_(b) {
  const pays = sheet_('Payments', PAY_HEAD);
  const recs = sheet_('Records', REC_HEAD);

  // each payment is typed 'Deposit' or 'Payment' by the calculator
  (b.payments || []).forEach(p => {
    if (p.amount > 0) {
      pays.appendRow([String(p.date), b.id, b.mommy, b.room, Number(p.amount), p.type, p.by, new Date()]);
    }
  });

  // deposit = the Deposit payment; paid = sum of all of this record's payments
  let paid = 0, deposit = 0;
  pays.getDataRange().getValues().slice(1).forEach(r => {
    if (r[1] !== b.id) return;
    paid += Number(r[4]) || 0;
    if (r[5] === 'Deposit') deposit += Number(r[4]) || 0;
  });

  const total = Number(b.total) || 0;
  const balance = Math.round((total - paid) * 100) / 100;
  const status = balance <= 0 ? 'Paid' : paid > 0 ? 'Partial' : 'Confirmed';
  const now = new Date();

  const ids = recs.getRange(1, 1, recs.getLastRow(), 1).getValues().map(r => r[0]);
  const idx = ids.indexOf(b.id);
  if (idx >= 1) {
    const row = idx + 1;
    const created = recs.getRange(row, 2).getValue();
    recs.getRange(row, 1, 1, REC_HEAD.length).setValues([[
      b.id, created, now, b.mommy, b.room, b.pkg, b.items, Number(b.discount) || 0, total, deposit, paid, balance, status, b.data, b.upgrade || "", b.pref || "", b.signup || "", b.babies || ""]]);
  } else {
    recs.appendRow([b.id, now, now, b.mommy, b.room, b.pkg, b.items, Number(b.discount) || 0, total, deposit, paid, balance, status, b.data, b.upgrade || "", b.pref || "", b.signup || "", b.babies || ""]);
  }
  return { ok: true, id: b.id, deposit, paid, balance, status };
}
