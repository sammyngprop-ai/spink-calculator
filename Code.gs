// Spink Calculator – Google Sheet backend (Google Apps Script)
// Paste into Extensions > Apps Script of your Google Sheet, then Deploy > Web app.

const KEY = 'CHANGE-ME'; // must match the "Secret key" in the calculator's ⚙ Sheet settings

const REC_HEAD = ['ID', 'Created', 'Updated', 'Mommy Name', 'Room No', 'Package', 'Items',
                  'Discount', 'Total', 'Deposit Paid', 'Paid', 'Balance', 'Status', 'Data'];
const PAY_HEAD = ['Payment Date', 'Record ID', 'Mommy Name', 'Room No', 'Amount', 'Type', 'Collected By', 'Logged At'];

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
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
    if (b.key !== KEY) return out_({ ok: false, error: 'Wrong secret key' });
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
    if (e.parameter.key !== KEY) return out_({ ok: false, error: 'Wrong secret key' });
    const rows = sheet_('Records', REC_HEAD).getDataRange().getValues().slice(1);
    const records = rows.reverse().slice(0, 300).map(r => ({
      id: r[0], mommy: r[3], room: r[4], pkg: r[5], total: r[8], deposit: r[9], paid: r[10],
      balance: r[11], status: r[12], data: r[13],
    }));
    return out_({ ok: true, records });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function save_(b) {
  const pays = sheet_('Payments', PAY_HEAD);
  const recs = sheet_('Records', REC_HEAD);

  // the first payment recorded for a customer is the deposit; later ones are plain payments
  const hasPayment = pays.getDataRange().getValues().slice(1).some(r => r[1] === b.id);
  if (b.payment && b.payment.amount > 0) {
    pays.appendRow([String(b.payment.date), b.id, b.mommy, b.room, Number(b.payment.amount),
                    hasPayment ? 'Payment' : 'Deposit', b.payment.by, new Date()]);
  }

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
      b.id, created, now, b.mommy, b.room, b.pkg, b.items, Number(b.discount) || 0, total, deposit, paid, balance, status, b.data]]);
  } else {
    recs.appendRow([b.id, now, now, b.mommy, b.room, b.pkg, b.items, Number(b.discount) || 0, total, deposit, paid, balance, status, b.data]);
  }
  return { ok: true, id: b.id, paid, balance, status };
}
