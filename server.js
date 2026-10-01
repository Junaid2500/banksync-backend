var express = require('express');
var multer = require('multer');
var cors = require('cors');
var ExcelJS = require('exceljs');
var PdfReader = require('pdfreader').PdfReader;

var app = express();
app.use(cors());
var upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

// ================= PDF se items =================
function readItems(buffer, password) {
  return new Promise(function(resolve, reject) {
    var items = [];
    var page = 0;
    var done = false;
    var reader = new PdfReader(password ? { password: password } : {});
    reader.parseBuffer(buffer, function(err, item) {
      if (done) return;
      if (err) { done = true; return reject(err); }
      if (!item) { done = true; return resolve(items); }
      if (item.page) page = item.page;
      if (item.text !== undefined && String(item.text).trim() !== '') {
        var tx = String(item.text).trim();
        items.push({ page: page, x: item.x, y: item.y, w: item.w || (tx.length * 0.5), text: tx });
      }
    });
  });
}

// ================= Helpers =================
var DATE_RES = [
  /^\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}$/,
  /^\d{1,2}[\-\/ ][A-Za-z]{3,9}[\-\/ ,']*\d{2,4}$/,
  /^\d{4}-\d{2}-\d{2}$/
];
function isDate(t) {
  for (var i = 0; i < DATE_RES.length; i++) {
    if (DATE_RES[i].test(t)) return true;
  }
  return false;
}

var AMT_RE = /^\(?-?[\d,]*\d\.\d{1,2}\)?\s*(Dr|Cr)?\.?$/i;
function isAmountText(t) {
  if (isDate(t)) return false;
  return AMT_RE.test(t);
}
function parseAmt(t) {
  var m = t.match(/(Dr|Cr)\.?$/i);
  var suffix = m ? m[1].toLowerCase() : '';
  var num = parseFloat(t.replace(/[^0-9.]/g, ''));
  if (isNaN(num)) num = 0;
  return { num: num, suffix: suffix };
}
function round2(n) { return Math.round(n * 100) / 100; }

// "01-May-" aur neeche "2026" alag alag line mein ho to jod do
function stackDates(items) {
  var used = {};
  var p1 = /^\d{1,2}[\-\/ ][A-Za-z]{3,9}[\-\/ ]?$/;
  var p2 = /^\d{1,2}[\/\-.]\d{1,2}[\/\-.]$/;
  for (var i = 0; i < items.length; i++) {
    var a = items[i];
    if (used[i]) continue;
    if (!(p1.test(a.text) || p2.test(a.text))) continue;
    var ac = a.x + a.w / 2;
    for (var j = 0; j < items.length; j++) {
      if (j === i || used[j]) continue;
      var b = items[j];
      if (b.page !== a.page) continue;
      if (!/^\d{2,4}$/.test(b.text)) continue;
      var dy = b.y - a.y;
      if (dy <= 0 || dy > 2.6) continue;
      if (Math.abs((b.x + b.w / 2) - ac) > 2) continue;
      a.text = a.text + (/[A-Za-z]$/.test(a.text) ? '-' : '') + b.text;
      used[j] = true;
      break;
    }
  }
  var out = [];
  for (var k = 0; k < items.length; k++) {
    if (!used[k]) out.push(items[k]);
  }
  return out;
}

function mergeRowItems(list) {
  list.sort(function(a, b) { return a.x - b.x; });
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var it = list[i];
    var last = out[out.length - 1];
    if (last && (it.x - (last.x + last.w)) < 0.45 && it.x >= last.x &&
        !isAmountText(last.text) && !isAmountText(it.text) &&
        !isDate(last.text) && !isDate(it.text)) {
      last.text = last.text + ' ' + it.text;
      last.w = (it.x + it.w) - last.x;
    } else {
      out.push({ page: it.page, x: it.x, y: it.y, w: it.w, text: it.text });
    }
  }
  return out;
}

function buildRows(items) {
  var sorted = items.slice().sort(function(a, b) {
    if (a.page !== b.page) return a.page - b.page;
    return a.y - b.y;
  });
  var rows = [];
  var cur = null;
  for (var i = 0; i < sorted.length; i++) {
    var it = sorted[i];
    if (!cur || cur.page !== it.page || Math.abs(it.y - cur.y) > 0.4) {
      cur = { page: it.page, y: it.y, raw: [] };
      rows.push(cur);
    }
    cur.raw.push(it);
  }
  for (var r = 0; r < rows.length; r++) {
    rows[r].items = mergeRowItems(rows[r].raw);
    var parts = [];
    for (var c = 0; c < rows[r].items.length; c++) parts.push(rows[r].items[c].text);
    rows[r].text = parts.join(' | ');
  }
  return rows;
}

// ================= Header dhundho =================
function findHeader(rows) {
  for (var pass = 0; pass < 2; pass++) {
    for (var i = 0; i < rows.length; i++) {
      var txt = '';
      var hasTxData = false;
      for (var j = i; j < rows.length && rows[j].page === rows[i].page && rows[j].y - rows[i].y <= 2.4; j++) {
        txt += rows[j].text.toUpperCase() + ' ';
      }
      var hasBal = /BALANCE/.test(txt);
      var hasMoney = /DEBIT|WITHDRAW|CREDIT|DEPOSIT|AMOUNT/.test(txt);
      var hasLabel = /DATE|PARTICULARS|NARRATION|DESCRIPTION/.test(txt);
      if (hasBal && hasMoney && (pass === 1 || hasLabel)) {
        var start = i;
        while (start > 0 && rows[start - 1].page === rows[i].page && rows[i].y - rows[start - 1].y <= 1.6) start--;
        var end = i;
        while (end + 1 < rows.length && rows[end + 1].page === rows[i].page && rows[end + 1].y - rows[i].y <= 2.4) end++;
        return { start: start, end: end };
      }
    }
  }
  return null;
}

function getAnchors(rows, hdr) {
  var a = { debit: null, credit: null, balance: null, amount: null, narr: null, chq: null };
  for (var r = hdr.start; r <= hdr.end; r++) {
    for (var i = 0; i < rows[r].items.length; i++) {
      var it = rows[r].items[i];
      var up = it.text.toUpperCase();
      var c = it.x + it.w / 2;
      if (/WITHDRAW|DEBIT/.test(up)) { if (a.debit === null) a.debit = c; }
      else if (/DEPOSIT|CREDIT/.test(up)) { if (a.credit === null) a.credit = c; }
      else if (/BALANCE/.test(up)) { if (a.balance === null) a.balance = c; }
      else if (/^AMOUNT/.test(up)) { if (a.amount === null) a.amount = c; }
      else if (/DESCRIPTION|NARRATION|PARTICULARS|DETAILS|REMARKS/.test(up)) { if (a.narr === null) a.narr = c; }
      else if (/CHEQUE|CHQ/.test(up)) { if (a.chq === null) a.chq = c; }
    }
  }
  return a;
}

function nearestKey(cx, list) {
  var best = null;
  var bd = 1e9;
  for (var i = 0; i < list.length; i++) {
    var d = Math.abs(cx - list[i].x);
    if (d < bd) { bd = d; best = list[i].k; }
  }
  return best;
}

// ================= Transactions banao =================
var STOP_RE = /^(opening balance|closing balance|total|grand total|page\s*\d|statement (summary|generated)|this is a computer|end of statement|\*{3,}|registered office|legends|disclaimer)/i;

function joinNarr(parts) {
  var anySpace = false;
  for (var i = 0; i < parts.length; i++) {
    if (parts[i].indexOf(' ') !== -1) anySpace = true;
  }
  var s = '';
  for (var j = 0; j < parts.length; j++) {
    if (j === 0) { s = parts[j]; continue; }
    if (!anySpace || /[\/\-]$/.test(s)) s += parts[j];
    else s += ' ' + parts[j];
  }
  return s;
}

function findTxDate(row, leftLimit) {
  var before = 0;
  for (var i = 0; i < row.items.length; i++) {
    var it = row.items[i];
    if (isDate(it.text)) {
      if (before <= 3 && (it.x + it.w / 2) < leftLimit) return it;
      return null;
    }
    before++;
  }
  return null;
}

function processRow(tx, row, dateItem, ctx) {
  for (var i = 0; i < row.items.length; i++) {
    var it = row.items[i];
    var t = it.text;
    var cx = it.x + it.w / 2;
    if (dateItem && it === dateItem) { tx.date = t; tx.dateX = it.x; continue; }
    if (isDate(t)) {
      if (!tx.date) tx.date = t;
      else if (!tx.valueDate) tx.valueDate = t;
      continue;
    }
    if (isAmountText(t) && cx >= ctx.leftLimit) {
      var p = parseAmt(t);
      tx.amounts.push({ num: p.num, suffix: p.suffix, x: cx });
      continue;
    }
    if (/^\(?(dr|cr)\)?\.?$/i.test(t)) continue;
    if (dateItem && tx.dateX !== null && it.x < tx.dateX - 0.5) {
      if (/^\d{1,5}$/.test(t)) continue;
      tx.refParts.push(t);
      continue;
    }
    if (/^\d{6,9}$/.test(t) && ctx.anchors.chq !== null && ctx.anchors.narr !== null &&
        Math.abs(cx - ctx.anchors.chq) < Math.abs(cx - ctx.anchors.narr)) {
      tx.chq = t;
      continue;
    }
    tx.narrParts.push(t);
  }
}

function finalizeTx(tx, ctx) {
  if (!tx.amounts.length || !tx.date) return null;
  var narration = joinNarr(tx.narrParts);
  if (/^(opening balance|closing balance|b\/f|c\/f|brought forward|carried forward)/i.test(narration)) return null;

  var o = {
    date: tx.date, valueDate: tx.valueDate, narration: narration,
    refNo: tx.refParts.join(' '), chqNo: tx.chq,
    withdrawal: null, deposit: null, balance: null, amt: null, suffix: ''
  };
  var i;
  if (ctx.money.length) {
    for (i = 0; i < tx.amounts.length; i++) {
      var a = tx.amounts[i];
      var k = nearestKey(a.x, ctx.money);
      if (k === 'debit' && o.withdrawal === null) o.withdrawal = a.num;
      else if (k === 'credit' && o.deposit === null) o.deposit = a.num;
      else if (k === 'balance' && o.balance === null) o.balance = a.num;
      else if (k === 'amount' && o.amt === null) { o.amt = a.num; o.suffix = a.suffix; }
    }
  } else {
    var sorted = tx.amounts.slice().sort(function(p, q) { return p.x - q.x; });
    if (sorted.length === 1) {
      o.amt = sorted[0].num; o.suffix = sorted[0].suffix;
    } else {
      o.balance = sorted[sorted.length - 1].num;
      for (i = 0; i < sorted.length - 1; i++) {
        if (sorted[i].num !== 0) { o.amt = sorted[i].num; o.suffix = sorted[i].suffix; break; }
      }
    }
  }
  if (o.withdrawal === 0 && o.deposit !== null && o.deposit !== 0) o.withdrawal = null;
  if (o.deposit === 0 && o.withdrawal !== null && o.withdrawal !== 0) o.deposit = null;
  if (o.amt !== null) {
    if (o.suffix === 'dr') { o.withdrawal = o.amt; o.amt = null; }
    else if (o.suffix === 'cr') { o.deposit = o.amt; o.amt = null; }
  }
  return o;
}

function moneyOf(o) {
  if (o.withdrawal !== null) return o.withdrawal;
  if (o.deposit !== null) return o.deposit;
  return o.amt;
}

function resolveSides(list) {
  function score(dir) {
    var c = 0;
    for (var i = 0; i < list.length; i++) {
      var p = i - dir;
      if (p < 0 || p >= list.length) continue;
      if (list[i].balance === null || list[p].balance === null) continue;
      var m = moneyOf(list[i]);
      if (m === null) continue;
      if (Math.abs(Math.abs(list[i].balance - list[p].balance) - m) < 0.02) c++;
    }
    return c;
  }
  var dir = score(-1) > score(1) ? -1 : 1;
  for (var i = 0; i < list.length; i++) {
    var o = list[i];
    var m = moneyOf(o);
    var fixed = false;
    var p = i - dir;
    if (m !== null && p >= 0 && p < list.length && o.balance !== null && list[p].balance !== null &&
        !(o.withdrawal !== null && o.deposit !== null)) {
      var d = round2(o.balance - list[p].balance);
      if (Math.abs(Math.abs(d) - m) < 0.02) {
        if (d > 0) { o.deposit = m; o.withdrawal = null; }
        else { o.withdrawal = m; o.deposit = null; }
        o.amt = null;
        fixed = true;
      }
    }
    if (!fixed && o.amt !== null) {
      if (/\bCR\b|CREDIT|SALARY|DEPOSIT|INTEREST|REFUND|REVERS/i.test(o.narration)) o.deposit = o.amt;
      else o.withdrawal = o.amt;
      o.amt = null;
    }
  }
  return list;
}

function categoryOf(n) {
  var u = n.toUpperCase();
  if (/UPI/.test(u)) return 'UPI';
  if (/IMPS/.test(u)) return 'IMPS';
  if (/NEFT/.test(u)) return 'NEFT';
  if (/RTGS/.test(u)) return 'RTGS';
  if (/ATM|CASH WDL|CASH DEP/.test(u)) return 'ATM/Cash';
  if (/CHQ|CHEQUE|CLG|CLEARING/.test(u)) return 'Cheque';
  if (/POS|CARD|ECOM/.test(u)) return 'Card';
  if (/SALARY/.test(u)) return 'Salary';
  if (/INT\.?PD|INTEREST|INT CR/.test(u)) return 'Interest';
  if (/CHARGE|CHRG|FEE|GST|SMS/.test(u)) return 'Charges';
  return 'Other';
}

// ================= Bank info =================
var IFSC_BANK = {
  ICIC: 'ICICI Bank', HDFC: 'HDFC Bank', INDB: 'IndusInd Bank', MAHB: 'Bank of Maharashtra',
  SBIN: 'State Bank of India', UTIB: 'Axis Bank', KKBK: 'Kotak Mahindra Bank', YESB: 'YES Bank',
  PUNB: 'Punjab National Bank', UBIN: 'Union Bank of India', CNRB: 'Canara Bank', BARB: 'Bank of Baroda'
};
var BANK_WORDS = [
  ['ICICI', 'ICICI Bank'], ['HDFC BANK', 'HDFC Bank'], ['INDUSIND', 'IndusInd Bank'],
  ['BANK OF MAHARASHTRA', 'Bank of Maharashtra'], ['STATE BANK OF INDIA', 'State Bank of India'],
  ['AXIS BANK', 'Axis Bank'], ['KOTAK', 'Kotak Mahindra Bank'], ['YES BANK', 'YES Bank'],
  ['PUNJAB NATIONAL', 'Punjab National Bank'], ['UNION BANK', 'Union Bank of India'],
  ['CANARA', 'Canara Bank'], ['BANK OF BARODA', 'Bank of Baroda']
];

function extractInfo(rows, headerStart) {
  var lim = headerStart > 0 ? headerStart : Math.min(rows.length, 40);
  var parts = [];
  for (var i = 0; i < lim && i < rows.length; i++) parts.push(rows[i].text);
  var text = parts.join('\n');
  var info = { bank: '', accountNo: '', period: '', ifsc: '', micr: '' };

  var m = text.match(/\b([A-Z]{4})0[A-Z0-9]{6}\b/);
  if (m) {
    info.ifsc = m[0];
    if (IFSC_BANK[m[1]]) info.bank = IFSC_BANK[m[1]];
  }
  if (!info.bank) {
    var up = text.toUpperCase();
    for (var b = 0; b < BANK_WORDS.length; b++) {
      if (up.indexOf(BANK_WORDS[b][0]) !== -1) { info.bank = BANK_WORDS[b][1]; break; }
    }
  }
  var a = text.match(/(?:A\/C|ACCOUNT)[^0-9X*\n]{0,25}([X*]*\d[X*\d]{5,19})/i);
  if (a) info.accountNo = a[1];
  var mi = text.match(/MICR[^0-9\n]{0,15}(\d{9})/i);
  if (mi) info.micr = mi[1];
  var p = text.match(/(\d{1,2}[\/\-\s][A-Za-z0-9]{2,3}[\/\-\s']*\d{2,4})\s*(?:to|-|TO)\s*(\d{1,2}[\/\-\s][A-Za-z0-9]{2,3}[\/\-\s']*\d{2,4})/);
  if (p) info.period = p[1] + ' to ' + p[2];
  return info;
}

// ================= Poora parse =================
function parseStatement(allItems) {
  var items = stackDates(allItems);
  var rows = buildRows(items);
  var docMaxX = 0;
  for (var i = 0; i < items.length; i++) {
    if (items[i].x + items[i].w > docMaxX) docMaxX = items[i].x + items[i].w;
  }

  var hdr = findHeader(rows);
  var anchors = { debit: null, credit: null, balance: null, amount: null, narr: null, chq: null };
  var startRow = 0;
  if (hdr) {
    anchors = getAnchors(rows, hdr);
    startRow = hdr.end + 1;
  }

  var money = [];
  var keys = ['debit', 'credit', 'balance', 'amount'];
  for (var k = 0; k < keys.length; k++) {
    if (anchors[keys[k]] !== null) money.push({ k: keys[k], x: anchors[keys[k]] });
  }
  var leftLimit = docMaxX * 0.5;
  if (money.length) {
    var xs = [];
    for (var q = 0; q < money.length; q++) xs.push(money[q].x);
    xs.sort(function(a, b) { return a - b; });
    var gap = 6;
    if (xs.length > 1) {
      gap = 1e9;
      for (var g = 1; g < xs.length; g++) {
        if (xs[g] - xs[g - 1] < gap) gap = xs[g] - xs[g - 1];
      }
    }
    leftLimit = xs[0] - Math.max(gap, 2) * 0.9;
  }
  var ctx = { anchors: anchors, money: money, leftLimit: leftLimit };

  var rawTx = [];
  var cur = null;
  var curPage = 0;
  function newTx() {
    return { date: '', valueDate: '', dateX: null, narrParts: [], refParts: [], chq: '', amounts: [], lastY: 0 };
  }
  for (var r = startRow; r < rows.length; r++) {
    var row = rows[r];
    if (row.page !== curPage) {
      if (cur) rawTx.push(cur);
      cur = null;
      curPage = row.page;
    }
    var dItem = findTxDate(row, leftLimit);
    if (dItem) {
      if (cur) rawTx.push(cur);
      cur = newTx();
      processRow(cur, row, dItem, ctx);
      cur.lastY = row.y;
      continue;
    }
    if (!cur) continue;
    if (STOP_RE.test(row.text) || (row.y - cur.lastY) > 3.5) {
      rawTx.push(cur);
      cur = null;
      continue;
    }
    processRow(cur, row, null, ctx);
    cur.lastY = row.y;
  }
  if (cur) rawTx.push(cur);

  var list = [];
  for (var t = 0; t < rawTx.length; t++) {
    var f = finalizeTx(rawTx[t], ctx);
    if (f) list.push(f);
  }
  resolveSides(list);
  for (var u = 0; u < list.length; u++) {
    list[u].category = categoryOf(list[u].narration);
    delete list[u].amt;
    delete list[u].suffix;
  }

  var sample = [];
  for (var s = 0; s < rows.length && s < 60; s++) {
    sample.push('p' + rows[s].page + ' y' + round2(rows[s].y) + ': ' + rows[s].text);
  }
  return {
    info: extractInfo(rows, hdr ? hdr.start : 0),
    transactions: list,
    debug: {
      totalItems: items.length,
      totalRows: rows.length,
      headerFound: !!hdr,
      anchors: anchors,
      leftLimit: round2(leftLimit),
      firstRows: sample
    }
  };
}

async function handleFile(req) {
  if (!req.file) {
    var e0 = new Error('No file uploaded');
    e0.status = 400;
    throw e0;
  }
  var password = (req.body && req.body.password) || '';
  var items;
  try {
    items = await readItems(req.file.buffer, password);
  } catch (e) {
    var msg = String(e && e.message ? e.message : e);
    var e1 = new Error(/password/i.test(msg) ? 'PDF password galat hai ya password chahiye' : 'PDF read nahi hui: ' + msg);
    e1.status = 400;
    throw e1;
  }
  if (!items.length) {
    var e2 = new Error('PDF me text nahi mila (scanned image PDF support nahi hai)');
    e2.status = 400;
    throw e2;
  }
  return parseStatement(items);
}

// ================= Excel =================
async function buildExcel(info, txs) {
  var wb = new ExcelJS.Workbook();
  var ws = wb.addWorksheet('Statement');
  var GREEN = 'FF1B8A5A';
  var LIGHT = 'FFEAF6F0';

  ws.columns = [
    { width: 7 }, { width: 13 }, { width: 50 }, { width: 13 }, { width: 18 },
    { width: 12 }, { width: 13 }, { width: 16 }, { width: 16 }, { width: 16 }
  ];

  ws.mergeCells('A1:J1');
  var title = ws.getCell('A1');
  title.value = 'ACCOUNT STATEMENT';
  title.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
  title.alignment = { horizontal: 'center', vertical: 'middle' };
  title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } };
  ws.getRow(1).height = 28;

  var infoRows = [
    ['Bank', info.bank || ''],
    ['Account', info.accountNo || ''],
    ['Period', info.period || ''],
    ['IFSC', info.ifsc || ''],
    ['MICR', info.micr || '']
  ];
  for (var i = 0; i < infoRows.length; i++) {
    var rn = i + 2;
    ws.getCell('A' + rn).value = infoRows[i][0];
    ws.getCell('A' + rn).font = { bold: true };
    ws.mergeCells('A' + rn + ':B' + rn);
    ws.getCell('C' + rn).value = infoRows[i][1];
    ws.mergeCells('C' + rn + ':F' + rn);
  }

  var headerRowNo = 8;
  var heads = ['S.No', 'Date', 'Narration', 'Category', 'Ref No.', 'Chq No.', 'Value Date', 'Withdrawal (₹)', 'Deposit (₹)', 'Balance (₹)'];
  var hr = ws.getRow(headerRowNo);
  for (var h = 0; h < heads.length; h++) {
    var cell = hr.getCell(h + 1);
    cell.value = heads[h];
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  }
  hr.height = 22;

  for (var k = 0; k < txs.length; k++) {
    var t = txs[k];
    var row = ws.getRow(headerRowNo + 1 + k);
    row.getCell(1).value = k + 1;
    row.getCell(2).value = t.date;
    row.getCell(3).value = t.narration;
    row.getCell(4).value = t.category;
    row.getCell(5).value = t.refNo;
    row.getCell(6).value = t.chqNo;
    row.getCell(7).value = t.valueDate || t.date;
    row.getCell(8).value = t.withdrawal;
    row.getCell(9).value = t.deposit;
    row.getCell(10).value = t.balance;
    row.getCell(3).alignment = { wrapText: true, vertical: 'top' };
    for (var c = 8; c <= 10; c++) row.getCell(c).numFmt = '#,##0.00';
    if (k % 2 === 1) {
      for (var cc = 1; cc <= 10; cc++) {
        row.getCell(cc).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT } };
      }
    }
  }
  ws.views = [{ state: 'frozen', ySplit: headerRowNo }];
  return await wb.xlsx.writeBuffer();
}

// ================= Routes =================
app.get('/', function(req, res) {
  res.send('BankSync Pro Backend Running! (no-AI parser) - test page: /test');
});

app.post('/parse', upload.single('pdf'), async function(req, res) {
  try {
    var result = await handleFile(req);
    res.json(result);
  } catch (e) {
    res.status(e.status || 500).json({ error: String(e.message || e) });
  }
});

app.post('/convert', upload.single('pdf'), async function(req, res) {
  try {
    var result = await handleFile(req);
    if (!result.transactions.length) {
      return res.status(400).json({ error: 'Transactions nahi mile. /test page pe debug dekho.' });
    }
    var buf = await buildExcel(result.info, result.transactions);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="statement.xlsx"');
    res.send(Buffer.from(buf));
  } catch (e) {
    console.log('Convert error:', e);
    res.status(e.status || 500).json({ error: String(e.message || e) });
  }
});

var TEST_PAGE = [
'<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">',
'<title>BankSync Test</title>',
'<style>body{font-family:Segoe UI,sans-serif;background:#0d1117;color:#e6edf3;margin:0;padding:20px}',
'.box{max-width:1100px;margin:0 auto}h1{color:#00d084}',
'input,button{padding:10px;border-radius:8px;border:1px solid #30363d;background:#161b22;color:#e6edf3;margin:4px 0}',
'button{background:#00a86b;border:0;cursor:pointer;font-weight:700}',
'table{border-collapse:collapse;width:100%;font-size:12px;margin-top:12px}',
'th{background:#1b8a5a;padding:6px;text-align:left;position:sticky;top:0}',
'td{padding:5px 6px;border-bottom:1px solid #30363d}',
'.w{max-height:420px;overflow:auto;border:1px solid #30363d;border-radius:8px}',
'pre{background:#161b22;padding:12px;border-radius:8px;overflow:auto;font-size:11px;max-height:300px}',
'.r{color:#ff6b6b}.g{color:#00d084}</style></head><body><div class="box">',
'<h1>BankSync - Test Page</h1>',
'<div><input type="file" id="f" accept=".pdf"><br>',
'<input type="text" id="p" placeholder="PDF password (agar ho)"><br>',
'<button onclick="runParse()">PARSE KARO</button> ',
'<button onclick="runExcel()">EXCEL DOWNLOAD</button></div>',
'<p id="s"></p><div class="w"><table><thead><tr><th>#</th><th>Date</th><th>Narration</th><th>Ref</th><th>Withdrawal</th><th>Deposit</th><th>Balance</th></tr></thead><tbody id="b"></tbody></table></div>',
'<h3>Debug (problem aaye to ye copy karke bhejo)</h3><pre id="d"></pre>',
'<script>',
'var BASE = "";',
'function getForm(){var f=document.getElementById("f").files[0];if(!f){alert("PDF chuno");return null;}',
'var fd=new FormData();fd.append("pdf",f);var pw=document.getElementById("p").value;if(pw)fd.append("password",pw);return fd;}',
'function cell(tr,txt,cls){var td=document.createElement("td");td.textContent=(txt===null||txt===undefined)?"":txt;if(cls)td.className=cls;tr.appendChild(td);}',
'function runParse(){var fd=getForm();if(!fd)return;var s=document.getElementById("s");s.textContent="Processing... (pehli baar 30-60 sec lag sakte hain)";',
'fetch(BASE+"/parse",{method:"POST",body:fd}).then(function(r){return r.json();}).then(function(j){',
'if(j.error){s.textContent="Error: "+j.error;return;}',
'var tb=document.getElementById("b");tb.innerHTML="";var t=j.transactions;',
'for(var i=0;i<t.length;i++){var tr=document.createElement("tr");cell(tr,i+1);cell(tr,t[i].date);cell(tr,t[i].narration);cell(tr,t[i].refNo);cell(tr,t[i].withdrawal,"r");cell(tr,t[i].deposit,"g");cell(tr,t[i].balance);tb.appendChild(tr);}',
's.textContent="Bank: "+(j.info.bank||"?")+" | Transactions: "+t.length+" | Header mila: "+j.debug.headerFound;',
'document.getElementById("d").textContent=JSON.stringify({info:j.info,debug:j.debug},null,2);',
'}).catch(function(e){s.textContent="Server error: "+e;});}',
'function runExcel(){var fd=getForm();if(!fd)return;var s=document.getElementById("s");s.textContent="Excel ban raha hai...";',
'fetch(BASE+"/convert",{method:"POST",body:fd}).then(function(r){if(!r.ok){return r.json().then(function(j){throw new Error(j.error);});}return r.blob();}).then(function(bl){',
'var a=document.createElement("a");a.href=URL.createObjectURL(bl);a.download="statement.xlsx";a.click();s.textContent="Excel download ho gaya";',
'}).catch(function(e){s.textContent="Error: "+e.message;});}',
'</script></div></body></html>'
].join('\n');

app.get('/test', function(req, res) {
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.send(TEST_PAGE);
});

var PORT = process.env.PORT || 3000;
app.listen(PORT, function() {
  console.log('BankSync Pro running on port ' + PORT);
});
