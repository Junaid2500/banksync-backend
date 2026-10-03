var express = require('express');
var multer = require('multer');
var cors = require('cors');
var ExcelJS = require('exceljs');
var PdfReader = require('pdfreader').PdfReader;

var app = express();
app.use(cors());
var upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

var MONEY_RE = /^\(?-?₹?\s?\d[\d,]*\.\d{2}\)?\s?(?:dr|cr)?\.?$/i;
var DATE_RE = /^(?:\d{1,2}[-\/. ](?:\d{1,2}|[A-Za-z]{3,9})[-\/. ,']*\d{2,4}|\d{4}-\d{2}-\d{2})$/;
var STACK_RE = /^\d{1,2}[-\/ ][A-Za-z]{3,9}[-\/ ]?$/;  // 01-Apr-
var STACK_YMD_RE = /^\d{4}-\d{2}-$/;                 // 2026-02-
var HDR_RE = /\b(s\.?\s?no|sl\.?\s?no|transaction|txn|date|cheque|chq|description|narration|particulars|details|withdrawals?|deposits?|debit|credit|balance|available|amount|ref|dr|cr)\b/gi;

var STOP_RE = /(generated on|page\s+\d+\s+of|closing balance|opening balance|statement summary|statement of account|total\s+(debit|credit|withdrawal|deposit)|grand total|end of statement|computer generated|system generated|registered office|regd\.?\s*office|corporate identity number|cin\s*[:\-]|toll[- ]?free|customer care|grievance|disclaimer|e\.?\s?&\.?\s?o\.?\s?e\.?|terms and condition|subject to realisation|this is a (system|computer)|swift\s*code|www\.[a-z]|https?:\/\/|member.*deposit insurance|all disputes|jurisdiction|for any quer(y|ies)|legends?\s*:)/i;

// spaces hata ke footer detect (HDFC me words chipke hote hain)
var STOP_COMPACT_RE = /(statementsummary|generatedon|generatedby|requestingbranchcode|thisisacomputergeneratedstatement|doesnotrequiresignature|hdfcbanklimited|closingbalanceincludes|contentsofthisstatement|stateaccountbranchgstn|registeredofficeaddress)/i;

var OPEN_RE = /^(brought\s*forward|balance\s*brought\s*forward|opening\s*balance|op\.?\s*bal\.?|b\/f\b|carried\s*forward|balance\s*carried\s*forward|c\/f\b)/i;

function r2(x) { return Math.round(x * 100) / 100; }

function num(t) {
  var s = String(t).replace(/[₹,\s]/g, '').replace(/(dr|cr)\.?$/i, '');
  var neg = /^\(.*\)$/.test(s) || s.charAt(0) === '-';
  s = s.replace(/[()\-]/g, '');
  var v = parseFloat(s);
  if (isNaN(v)) v = 0;
  return neg ? -v : v;
}

function balOf(t) {
  var v = num(t);
  if (/dr\.?$/i.test(String(t).trim())) v = -Math.abs(v);
  return v;
}

function isStopRow(text) {
  if (!text) return false;
  if (STOP_RE.test(text)) return true;
  var compact = String(text).toLowerCase().replace(/\s+/g, '');
  if (STOP_COMPACT_RE.test(compact)) return true;
  return false;
}

/* ---------- PDF read ---------- */
function readItems(buffer, password) {
  return new Promise(function (resolve, reject) {
    var items = [];
    var page = 0;
    var finished = false;
    var opts = password ? { password: password } : {};
    try {
      var reader = new PdfReader(opts);
      reader.parseBuffer(buffer, function (err, item) {
        if (finished) return;
        if (err) { finished = true; reject(err); return; }
        if (!item) { finished = true; resolve(items); return; }
        if (item.page) { page = item.page; return; }
        if (item.text !== undefined && String(item.text).trim() !== '') {
          items.push({ page: page, x: item.x, y: item.y, w: item.w || 0, text: String(item.text).trim() });
        }
      });
    } catch (e) {
      if (!finished) { finished = true; reject(e); }
    }
  });
}

/* date split join:
   01-Apr- + 2026  => 01-Apr-2026
   2026-02- + 14   => 2026-02-14
*/
function stackDates(items) {
  var used = {};
  var i, j;

  function closeX(a, b) {
    var ac = a.x + a.w / 2;
    var bc = b.x + b.w / 2;
    var dx1 = Math.abs(bc - ac);
    var dx2 = Math.abs(b.x - a.x);
    return Math.min(dx1, dx2) <= 1.5;
  }

  for (i = 0; i < items.length; i++) {
    var a = items[i];

    if (STACK_RE.test(a.text)) {
      for (j = 0; j < items.length; j++) {
        var b1 = items[j];
        if (used[j] || b1.page !== a.page) continue;
        if (!/^\d{4}$/.test(b1.text)) continue;
        var dy1 = b1.y - a.y;
        if (dy1 < 0.2 || dy1 > 1.6) continue;
        if (!closeX(a, b1)) continue;
        a.text = a.text + b1.text;
        used[j] = true;
        break;
      }
      continue;
    }

    if (STACK_YMD_RE.test(a.text)) {
      for (j = 0; j < items.length; j++) {
        var b2 = items[j];
        if (used[j] || b2.page !== a.page) continue;
        if (!/^\d{1,2}$/.test(b2.text)) continue;
        var dy2 = b2.y - a.y;
        if (dy2 < 0.2 || dy2 > 1.8) continue;
        if (!closeX(a, b2)) continue;
        var dd = b2.text.length === 1 ? ('0' + b2.text) : b2.text;
        a.text = a.text + dd;
        used[j] = true;
        break;
      }
    }
  }

  return items.filter(function (it, idx) { return !used[idx]; });
}

function mergeClose(list) {
  var out = [];
  list.forEach(function (it) {
    var prev = out[out.length - 1];
    if (prev) {
      var gap = it.x - (prev.x + prev.w);

      var special =
        MONEY_RE.test(prev.text) || MONEY_RE.test(it.text) ||
        DATE_RE.test(prev.text) || DATE_RE.test(it.text) ||
        STACK_RE.test(prev.text) || STACK_RE.test(it.text) ||
        STACK_YMD_RE.test(prev.text) || STACK_YMD_RE.test(it.text);

      if (!special && gap < 0.25) {
        prev.text += (gap < 0.05 ? '' : ' ') + it.text;
        prev.w = (it.x + it.w) - prev.x;
        return;
      }
    }
    out.push({ page: it.page, x: it.x, y: it.y, w: it.w, text: it.text });
  });
  return out;
}

function buildRows(items) {
  var pages = {};
  items.forEach(function (it) {
    if (!pages[it.page]) pages[it.page] = [];
    pages[it.page].push(it);
  });

  var rows = [];
  Object.keys(pages).map(Number).sort(function (a, b) { return a - b; }).forEach(function (p) {
    var list = pages[p].slice().sort(function (a, b) { return (a.y - b.y) || (a.x - b.x); });
    var cur = null;
    list.forEach(function (it) {
      if (!cur || Math.abs(it.y - cur.y) > 0.4) {
        cur = { page: p, y: it.y, items: [] };
        rows.push(cur);
      }
      cur.items.push(it);
    });
  });

  rows.forEach(function (r) {
    r.items.sort(function (a, b) { return a.x - b.x; });
    r.items = mergeClose(r.items);
    r.text = r.items.map(function (i) { return i.text; }).join(' ');
  });
  return rows;
}

function computePitch(rows) {
  var gaps = [];
  var i;
  for (i = 1; i < rows.length; i++) {
    if (rows[i].page === rows[i - 1].page) {
      var g = rows[i].y - rows[i - 1].y;
      if (g > 0.05 && g < 3) gaps.push(g);
    }
  }
  if (!gaps.length) return 0.75;
  gaps.sort(function (a, b) { return a - b; });
  return gaps[Math.floor(gaps.length / 2)];
}

/* row mein date kahan hai (pehle 5 items mein, paise se pehle) */
function dateIdx(row, dateX) {
  var i;
  for (i = 0; i < row.items.length && i < 5; i++) {
    var it = row.items[i];
    if (MONEY_RE.test(it.text)) return -1;
    if (DATE_RE.test(it.text)) {
      if (dateX === null || Math.abs(it.x - dateX) <= 1.0) return i;
    }
  }
  return -1;
}

function pickDateX(c) {
  if (!c.length) return null;
  var counts = [];
  var best = 0;
  var i, j;
  for (i = 0; i < c.length; i++) {
    var n = 0;
    for (j = 0; j < c.length; j++) { if (Math.abs(c[j] - c[i]) <= 1.0) n++; }
    counts.push(n);
    if (n > best) best = n;
  }
  var m = null;
  for (i = 0; i < c.length; i++) {
    if (counts[i] >= best * 0.5 && (m === null || c[i] < m)) m = c[i];
  }
  return m;
}

function isHeaderRow(text) {
  if (/\d[\d,]*\.\d{2}/.test(text)) return false;
  var found = {};
  var m = text.match(HDR_RE) || [];
  m.forEach(function (w) { found[w.toLowerCase().replace(/\s+/g, '')] = 1; });
  return Object.keys(found).length >= 3;
}

function joinParts(parts) {
  var s = '';
  parts.forEach(function (p) {
    if (!s) { s = p; return; }
    if (/[-\/]$/.test(s) || /^[-\/]/.test(p)) s += p; else s += ' ' + p;
  });
  return s.replace(/\s+/g, ' ').trim();
}

function pickRef(left) {
  var words = left.join(' ').split(/\s+/);
  var i;
  for (i = 0; i < words.length; i++) {
    if (words[i].length >= 5 && /\d/.test(words[i])) return words[i];
  }
  return '';
}

function categoryOf(s) {
  var u = String(s).toUpperCase();
  if (/UPI/.test(u)) return 'UPI';
  if (/IMPS/.test(u)) return 'IMPS';
  if (/NEFT/.test(u)) return 'NEFT';
  if (/RTGS/.test(u)) return 'RTGS';
  if (/ATM|CASH/.test(u)) return 'ATM/Cash';
  if (/CHQ|CHEQUE|CLG/.test(u)) return 'Cheque';
  if (/POS|CARD/.test(u)) return 'Card';
  if (/SALARY/.test(u)) return 'Salary';
  if (/INTEREST|INT\.?\s?PD/.test(u)) return 'Interest';
  if (/CHARGES|CHRG|FEE|GST|SMS/.test(u)) return 'Charges';
  return 'Other';
}

function guessSide(narr) {
  var u = String(narr).toUpperCase();
  if (/INF\/|BY CASH|BY TRANSFER|BY CLG/.test(u)) return 'd';
  if (/(^|[^A-Z])(CR|CREDIT|DEPOSIT|REFUND|REVERSAL|SALARY|INTEREST|CASHBACK|RECEIVED)([^A-Z]|$)/.test(u)) return 'd';
  return 'w';
}

function resolveSides(recs, dbg) {
  var asc = 0, desc = 0, i;
  for (i = 0; i < recs.length; i++) {
    var t = recs[i];
    if (t.amt === null || t.bal === null) continue;
    var p = recs[i - 1], q = recs[i + 1];
    if (p && p.bal !== null && Math.abs(Math.abs(t.bal - p.bal) - t.amt) < 0.02) asc++;
    if (q && q.bal !== null && Math.abs(Math.abs(t.bal - q.bal) - t.amt) < 0.02) desc++;
  }
  var dir = asc >= desc ? 1 : -1;
  var how = { delta: 0, suffix: 0, position: 0, keyword: 0 };
  dbg.order = dir === 1 ? 'purani-se-nayi' : 'nayi-se-purani';
  dbg.deltaMatches = { asc: asc, desc: desc };

  var side = [];
  for (i = 0; i < recs.length; i++) {
    var r = recs[i];
    side[i] = null;
    if (r.amt === null) continue;
    var e = recs[i - dir];
    if (e && e.bal !== null && r.bal !== null) {
      var diff = r.bal - e.bal;
      if (Math.abs(diff - r.amt) < 0.02) { side[i] = 'd'; how.delta++; continue; }
      if (Math.abs(diff + r.amt) < 0.02) { side[i] = 'w'; how.delta++; continue; }
    }
    if (r.hint) { side[i] = r.hint; how.suffix++; }
  }

  var sw = 0, nw = 0, sd = 0, nd = 0;
  for (i = 0; i < recs.length; i++) {
    if (side[i] && recs[i].tok) {
      var xr = recs[i].tok.x + recs[i].tok.w;
      if (side[i] === 'w') { sw += xr; nw++; } else { sd += xr; nd++; }
    }
  }
  var avgW = nw ? sw / nw : null;
  var avgD = nd ? sd / nd : null;
  dbg.avgWithdrawalX = avgW === null ? null : r2(avgW);
  dbg.avgDepositX = avgD === null ? null : r2(avgD);

  for (i = 0; i < recs.length; i++) {
    var rr = recs[i];
    if (rr.amt === null) continue;
    if (!side[i]) {
      if (avgW !== null && avgD !== null && Math.abs(avgW - avgD) > 1 && rr.tok) {
        var xx = rr.tok.x + rr.tok.w;
        side[i] = Math.abs(xx - avgW) <= Math.abs(xx - avgD) ? 'w' : 'd';
        how.position++;
      } else {
        side[i] = guessSide(rr.narr);
        how.keyword++;
      }
    }
    if (side[i] === 'd') { rr.d = rr.amt; rr.w = 0; } else { rr.w = rr.amt; rr.d = 0; }
  }
  dbg.sideBy = how;
}

var IFSC_BANK = {
  ICIC: 'ICICI Bank', HDFC: 'HDFC Bank', INDB: 'IndusInd Bank', MAHB: 'Bank of Maharashtra',
  SBIN: 'State Bank of India', UTIB: 'Axis Bank', KKBK: 'Kotak Mahindra Bank', YESB: 'YES Bank',
  PUNB: 'Punjab National Bank', UBIN: 'Union Bank of India', CNRB: 'Canara Bank', BARB: 'Bank of Baroda'
};

function extractInfo(rows, firstStart, txs) {
  var headRows = firstStart > 0 ? rows.slice(0, firstStart) : rows.slice(0, 80);
  var head = headRows.map(function (r) { return r.text; }).join('\n');
  var info = { bank: 'Unknown', accountNo: '', period: '', ifsc: '', micr: '' };

  // IFSC (word boundary mat rakho, HDFC me chipka hota hai)
  var m = head.match(/([A-Z]{4}0[A-Z0-9]{6})/);
  if (m) {
    info.ifsc = m[1];
    if (IFSC_BANK[m[1].substring(0, 4)]) info.bank = IFSC_BANK[m[1].substring(0, 4)];
  }

  if (info.bank === 'Unknown') {
    var names = [
      [/bank of maharashtra/i, 'Bank of Maharashtra'], [/indusind/i, 'IndusInd Bank'], [/icici/i, 'ICICI Bank'],
      [/hdfc/i, 'HDFC Bank'], [/state bank/i, 'State Bank of India'], [/axis bank/i, 'Axis Bank'],
      [/kotak/i, 'Kotak Mahindra Bank'], [/yes bank/i, 'YES Bank'], [/punjab national/i, 'Punjab National Bank'],
      [/union bank/i, 'Union Bank of India'], [/canara/i, 'Canara Bank'], [/bank of baroda/i, 'Bank of Baroda']
    ];
    var i;
    for (i = 0; i < names.length; i++) {
      if (names[i][0].test(head)) { info.bank = names[i][1]; break; }
    }
  }

  m = head.match(/(?:a\/c|account)\s*(?:number|no\.?|#)?\s*[:\-]?\s*(\d{9,18})/i);
  if (m) info.accountNo = m[1];
  else {
    m = head.match(/([xX*]{3,}\d{3,6})/);
    if (m) info.accountNo = m[1];
    else { m = head.match(/\b(\d{11,16})\b/); if (m) info.accountNo = m[1]; }
  }

  m = head.match(/MICR\s*(?:code|no\.?)?\s*[:\-]?\s*(\d{9})/i);
  if (m) info.micr = m[1];

  var D = "\\d{1,2}[-\\/. ]+[A-Za-z0-9]{2,9}[-\\/. ,]*'?\\d{2,4}";
  var pr = new RegExp('(' + D + ')\\s*(?:to|-|–|—|till|through)\\s*(' + D + ')', 'i');
  m = head.match(pr);
  if (m) info.period = m[1] + ' to ' + m[2];
  else if (txs.length) info.period = txs[0].date + ' to ' + txs[txs.length - 1].date;

  return info;
}

function analyze(rawItems) {
  var items = stackDates(rawItems);
  var rows = buildRows(items);
  var dbg = { totalItems: items.length, totalRows: rows.length };

  var pitch = computePitch(rows);
  var maxGap = Math.min(Math.max(pitch * 4.5, 1.8), 6);
  dbg.linePitch = r2(pitch);
  dbg.maxRowGap = r2(maxGap);

  var cands = [];
  rows.forEach(function (r) {
    var di = dateIdx(r, null);
    if (di >= 0) cands.push(r.items[di].x);
  });
  var dateX = pickDateX(cands);
  dbg.dateX = dateX === null ? null : r2(dateX);

  var firstStart = -1;
  var lefts = [];
  if (dateX !== null) {
    rows.forEach(function (r, idx) {
      if (dateIdx(r, dateX) < 0) return;
      if (firstStart < 0) firstStart = idx;
      var mx = null;
      r.items.forEach(function (it) {
        if (MONEY_RE.test(it.text) && (mx === null || it.x < mx)) mx = it.x;
      });
      if (mx !== null) lefts.push(mx);
    });
  }

  var moneyMinX = null;
  if (lefts.length) {
    lefts.sort(function (a, b) { return a - b; });
    moneyMinX = lefts[Math.floor(lefts.length * 0.05)] - 0.5;
  }
  dbg.moneyMinX = moneyMinX === null ? null : r2(moneyMinX);

  var fr = [];
  rows.slice(0, 60).forEach(function (r) { fr.push('p' + r.page + ' y' + r2(r.y) + ': ' + r.text); });
  dbg.firstRows = fr;

  if (dateX === null || moneyMinX === null) {
    dbg.problem = 'Date column ya amount column nahi mila';
    return { info: extractInfo(rows, firstStart, []), txs: [], debug: dbg };
  }

  var raw = [];
  var cur = null;
  var lastRow = null;
  var startCount = 0;
  var sample = [];

  function addItem(tx, it) {
    if (it.x >= moneyMinX) {
      if (MONEY_RE.test(it.text)) tx.money.push(it);
      else if (!/^(dr|cr)\.?$/i.test(it.text)) tx.extras.push(it.text);
    } else {
      tx.parts.push(it.text);
    }
  }

  rows.forEach(function (row) {
    var di = dateIdx(row, dateX);

    if (di >= 0) {
      startCount++;
      cur = { date: row.items[di].text, valueDate: '', chq: '', left: [], parts: [], extras: [], money: [] };
      raw.push(cur);
      lastRow = row;

      var k;
      for (k = 0; k < di; k++) cur.left.push(row.items[k].text);

      var rest = row.items.slice(di + 1);
      var p = 0;
      if (rest[p] && DATE_RE.test(rest[p].text) && rest[p].x < moneyMinX) { cur.valueDate = rest[p].text; p++; }
      if (rest[p] && /^\d{6,12}$/.test(rest[p].text) && rest[p].x < moneyMinX) { cur.chq = rest[p].text; p++; }
      for (k = p; k < rest.length; k++) addItem(cur, rest[k]);

      if (sample.length < 4) {
        sample.push(row.items.map(function (it) {
          return it.text + '@' + r2(it.x) + (it.w ? '+' + r2(it.w) : '');
        }).join(' | '));
      }
      return;
    }

    if (!cur) return;

    if (row.page !== lastRow.page || (row.y - lastRow.y) > maxGap || isStopRow(row.text) || isHeaderRow(row.text)) {
      cur = null;
      return;
    }

    row.items.forEach(function (it) {
      if (it.x < dateX - 0.5) cur.left.push(it.text);
      else addItem(cur, it);
    });
    lastRow = row;
  });

  dbg.startRows = startCount;
  dbg.sampleStartRows = sample;

  var recs = [];
  var skippedOpening = 0;

  raw.forEach(function (t) {
    var n = t.money.length;
    if (!n) return;

    var narrText = joinParts(t.parts.concat(t.extras));
    if (OPEN_RE.test(narrText.trim())) { skippedOpening++; return; }

    var rec = {
      date: t.date, valueDate: t.valueDate, chq: t.chq, ref: pickRef(t.left),
      narr: narrText,
      amt: null, bal: null, w: null, d: null, tok: null, hint: null
    };

    var m = t.money;
    if (n >= 3) {
      rec.w = Math.abs(num(m[n - 3].text));
      rec.d = Math.abs(num(m[n - 2].text));
      rec.bal = balOf(m[n - 1].text);
    } else if (n === 2) {
      rec.amt = Math.abs(num(m[0].text));
      rec.tok = m[0];
      rec.bal = balOf(m[1].text);
    } else {
      rec.amt = Math.abs(num(m[0].text));
      rec.tok = m[0];
    }

    if (rec.tok) {
      if (/cr\.?$/i.test(rec.tok.text)) rec.hint = 'd';
      else if (/dr\.?$/i.test(rec.tok.text)) rec.hint = 'w';
    }

    recs.push(rec);
  });

  dbg.skippedOpeningRows = skippedOpening;

  resolveSides(recs, dbg);

  var txs = recs.map(function (r) {
    return {
      date: r.date,
      narration: r.narr,
      category: categoryOf(r.narr),
      ref: r.ref,
      chq: r.chq,
      valueDate: r.valueDate,
      withdrawal: r.w ? r2(r.w) : null,
      deposit: r.d ? r2(r.d) : null,
      balance: r.bal === null ? null : r2(r.bal)
    };
  });

  dbg.txCount = txs.length;
  return { info: extractInfo(rows, firstStart, txs), txs: txs, debug: dbg };
}

/* ---------- Excel ---------- */
function solid(argb) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: argb } }; }

function buildExcel(info, txs) {
  var GREEN = 'FF1B8A5A';
  var wb = new ExcelJS.Workbook();
  var ws = wb.addWorksheet('Statement', { views: [{ state: 'frozen', ySplit: 8 }] });
  ws.columns = [{ width: 7 }, { width: 14 }, { width: 60 }, { width: 14 }, { width: 20 }, { width: 12 }, { width: 14 }, { width: 16 }, { width: 16 }, { width: 18 }];

  ws.mergeCells('A1:J1');
  var t = ws.getCell('A1');
  t.value = 'ACCOUNT STATEMENT';
  t.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
  t.fill = solid(GREEN);
  t.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 28;

  var meta = [['Bank', info.bank], ['Account', info.accountNo], ['Period', info.period], ['IFSC', info.ifsc], ['MICR', info.micr]];
  meta.forEach(function (m, i) {
    var r = i + 2;
    ws.mergeCells('A' + r + ':B' + r);
    ws.mergeCells('C' + r + ':F' + r);
    ws.getCell('A' + r).value = m[0];
    ws.getCell('A' + r).font = { bold: true };
    ws.getCell('C' + r).value = m[1] ? String(m[1]) : '';
    ws.getCell('C' + r).alignment = { horizontal: 'left' };
  });

  var heads = ['S.No', 'Date', 'Narration', 'Category', 'Ref No.', 'Chq No.', 'Value Date', 'Withdrawal (₹)', 'Deposit (₹)', 'Balance (₹)'];
  var hr = ws.getRow(8);
  heads.forEach(function (h, c) {
    var cell = hr.getCell(c + 1);
    cell.value = h;
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = solid(GREEN);
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  hr.height = 22;

  txs.forEach(function (tx, i) {
    var row = ws.getRow(9 + i);
    var vals = [i + 1, tx.date, tx.narration, tx.category, tx.ref, tx.chq, tx.valueDate, tx.withdrawal, tx.deposit, tx.balance];
    var c;
    for (c = 0; c < 10; c++) {
      var cell = row.getCell(c + 1);
      cell.value = vals[c] === null || vals[c] === '' ? null : vals[c];
      if (i % 2 === 1) cell.fill = solid('FFEAF6F0');
      cell.alignment = { vertical: 'top', wrapText: c === 2 };
      if (c >= 7) cell.numFmt = '#,##0.00';
    }
  });

  return wb.xlsx.writeBuffer();
}

/* ---------- Routes ---------- */
function errMsg(e) {
  var s = '';
  if (e && e.parserError) s = e.parserError.message ? e.parserError.message : String(e.parserError);
  else if (e && e.message) s = e.message;
  else if (typeof e === 'string') s = e;
  else { try { s = JSON.stringify(e); } catch (x) { s = ''; } }

  if (/password/i.test(s)) return 'PDF password galat hai ya password chahiye';
  if (/compression|flate stream|bad xref|invalid pdf structure|unexpected end of file|fcheck/i.test(s)) {
    return 'Ye PDF file ka format non-standard/corrupt hai (shayad kisi unlock-tool se banayi gayi). PDF ko Chrome mein kholkar Print > Save as PDF karke naya file banao, phir usi ko upload karo.';
  }
  if (!s || s === '{}' || s === '[object Object]') {
    return 'PDF padhi nahi ja saki (password-protected, scan ki hui image ya damaged ho sakti hai)';
  }
  return 'PDF error: ' + s;
}

function run(req) {
  if (!req.file) return Promise.reject(new Error('PDF file nahi mili'));
  return readItems(req.file.buffer, req.body && req.body.password).then(analyze);
}

app.get('/', function (req, res) { res.send('BankSync Pro backend chal raha hai'); });

app.post('/parse', upload.single('pdf'), function (req, res) {
  run(req).then(function (r) {
    res.json({ info: r.info, transactions: r.txs, debug: r.debug });
  }).catch(function (e) {
    console.log('parse error:', e);
    res.status(400).json({ error: errMsg(e) });
  });
});

app.post('/convert', upload.single('pdf'), function (req, res) {
  run(req).then(function (r) {
    if (!r.txs.length) {
      res.status(422).json({ error: 'No transactions found' });
      return null;
    }
    return buildExcel(r.info, r.txs).then(function (buf) {
      res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.set('Content-Disposition', 'attachment; filename="BankSync_Statement.xlsx"');
      res.send(Buffer.from(buf));
    });
  }).catch(function (e) {
    console.log('convert error:', e);
    res.status(400).json({ error: errMsg(e) });
  });
});

var TEST_PAGE = '<!DOCTYPE html><html><head><meta charset="utf-8"><title>BankSync Test</title>' +
'<style>body{font-family:Arial,sans-serif;background:#0d1117;color:#e6edf3;margin:0;padding:20px}' +
'h2{color:#2ecc8f}input,button{padding:10px;margin:4px;border-radius:6px;border:1px solid #30363d;background:#161b22;color:#e6edf3}' +
'button{background:#1b8a5a;border:0;cursor:pointer;font-weight:bold}table{border-collapse:collapse;width:100%;font-size:12px;margin-top:12px}' +
'th{background:#1b8a5a;color:#fff;padding:6px;position:sticky;top:0}td{border-bottom:1px solid #30363d;padding:5px;vertical-align:top}' +
'tr:nth-child(even) td{background:#161b22}.num{text-align:right;white-space:nowrap}' +
'textarea{width:100%;height:260px;background:#010409;color:#9ef0c4;border:1px solid #30363d;font-size:11px}#wrap{max-height:520px;overflow:auto}</style>' +
'</head><body><h2>BankSync Pro - Test Page</h2><div>' +
'<input type="file" id="f" accept="application/pdf">' +
'<input type="password" id="pw" placeholder="PDF password (agar ho)">' +
'<button onclick="parseIt()">PARSE KARO</button>' +
'<button onclick="dl()">EXCEL DOWNLOAD</button></div>' +
'<div id="st" style="margin:8px 4px"></div><div id="info" style="margin:4px"></div><div id="wrap"></div>' +
'<h3>Debug box</h3><textarea id="dbg" readonly></textarea>' +
'<script>' +
'function esc(s){return String(s===null||s===undefined?\"\":s).replace(/&/g,\"&amp;\").replace(/</g,\"&lt;\");}' +
'function fmt(v){return v===null||v===undefined?\"\":Number(v).toLocaleString(\"en-IN\",{minimumFractionDigits:2});}' +
'function post(path,type,done){var f=document.getElementById(\"f\").files[0];if(!f){alert(\"Pehle PDF choose karo\");return;}' +
'var fd=new FormData();fd.append(\"pdf\",f);var pw=document.getElementById(\"pw\").value;if(pw){fd.append(\"password\",pw);}var x=new XMLHttpRequest();' +
'x.open(\"POST\",path);x.responseType=type;x.onload=function(){done(x);};x.onerror=function(){document.getElementById(\"st\").textContent=\"Network error\";};x.send(fd);}' +
'function parseIt(){var st=document.getElementById(\"st\");st.textContent=\"Parse ho raha hai... (pehli baar 30-60 sec lag sakte hain)\";' +
'post(\"/parse\",\"json\",function(x){var r=x.response;if(!r||x.status!==200){st.textContent=\"Error: \"+(r&&r.error?r.error:x.status);return;}show(r);});}' +
'function show(r){var t=r.transactions;document.getElementById(\"st\").textContent=t.length+\" transactions mile\";var i=r.info;' +
'document.getElementById(\"info\").innerHTML=\"<b>Bank:</b> \"+esc(i.bank)+\" &nbsp; <b>Account:</b> \"+esc(i.accountNo)+\" &nbsp; <b>Period:</b> \"+esc(i.period)+\" &nbsp; <b>IFSC:</b> \"+esc(i.ifsc)+\" &nbsp; <b>MICR:</b> \"+esc(i.micr);' +
'var h=\"<table><tr><th>#</th><th>Date</th><th>Narration</th><th>Category</th><th>Ref</th><th>Chq</th><th>Value Date</th><th>Withdrawal</th><th>Deposit</th><th>Balance</th></tr>\";var k;' +
'for(k=0;k<t.length;k++){var a=t[k];h+=\"<tr><td>\"+(k+1)+\"</td><td>\"+esc(a.date)+\"</td><td>\"+esc(a.narration)+\"</td><td>\"+esc(a.category)+\"</td><td>\"+esc(a.ref)+\"</td><td>\"+esc(a.chq)+\"</td><td>\"+esc(a.valueDate)+\"</td><td class=\\\"num\\\">\"+fmt(a.withdrawal)+\"</td><td class=\\\"num\\\">\"+fmt(a.deposit)+\"</td><td class=\\\"num\\\">\"+fmt(a.balance)+\"</td></tr>\";}' +
'h+=\"</table>\";document.getElementById(\"wrap\").innerHTML=h;document.getElementById(\"dbg\").value=JSON.stringify({info:r.info,debug:r.debug},null,2);}' +
'function dl(){var st=document.getElementById(\"st\");st.textContent=\"Excel ban raha hai...\";post(\"/convert\",\"blob\",function(x){if(x.status!==200){st.textContent=\"Excel error: \"+x.status;return;}' +
'var url=URL.createObjectURL(x.response);var a=document.createElement(\"a\");a.href=url;a.download=\"BankSync_Statement.xlsx\";document.body.appendChild(a);a.click();document.body.removeChild(a);st.textContent=\"Excel download ho gaya\";});}' +
'</script></body></html>';

app.get('/test', function (req, res) { res.send(TEST_PAGE); });

var PORT = process.env.PORT || 3000;
app.listen(PORT, function () { console.log('BankSync Pro running on port ' + PORT); });
