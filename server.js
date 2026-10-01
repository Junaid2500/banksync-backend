var express = require('express');
var multer = require('multer');
var cors = require('cors');
var ExcelJS = require('exceljs');
var PdfReader = require('pdfreader').PdfReader;

var app = express();
app.use(cors());
var upload = multer({ storage: multer.memoryStorage() });

var GEMINI_KEY = process.env.GEMINI_API_KEY;
var GEMINI_MODEL = process.env.GEMINI_MODEL || '';
var cachedModel = '';

// ---------- Model khud chuno ----------
async function pickModel() {
  if (GEMINI_MODEL) return GEMINI_MODEL;
  if (cachedModel) return cachedModel;

  var resp = await fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', {
    headers: { 'x-goog-api-key': GEMINI_KEY }
  });
  var data = await resp.json();
  if (!resp.ok) {
    throw new Error('AI error: ' + ((data.error && data.error.message) || resp.status));
  }
  var list = data.models || [];
  var best = '';
  var bestVer = -1;

  for (var i = 0; i < list.length; i++) {
    var name = String(list[i].name || '').replace('models/', '');
    var methods = list[i].supportedGenerationMethods || [];
    if (methods.indexOf('generateContent') === -1) continue;
    var m = name.match(/^gemini-(\d+(?:\.\d+)?)-flash$/);
    if (m) {
      var ver = parseFloat(m[1]);
      if (ver > bestVer) { bestVer = ver; best = name; }
    }
  }

  // Agar seedha naam na mile to koi bhi flash model
  if (!best) {
    for (var j = 0; j < list.length; j++) {
      var n2 = String(list[j].name || '').replace('models/', '');
      var m2 = list[j].supportedGenerationMethods || [];
      if (m2.indexOf('generateContent') === -1) continue;
      if (n2.indexOf('flash') === -1) continue;
      if (/lite|image|tts|live|audio|thinking|exp/.test(n2)) continue;
      best = n2;
      break;
    }
  }

  if (!best) throw new Error('AI error: koi Flash model nahi mila');
  console.log('Using model:', best);
  cachedModel = best;
  return best;
}

// ---------- 1. PDF se rows nikalo ----------
function readPdfRows(buffer, password) {
  return new Promise(function(resolve, reject) {
    var items = [];
    var page = 0;
    var opts = password ? { password: password } : {};
    var reader = new PdfReader(opts);
    reader.parseBuffer(buffer, function(err, item) {
      if (err) return reject(err);
      if (!item) return resolve(buildRows(items));
      if (item.page) page = item.page;
      if (item.text && item.text.trim()) {
        items.push({ page: page, x: item.x, y: item.y, text: item.text.trim() });
      }
    });
  });
}

function buildRows(items) {
  items.sort(function(a, b) {
    if (a.page !== b.page) return a.page - b.page;
    return a.y - b.y;
  });
  var rows = [];
  var cur = null;
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    if (!cur || cur.page !== it.page || Math.abs(it.y - cur.y) > 0.4) {
      cur = { page: it.page, y: it.y, cells: [] };
      rows.push(cur);
    }
    cur.cells.push(it);
  }
  var lines = [];
  for (var r = 0; r < rows.length; r++) {
    rows[r].cells.sort(function(a, b) { return a.x - b.x; });
    var parts = [];
    for (var c = 0; c < rows[r].cells.length; c++) parts.push(rows[r].cells[c].text);
    lines.push(parts.join(' | '));
  }
  return lines;
}

// ---------- 2. Gemini AI ----------
var PROMPT =
  'You are extracting data from an Indian bank statement. Each line below is one text row of the PDF, ' +
  'cells separated by " | " in left-to-right order.\n' +
  'Return ONLY JSON in this exact shape:\n' +
  '{"info":{"bank":"","accountNo":"","holder":"","period":"","ifsc":"","micr":""},' +
  '"transactions":[{"date":"","valueDate":"","narration":"","refNo":"","chqNo":"","withdrawal":null,"deposit":null,"balance":null,"category":""}]}\n' +
  'Rules:\n' +
  '- One object per transaction. If narration wraps onto several lines, join them into one narration.\n' +
  '- Skip headers, footers, page numbers, opening balance, totals, summaries.\n' +
  '- Numbers must be plain numbers without commas or currency (e.g. 54631.00 -> 54631).\n' +
  '- withdrawal = money going out (Debit/Dr/Withdrawal). deposit = money coming in (Credit/Cr/Deposit). The other one must be null.\n' +
  '- Use column headings and how the balance changes to decide debit or credit.\n' +
  '- Keep dates exactly as written in the PDF.\n' +
  '- category: one of UPI, IMPS, NEFT, RTGS, ATM, Cheque, Card, Salary, Interest, Charges, Transfer, Other.\n' +
  '- Do not invent data. If something is missing use empty string or null.\n' +
  '- If this text has no transactions, return an empty transactions array.\n\n' +
  'TEXT:\n';

function sleep(ms) {
  return new Promise(function(r) { setTimeout(r, ms); });
}

async function askGemini(text) {
  var body = {
    contents: [{ parts: [{ text: PROMPT + text }] }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json' }
  };
  var lastErr = null;

  for (var attempt = 1; attempt <= 3; attempt++) {
    try {
      var modelName = await pickModel();
      var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + modelName + ':generateContent';
      var resp = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-goog-api-key': GEMINI_KEY },
        body: JSON.stringify(body)
      });

      if (resp.status === 429 || resp.status >= 500) {
        lastErr = new Error('AI busy (' + resp.status + ')');
        await sleep(4000 * attempt);
        continue;
      }

      var data = await resp.json();

      if (!resp.ok) {
        var msg = (data.error && data.error.message) || String(resp.status);
        // Model purana/galat hai to auto-pick pe aa jao
        if (resp.status === 404 || /no longer available|not found/i.test(msg)) {
          GEMINI_MODEL = '';
          cachedModel = '';
          lastErr = new Error('AI error: ' + msg);
          continue;
        }
        throw new Error('AI error: ' + msg);
      }

      if (!data.candidates || !data.candidates[0] || !data.candidates[0].content) {
        lastErr = new Error('AI error: khali jawab aaya');
        await sleep(2000 * attempt);
        continue;
      }

      var out = data.candidates[0].content.parts[0].text;
      return JSON.parse(out);
    } catch (e) {
      lastErr = e;
      if (String(e.message).indexOf('AI error') === 0) throw e;
      await sleep(2000 * attempt);
    }
  }
  throw lastErr || new Error('AI failed');
}

async function runChunks(chunks) {
  var results = new Array(chunks.length);
  var next = 0;
  async function worker() {
    while (true) {
      var idx = next;
      next++;
      if (idx >= chunks.length) return;
      results[idx] = await askGemini(chunks[idx]);
    }
  }
  var workers = [];
  for (var w = 0; w < 3; w++) workers.push(worker());
  await Promise.all(workers);
  return results;
}

// ---------- 3. Safai + balance check ----------
function toNum(v) {
  if (v === null || v === undefined || v === '') return null;
  var n = parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

function cleanTransactions(list) {
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var t = list[i];
    if (!t || !t.date) continue;
    var tx = {
      date: String(t.date || ''),
      valueDate: String(t.valueDate || ''),
      narration: String(t.narration || ''),
      refNo: String(t.refNo || ''),
      chqNo: String(t.chqNo || ''),
      category: String(t.category || ''),
      withdrawal: toNum(t.withdrawal),
      deposit: toNum(t.deposit),
      balance: toNum(t.balance)
    };
    if (tx.withdrawal === null && tx.deposit === null && tx.balance === null) continue;
    out.push(tx);
  }
  for (var j = 1; j < out.length; j++) {
    var prev = out[j - 1].balance;
    var cur = out[j];
    if (prev === null || cur.balance === null) continue;
    var amt = cur.withdrawal !== null ? cur.withdrawal : cur.deposit;
    if (amt === null) continue;
    var delta = Math.round((cur.balance - prev) * 100) / 100;
    if (Math.abs(Math.abs(delta) - amt) < 0.02) {
      if (delta > 0) { cur.deposit = amt; cur.withdrawal = null; }
      else if (delta < 0) { cur.withdrawal = amt; cur.deposit = null; }
    }
  }
  return out;
}

// ---------- 4. Excel ----------
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
    ['Account', (info.accountNo || '') + (info.holder ? '  (' + info.holder + ')' : '')],
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

// ---------- 5. Routes ----------
app.get('/', function(req, res) {
  res.send('BankSync Pro Backend Running! AI key set: ' + (GEMINI_KEY ? 'YES' : 'NO'));
});

app.post('/convert', upload.single('pdf'), async function(req, res) {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    if (!GEMINI_KEY) return res.status(500).json({ error: 'Server me GEMINI_API_KEY set nahi hai' });

    var password = (req.body && req.body.password) || '';
    var lines;
    try {
      lines = await readPdfRows(req.file.buffer, password);
    } catch (e) {
      var m = String(e && e.message ? e.message : e).toLowerCase();
      if (m.indexOf('password') !== -1) {
        return res.status(400).json({ error: 'PDF password galat hai ya password chahiye' });
      }
      return res.status(400).json({ error: 'PDF read nahi hui: ' + (e.message || e) });
    }

    if (!lines.length) {
      return res.status(400).json({ error: 'PDF me text nahi mila. Scanned image PDF support nahi hai.' });
    }

    var CHUNK = 150;
    var chunks = [];
    for (var i = 0; i < lines.length; i += CHUNK) {
      chunks.push(lines.slice(i, i + CHUNK).join('\n'));
    }

    var results = await runChunks(chunks);

    var info = {};
    var all = [];
    for (var r = 0; r < results.length; r++) {
      var res1 = results[r] || {};
      if (res1.info) {
        var keys = Object.keys(res1.info);
        for (var k = 0; k < keys.length; k++) {
          if (res1.info[keys[k]] && !info[keys[k]]) info[keys[k]] = res1.info[keys[k]];
        }
      }
      if (res1.transactions) all = all.concat(res1.transactions);
    }

    var txs = cleanTransactions(all);
    if (!txs.length) {
      return res.status(400).json({ error: 'Transactions nahi mile. PDF check karo.' });
    }

    var buf = await buildExcel(info, txs);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="statement.xlsx"');
    return res.send(Buffer.from(buf));
  } catch (e) {
    console.log('Convert error:', e);
    return res.status(500).json({ error: String(e.message || e) });
  }
});

var PORT = process.env.PORT || 3000;
app.listen(PORT, function() {
  console.log('BankSync Pro running on port ' + PORT);
});
