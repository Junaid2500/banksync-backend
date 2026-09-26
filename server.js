var express = require('express');
var multer = require('multer');
var pdfParse = require('pdf-parse');
var XLSX = require('xlsx');
var cors = require('cors');
var app = express();
app.use(cors());
app.use(express.json());
var upload = multer({ storage: multer.memoryStorage() });

app.get('/', function(req, res) {
  res.json({ status: 'BankSync API Running!' });
});

app.post('/convert', upload.single('pdf'), function(req, res) {
  if (!req.file) {
    return res.status(400).json({ error: 'No PDF uploaded' });
  }
  pdfParse(req.file.buffer).then(function(data) {
    var lines = data.text.split('\n');
    var bankName = detectBank(data.text.toLowerCase());
    var transactions = parseTransactions(lines);
    if (transactions.length === 0) {
      return res.json({ error: 'No transactions found', bank: bankName });
    }
    var wb = XLSX.utils.book_new();
    var headers = ['S.No','Date','Narration','Category','Ref No.','Chq No.','Value Date','Withdrawal (Rs)','Deposit (Rs)','Balance (Rs)'];
    var sheetData = [headers];
    transactions.forEach(function(t, i) {
      sheetData.push([i+1, t.date, t.narration, t.category, t.refNo||'', t.chqNo||'', t.valueDate||t.date, t.debit||'', t.credit||'', t.balance||'']);
    });
    var ws = XLSX.utils.aoa_to_sheet(sheetData);
    ws['!cols'] = [{wch:6},{wch:12},{wch:45},{wch:12},{wch:22},{wch:12},{wch:12},{wch:16},{wch:16},{wch:16}];
    XLSX.utils.book_append_sheet(wb, ws, 'Transactions');
    var summary = buildSummary(transactions, bankName);
    var ws2 = XLSX.utils.aoa_to_sheet(summary);
    XLSX.utils.book_append_sheet(wb, ws2, 'Summary');
    var buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=BankSync.xlsx');
    res.send(buf);
  }).catch(function(err) {
    res.status(500).json({ error: err.message });
  });
});

function detectBank(txt) {
  if (txt.indexOf('indusind') !== -1) return 'IndusInd Bank';
  if (txt.indexOf('state bank of india') !== -1) return 'State Bank of India';
  if (txt.indexOf('hdfc bank') !== -1) return 'HDFC Bank';
  if (txt.indexOf('icici bank') !== -1) return 'ICICI Bank';
  if (txt.indexOf('axis bank') !== -1) return 'Axis Bank';
  if (txt.indexOf('kotak') !== -1) return 'Kotak Bank';
  if (txt.indexOf('bank of maharashtra') !== -1) return 'Bank of Maharashtra';
  if (txt.indexOf('punjab national') !== -1) return 'Punjab National Bank';
  if (txt.indexOf('bank of baroda') !== -1) return 'Bank of Baroda';
  if (txt.indexOf('canara bank') !== -1) return 'Canara Bank';
  if (txt.indexOf('union bank') !== -1) return 'Union Bank';
  if (txt.indexOf('yes bank') !== -1) return 'Yes Bank';
  if (txt.indexOf('federal bank') !== -1) return 'Federal Bank';
  if (txt.indexOf('idfc') !== -1) return 'IDFC First Bank';
  return 'Universal Bank';
}

function parseTransactions(lines) {
  var transactions = [];
  var current = null;
  var datePatterns = [
    /\b(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})\b/,
    /\b(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2})\b/,
    /\b(\d{1,2}[\/\-][A-Za-z]{3}[\/\-]\d{4})\b/,
    /\b(\d{1,2}[\/\-][A-Za-z]{3}[\/\-]\d{2})\b/,
    /\b(\d{1,2}\s[A-Za-z]{3}\s\d{4})\b/
  ];
  var skipWords = ['brought forward','carried forward','opening balance','statement of','account number','branch','ifsc','customer id','page no','total','subtotal','dear customer','toll free','www.','http','note:','disclaimer','system generated'];
  function hasSkip(line) {
    var l = line.toLowerCase();
    for (var i = 0; i < skipWords.length; i++) {
      if (l.indexOf(skipWords[i]) !== -1) return true;
    }
    return false;
  }
  function findDate(line) {
    for (var i = 0; i < datePatterns.length; i++) {
      var m = line.match(datePatterns[i]);
      if (m) return m[1];
    }
    return null;
  }
  function findAmounts(line) {
    var matches = line.match(/[\d,]+\.\d{2}/g);
    if (!matches) return [];
    return matches.map(function(m) { return parseFloat(m.replace(/,/g,'')); });
  }
  function cleanNarr(line, date) {
    var n = line;
    if (date) n = n.replace(date, '');
    n = n.replace(/[\d,]+\.\d{2}/g, '');
    n = n.replace(/\b\d{10}\b/g, '');
    n = n.replace(/\s+/g, ' ').trim();
    return n.substring(0, 150);
  }
  lines.forEach(function(line) {
    line = line.trim();
    if (!line || line.length < 3) return;
    if (hasSkip(line)) return;
    var date = findDate(line);
    var amounts = findAmounts(line);
    if (date) {
      if (current && (current.debit || current.credit || current.balance)) {
        transactions.push(current);
      }
      var narration = cleanNarr(line, date);
      var debit = '', credit = '', balance = '';
      if (amounts.length >= 3) { debit = amounts[amounts.length-3]; credit = amounts[amounts.length-2]; balance = amounts[amounts.length-1]; }
      else if (amounts.length === 2) { debit = amounts[0]; balance = amounts[1]; }
      else if (amounts.length === 1) { balance = amounts[0]; }
      current = { date:date, narration:narration, category:categorize(narration), debit:debit, credit:credit, balance:balance, refNo:'', chqNo:'', valueDate:date };
    } else if (current) {
      var amounts2 = findAmounts(line);
      if (amounts2.length > 0 && !current.balance) {
        if (amounts2.length >= 3) { current.debit = amounts2[amounts2.length-3]; current.credit = amounts2[amounts2.length-2]; current.balance = amounts2[amounts2.length-1]; }
        else if (amounts2.length === 2) { current.debit = amounts2[0]; current.balance = amounts2[1]; }
        else if (amounts2.length === 1) { current.balance = amounts2[0]; }
      } else if (amounts2.length === 0 && !hasSkip(line)) {
        current.narration = (current.narration + ' ' + line).trim().substring(0, 150);
      }
    }
  });
  if (current && (current.debit || current.credit || current.balance)) transactions.push(current);
  return transactions;
}

function categorize(narr) {
  if (!narr) return 'Other';
  var n = narr.toUpperCase();
  if (n.indexOf('NEFT') !== -1) return 'NEFT';
  if (n.indexOf('RTGS') !== -1) return 'RTGS';
  if (n.indexOf('IMPS') !== -1) return 'IMPS';
  if (n.indexOf('UPI') !== -1) return 'UPI';
  if (n.indexOf('ATM') !== -1 || n.indexOf('CASH') !== -1) return 'Cash';
  if (n.indexOf('SALARY') !== -1) return 'Salary';
  if (n.indexOf('EMI') !== -1 || n.indexOf('LOAN') !== -1) return 'Loan/EMI';
  if (n.indexOf('BILL') !== -1 || n.indexOf('BBPS') !== -1) return 'Bill Pay';
  if (n.indexOf('CHQ') !== -1 || n.indexOf('CHEQUE') !== -1) return 'Cheque';
  if (n.indexOf('RETURN') !== -1 || n.indexOf('BOUNCE') !== -1) return 'Return';
  if (n.indexOf('INTEREST') !== -1) return 'Interest';
  if (n.indexOf('TAX') !== -1 || n.indexOf('TDS') !== -1) return 'Tax';
  return 'Other';
}

function buildSummary(transactions, bankName) {
  var cat = {};
  var tDr = 0, tCr = 0, lBal = 0;
  transactions.forEach(function(t) {
    if (!cat[t.category]) cat[t.category] = { dr:0, cr:0, count:0 };
    var dr = parseFloat(t.debit) || 0;
    var cr = parseFloat(t.credit) || 0;
    cat[t.category].dr += dr;
    cat[t.category].cr += cr;
    cat[t.category].count++;
    tDr += dr; tCr += cr;
    if (t.balance) lBal = parseFloat(t.balance) || 0;
  });
  var rows = [
    ['BankSync Pro — Summary'],
    ['Bank', bankName],
    ['Generated', new Date().toLocaleString('en-IN')],
    ['Total Transactions', transactions.length],
    [],
    ['Category','Count','Withdrawal','Deposit','Net']
  ];
  Object.keys(cat).forEach(function(c) {
    var v = cat[c];
    rows.push([c, v.count, v.dr||'', v.cr||'', v.cr-v.dr]);
  });
  rows.push([]);
  rows.push(['TOTAL', transactions.length, tDr, tCr, tCr-tDr]);
  rows.push(['Closing Balance', '', '', '', lBal]);
  return rows;
}

var PORT = process.env.PORT || 3000;
app.listen(PORT, function() {
  console.log('BankSync API running on port ' + PORT);
});