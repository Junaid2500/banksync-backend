const express = require('express');
const multer = require('multer');
const pdfParse = require('pdf-parse');
const XLSX = require('xlsx');
const cors = require('cors');

const app = express();
app.use(cors());
const upload = multer({ storage: multer.memoryStorage() });

app.get('/', function(req, res) {
  res.json({ status: 'BankSync API Running!' });
});

app.post('/convert', upload.single('pdf'), function(req, res) {
  if (!req.file) {
    return res.status(400).json({ error: 'No PDF uploaded' });
  }

  pdfParse(req.file.buffer).then(function(data) {
    var text = data.text;
    var lines = text.split('\n');

    var bankInfo = detectBankInfo(lines);
    var transactions = parseTransactions(lines, bankInfo.bank);

    if (transactions.length === 0) {
      return res.status(400).json({ error: 'No transactions found. Please check PDF format.' });
    }

    var wb = buildExcel(transactions, bankInfo);
    var buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    res.setHeader('Content-Disposition', 'attachment; filename="bank-statement.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);

  }).catch(function(err) {
    res.status(500).json({ error: 'PDF parse error: ' + err.message });
  });
});

function detectBankInfo(lines) {
  var text = lines.join(' ').toUpperCase();
  var bank = 'UNKNOWN';
  var account = '';
  var period = '';
  var ifsc = '';
  var micr = '';

  if (text.indexOf('INDUSIND') !== -1) bank = 'INDUSIND';
  else if (text.indexOf('STATE BANK') !== -1 || text.indexOf('SBI') !== -1) bank = 'SBI';
  else if (text.indexOf('ICICI') !== -1) bank = 'ICICI';
  else if (text.indexOf('HDFC') !== -1) bank = 'HDFC';
  else if (text.indexOf('AXIS') !== -1) bank = 'AXIS';
  else if (text.indexOf('BANK OF MAHARASHTRA') !== -1 || text.indexOf('MAHABANK') !== -1) bank = 'BOM';
  else if (text.indexOf('KOTAK') !== -1) bank = 'KOTAK';
  else if (text.indexOf('YES BANK') !== -1) bank = 'YES';
  else if (text.indexOf('PUNJAB NATIONAL') !== -1 || text.indexOf('PNB') !== -1) bank = 'PNB';

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var lu = line.toUpperCase();

    if (lu.indexOf('ACCOUNT NO') !== -1 || lu.indexOf('A/C NO') !== -1 || lu.indexOf('ACCOUNT NUMBER') !== -1) {
      var match = line.match(/[\d]{8,}/);
      if (match) account = match[0];
    }
    if (lu.indexOf('IFSC') !== -1) {
      var m = line.match(/[A-Z]{4}0[A-Z0-9]{6}/);
      if (m) ifsc = m[0];
    }
    if (lu.indexOf('MICR') !== -1) {
      var m2 = line.match(/\d{9}/);
      if (m2) micr = m2[0];
    }
    if (lu.indexOf('PERIOD') !== -1 || lu.indexOf('STATEMENT') !== -1) {
      var m3 = line.match(/\d{2}[-\/]\w{3}[-\/]\d{4}\s*to\s*\d{2}[-\/]\w{3}[-\/]\d{4}/i);
      if (m3) period = m3[0];
    }
  }

  return { bank: bank, account: account, period: period, ifsc: ifsc, micr: micr };
}

function parseTransactions(lines, bank) {
  var transactions = [];
  var datePattern = /^(\d{2}[-\/]\w{3}[-\/]\d{4}|\d{2}[-\/]\d{2}[-\/]\d{4}|\d{2}\s\w{3}\s\d{4})/;

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;

    var dateMatch = line.match(datePattern);
    if (!dateMatch) continue;

    var date = dateMatch[0].trim();
    var rest = line.substring(date.length).trim();

    var numbers = [];
    var narration = '';
    var refNo = '';
    var chqNo = '';
    var valueDate = '';

    var numPattern = /[\d,]+\.\d{2}/g;
    var numMatches = rest.match(numPattern);
    var cleanRest = rest.replace(numPattern, '').trim();

    if (numMatches) {
      for (var n = 0; n < numMatches.length; n++) {
        numbers.push(parseFloat(numMatches[n].replace(/,/g, '')));
      }
    }

    var parts = cleanRest.split(/\s{2,}|\t/);
    narration = parts[0] ? parts[0].trim() : '';

    var refMatch = rest.match(/[A-Z0-9]{10,20}/);
    if (refMatch && refMatch[0] !== narration) refNo = refMatch[0];

    var chqMatch = rest.match(/\b\d{6,10}\b/);
    if (chqMatch) chqNo = chqMatch[0];

    var vdMatch = rest.match(/\d{2}[-\/]\w{3}[-\/]\d{4}/g);
    if (vdMatch && vdMatch.length > 1) valueDate = vdMatch[1];

    var withdrawal = '';
    var deposit = '';
    var balance = '';

    if (numbers.length >= 3) {
      withdrawal = numbers[0] > 0 ? numbers[0] : '';
      deposit = numbers[1] > 0 ? numbers[1] : '';
      balance = numbers[2];
    } else if (numbers.length === 2) {
      balance = numbers[1];
      var prevBalance = transactions.length > 0 ? transactions[transactions.length-1].balance : 0;
      if (numbers[0] < prevBalance) {
        withdrawal = numbers[0];
      } else {
        deposit = numbers[0];
      }
    } else if (numbers.length === 1) {
      balance = numbers[0];
    }

    if (!narration) {
      if (i + 1 < lines.length && !lines[i+1].match(datePattern)) {
        narration = lines[i+1].trim();
      }
    }

    var category = categorize(narration);

    transactions.push({
      date: date,
      narration: narration,
      category: category,
      refNo: refNo,
      chqNo: chqNo,
      valueDate: valueDate,
      withdrawal: withdrawal,
      deposit: deposit,
      balance: balance
    });
  }

  return transactions;
}

function categorize(narration) {
  if (!narration) return 'Other';
  var n = narration.toUpperCase();

  if (n.indexOf('CHEQUE') !== -1 || n.indexOf('CHQ') !== -1 || n.indexOf('CHQ') !== -1) return 'Cheque';
  if (n.indexOf('NEFT') !== -1 || n.indexOf('RTGS') !== -1 || n.indexOf('IMPS') !== -1) return 'Transfer';
  if (n.indexOf('ATM') !== -1 || n.indexOf('CASH') !== -1) return 'Cash/ATM';
  if (n.indexOf('UPI') !== -1) return 'UPI';
  if (n.indexOf('EMI') !== -1 || n.indexOf('LOAN') !== -1) return 'Loan/EMI';
  if (n.indexOf('SALARY') !== -1 || n.indexOf('SAL') !== -1) return 'Salary';
  if (n.indexOf('INTEREST') !== -1 || n.indexOf('INT') !== -1) return 'Interest';
  if (n.indexOf('TAX') !== -1 || n.indexOf('GST') !== -1) return 'Tax';
  if (n.indexOf('INSURANCE') !== -1) return 'Insurance';
  if (n.indexOf('DIVIDEND') !== -1) return 'Dividend';
  return 'Transfer';
}

function buildExcel(transactions, bankInfo) {
  var wb = XLSX.utils.book_new();
  var wsData = [];

  // Header info rows
  wsData.push(['ACCOUNT STATEMENT', '', '', '', '', '', '', '', '', '']);
  wsData.push(['', '', '', '', '', '', '', '', '', '']);
  wsData.push(['Bank', bankInfo.bank, '', '', '', '', '', '', '', '']);
  wsData.push(['Account', bankInfo.account, '', '', '', '', '', '', '', '']);
  wsData.push(['Period', bankInfo.period, '', '', '', '', '', '', '', '']);
  wsData.push(['IFSC', bankInfo.ifsc, '', '', '', '', '', '', '', '']);
  wsData.push(['MICR', bankInfo.micr, '', '', '', '', '', '', '', '']);
  wsData.push(['', '', '', '', '', '', '', '', '', '']);

  // Column headers
  wsData.push(['S.No', 'Date', 'Narration', 'Category', 'Ref No.', 'Chq No.', 'Value Date', 'Withdrawal (₹)', 'Deposit (₹)', 'Balance (₹)']);

  // Transaction rows
  for (var i = 0; i < transactions.length; i++) {
    var t = transactions[i];
    wsData.push([
      i + 1,
      t.date,
      t.narration,
      t.category,
      t.refNo,
      t.chqNo,
      t.valueDate,
      t.withdrawal || '',
      t.deposit || '',
      t.balance || ''
    ]);
  }

  var ws = XLSX.utils.aoa_to_sheet(wsData);

  // Column widths
  ws['!cols'] = [
    { wch: 6 },
    { wch: 14 },
    { wch: 45 },
    { wch: 12 },
    { wch: 18 },
    { wch: 12 },
    { wch: 12 },
    { wch: 16 },
    { wch: 14 },
    { wch: 14 }
  ];

  // Styling
  var darkGreen = '1F7A4D';
  var lightGreen = 'E8F5E9';
  var white = 'FFFFFF';

  // Title row style
  if (ws['A1']) {
    ws['A1'].s = {
      font: { bold: true, sz: 14, color: { rgb: white } },
      fill: { fgColor: { rgb: darkGreen } },
      alignment: { horizontal: 'left' }
    };
  }

  // Header row (row 9 = index 8)
  var headerCols = ['A','B','C','D','E','F','G','H','I','J'];
  for (var c = 0; c < headerCols.length; c++) {
    var cell = headerCols[c] + '9';
    if (ws[cell]) {
      ws[cell].s = {
        font: { bold: true, color: { rgb: white } },
        fill: { fgColor: { rgb: darkGreen } },
        alignment: { horizontal: 'center' }
      };
    }
  }

  // Alternating row colors
  for (var r = 0; r < transactions.length; r++) {
    var rowNum = r + 10;
    var bgColor = (r % 2 === 0) ? lightGreen : white;
    for (var c2 = 0; c2 < headerCols.length; c2++) {
      var cellRef = headerCols[c2] + rowNum;
      if (ws[cellRef]) {
        ws[cellRef].s = {
          fill: { fgColor: { rgb: bgColor } },
          alignment: { horizontal: c2 >= 7 ? 'right' : 'left' }
        };
      }
    }
  }

  XLSX.utils.book_append_sheet(wb, ws, 'Statement');
  return wb;
}

var PORT = process.env.PORT || 3000;
app.listen(PORT, function() {
  console.log('BankSync server running on port ' + PORT);
});
