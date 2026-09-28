var express = require('express');
var multer = require('multer');
var xlsx = require('xlsx');
var cors = require('cors');
var { PdfReader } = require('pdfreader');

var app = express();
app.use(cors());
var upload = multer({ storage: multer.memoryStorage() });

function findHeaderLine(lines, keywords) {
  for (var i = 0; i < lines.length; i++) {
    var upper = lines[i].text.toUpperCase();
    var found = 0;
    for (var k = 0; k < keywords.length; k++) {
      if (upper.indexOf(keywords[k]) !== -1) found++;
    }
    if (found >= 2) return i;
  }
  return -1;
}

function parseICICI(lines) {
  var headerIdx = findHeaderLine(lines, ['DATE', 'PARTICULARS', 'AMOUNT']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['DATE', 'DEBIT', 'CREDIT']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['DATE', 'NARRATION', 'BALANCE']);
  if (headerIdx === -1) headerIdx = 0;

  var transactions = [];
  var dateRegex = /^(\d{2}[\/\-]\d{2}[\/\-]\d{4}|\d{2}[\/\-]\d{2}[\/\-]\d{2})$/;
  var amountRegex = /^[\d,]+\.\d{2}$/;

  var i = headerIdx + 1;
  while (i < lines.length) {
    var text = lines[i].text.trim();
    if (dateRegex.test(text)) {
      var date = text;
      var narration = '';
      var refNo = '';
      var valueDate = '';
      var withdrawal = '';
      var deposit = '';
      var balance = '';
      i++;

      while (i < lines.length && !dateRegex.test(lines[i].text.trim())) {
        var t = lines[i].text.trim();
        if (amountRegex.test(t)) {
          if (!withdrawal && !deposit && !balance) {
            balance = t;
          } else if (!withdrawal && !deposit) {
            deposit = balance;
            balance = t;
          } else if (!withdrawal) {
            withdrawal = deposit;
            deposit = balance;
            balance = t;
          }
        } else if (dateRegex.test(t) && !valueDate) {
          valueDate = t;
        } else if (t.length > 2) {
          narration = narration ? narration + ' ' + t : t;
        }
        i++;
      }

      if (balance) {
        transactions.push({
          date: date,
          narration: narration,
          refNo: refNo,
          valueDate: valueDate || date,
          withdrawal: withdrawal,
          deposit: deposit,
          balance: balance
        });
      }
    } else {
      i++;
    }
  }
  return transactions;
}

function parseBOM(lines) {
  var headerIdx = findHeaderLine(lines, ['DATE', 'DEBIT', 'CREDIT']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['DATE', 'PARTICULARS', 'BALANCE']);
  if (headerIdx === -1) headerIdx = 0;

  var transactions = [];
  var dateRegex = /^\d{2}\/\d{2}\/\d{4}$/;
  var amountRegex = /^[\d,]+\.\d{2}$/;

  var i = headerIdx + 1;
  while (i < lines.length) {
    var text = lines[i].text.trim();
    if (dateRegex.test(text)) {
      var date = text;
      var narration = '';
      var withdrawal = '';
      var deposit = '';
      var balance = '';
      i++;

      while (i < lines.length && !dateRegex.test(lines[i].text.trim())) {
        var t = lines[i].text.trim();
        if (amountRegex.test(t)) {
          if (!withdrawal && !deposit && !balance) {
            balance = t;
          } else if (!withdrawal && !deposit) {
            deposit = balance;
            balance = t;
          } else if (!withdrawal) {
            withdrawal = deposit;
            deposit = balance;
            balance = t;
          }
        } else if (t.length > 2) {
          narration = narration ? narration + ' ' + t : t;
        }
        i++;
      }

      if (balance) {
        transactions.push({
          date: date,
          narration: narration,
          refNo: '',
          valueDate: date,
          withdrawal: withdrawal,
          deposit: deposit,
          balance: balance
        });
      }
    } else {
      i++;
    }
  }
  return transactions;
}

function parseIndusInd(lines) {
  var headerIdx = findHeaderLine(lines, ['DATE', 'WITHDRAWAL', 'DEPOSIT']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['DATE', 'DEBIT', 'CREDIT']);
  if (headerIdx === -1) headerIdx = 0;

  var transactions = [];
  var dateRegex = /^\d{2}[\-\/]\w{3}[\-\/]\d{4}$|^\d{2}[\-\/]\d{2}[\-\/]\d{4}$/;
  var amountRegex = /^[\d,]+\.\d{2}$/;

  var i = headerIdx + 1;
  while (i < lines.length) {
    var text = lines[i].text.trim();
    if (dateRegex.test(text)) {
      var date = text;
      var narration = '';
      var withdrawal = '';
      var deposit = '';
      var balance = '';
      i++;

      while (i < lines.length && !dateRegex.test(lines[i].text.trim())) {
        var t = lines[i].text.trim();
        if (amountRegex.test(t)) {
          if (!withdrawal && !deposit && !balance) {
            balance = t;
          } else if (!withdrawal && !deposit) {
            deposit = balance;
            balance = t;
          } else if (!withdrawal) {
            withdrawal = deposit;
            deposit = balance;
            balance = t;
          }
        } else if (t.length > 2) {
          narration = narration ? narration + ' ' + t : t;
        }
        i++;
      }

      if (balance) {
        transactions.push({
          date: date,
          narration: narration,
          refNo: '',
          valueDate: date,
          withdrawal: withdrawal,
          deposit: deposit,
          balance: balance
        });
      }
    } else {
      i++;
    }
  }
  return transactions;
}

function parseSBI(lines) {
  var headerIdx = findHeaderLine(lines, ['DATE', 'DEBIT', 'CREDIT']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['DATE', 'PARTICULARS', 'BALANCE']);
  if (headerIdx === -1) headerIdx = 0;

  var transactions = [];
  var dateRegex = /^\d{2}\/\d{2}\/\d{4}$|^\d{2}-\d{2}-\d{4}$/;
  var amountRegex = /^[\d,]+\.\d{2}$/;

  var i = headerIdx + 1;
  while (i < lines.length) {
    var text = lines[i].text.trim();
    if (dateRegex.test(text)) {
      var date = text;
      var narration = '';
      var withdrawal = '';
      var deposit = '';
      var balance = '';
      i++;

      while (i < lines.length && !dateRegex.test(lines[i].text.trim())) {
        var t = lines[i].text.trim();
        if (amountRegex.test(t)) {
          if (!withdrawal && !deposit && !balance) {
            balance = t;
          } else if (!withdrawal && !deposit) {
            deposit = balance;
            balance = t;
          } else if (!withdrawal) {
            withdrawal = deposit;
            deposit = balance;
            balance = t;
          }
        } else if (t.length > 2) {
          narration = narration ? narration + ' ' + t : t;
        }
        i++;
      }

      if (balance) {
        transactions.push({
          date: date,
          narration: narration,
          refNo: '',
          valueDate: date,
          withdrawal: withdrawal,
          deposit: deposit,
          balance: balance
        });
      }
    } else {
      i++;
    }
  }
  return transactions;
}

function detectBank(lines) {
  for (var i = 0; i < Math.min(lines.length, 30); i++) {
    var t = lines[i].text.toUpperCase();
    if (t.indexOf('ICICI') !== -1) return 'ICICI';
    if (t.indexOf('INDUSIND') !== -1 || t.indexOf('INDUS IND') !== -1) return 'INDUSIND';
    if (t.indexOf('BANK OF MAHARASHTRA') !== -1 || t.indexOf('MAHABANK') !== -1) return 'BOM';
    if (t.indexOf('STATE BANK OF INDIA') !== -1 || t.indexOf('SBI') !== -1) return 'SBI';
    if (t.indexOf('HDFC') !== -1) return 'HDFC';
    if (t.indexOf('AXIS') !== -1) return 'AXIS';
  }
  return 'UNKNOWN';
}

function buildExcel(transactions, bankName) {
  var wb = xlsx.utils.book_new();
  var wsData = [];

  wsData.push(['ACCOUNT STATEMENT - ' + bankName]);
  wsData.push([]);
  wsData.push(['S.No', 'Date', 'Narration', 'Category', 'Ref No.', 'Chq No.', 'Value Date', 'Withdrawal (₹)', 'Deposit (₹)', 'Balance (₹)']);

  for (var i = 0; i < transactions.length; i++) {
    var t = transactions[i];
    wsData.push([
      i + 1,
      t.date,
      t.narration,
      '',
      t.refNo || '',
      '',
      t.valueDate || t.date,
      t.withdrawal || '',
      t.deposit || '',
      t.balance || ''
    ]);
  }

  var ws = xlsx.utils.aoa_to_sheet(wsData);
  xlsx.utils.book_append_sheet(wb, ws, 'Statement');
  return xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

app.get('/', function(req, res) {
  res.send('BankSync Pro Backend Running!');
});

app.post('/convert', upload.single('pdf'), function(req, res) {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  var lines = [];
  var reader = new PdfReader();

  reader.parseBuffer(req.file.buffer, function(err, item) {
    if (err) {
      return res.status(500).json({ error: 'PDF read error: ' + err.message });
    }

    if (!item) {
      // PDF parsing done
      var bank = detectBank(lines);
      var transactions = [];

      if (bank === 'ICICI') transactions = parseICICI(lines);
      else if (bank === 'BOM') transactions = parseBOM(lines);
      else if (bank === 'INDUSIND') transactions = parseIndusInd(lines);
      else if (bank === 'SBI') transactions = parseSBI(lines);
      else transactions = parseICICI(lines); // fallback

      if (transactions.length === 0) {
        return res.status(400).json({ error: 'No transactions found. Bank: ' + bank });
      }

      var excelBuffer = buildExcel(transactions, bank);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="statement.xlsx"');
      return res.send(excelBuffer);
    }

    if (item.text) {
      lines.push({ text: item.text, x: item.x, y: item.y, page: item.page });
    }
  });
});

var PORT = process.env.PORT || 3000;
app.listen(PORT, function() {
  console.log('BankSync Pro running on port ' + PORT);
});
