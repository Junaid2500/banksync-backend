var express = require('express');
var multer = require('multer');
var XLSX = require('xlsx');
var cors = require('cors');
var pdfreader = require('pdfreader');

var app = express();
app.use(cors());
var upload = multer({ storage: multer.memoryStorage() });

app.get('/', function(req, res) {
  res.json({ status: 'BankSync API Running!' });
});

app.post('/convert', upload.single('pdf'), function(req, res) {
  if (!req.file) {
    return res.status(400).json({ error: 'No PDF uploaded' });
  }

  var rows = {};
  var password = req.body.password || '';

  new pdfreader.PdfReader({ password: password }).parseBuffer(req.file.buffer, function(err, item) {
    if (err) {
      return res.status(500).json({ error: 'PDF Error: ' + err });
    }

    if (!item) {
      try {
        var lines = buildLines(rows);
        var bankInfo = detectBank(lines);
        var transactions = parseTransactions(lines, bankInfo.bank);

        if (transactions.length === 0) {
          return res.status(400).json({
            error: 'No transactions found. Bank: ' + bankInfo.bank
          });
        }

        var wb = buildExcel(transactions, bankInfo);
        var buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        res.setHeader('Content-Disposition', 'attachment; filename="bank-statement.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buf);
      } catch(e) {
        res.status(500).json({ error: 'Processing error: ' + e.message });
      }
      return;
    }

    if (item.text) {
      var y = (item.y).toFixed(1);
      if (!rows[y]) rows[y] = [];
      rows[y].push({ x: item.x, text: item.text });
    }
  });
});

function buildLines(rows) {
  var lines = [];
  var yKeys = Object.keys(rows).sort(function(a, b) {
    return parseFloat(a) - parseFloat(b);
  });

  yKeys.forEach(function(y) {
    var items = rows[y].sort(function(a, b) { return a.x - b.x; });
    var lineText = items.map(function(i) { return i.text; }).join(' ');
    lines.push({
      text: lineText.trim(),
      items: items,
      y: parseFloat(y)
    });
  });

  return lines;
}

function detectBank(lines) {
  var fullText = lines.map(function(l) { return l.text; }).join(' ').toUpperCase();

  var bank = 'UNKNOWN';
  var account = '';
  var period = '';
  var ifsc = '';

  if (fullText.indexOf('BANK OF MAHARASHTRA') !== -1 || fullText.indexOf('MAHABANK') !== -1) bank = 'BOM';
  else if (fullText.indexOf('INDUSIND') !== -1) bank = 'INDUSIND';
  else if (fullText.indexOf('ICICI') !== -1) bank = 'ICICI';
  else if (fullText.indexOf('STATE BANK OF INDIA') !== -1) bank = 'SBI';
  else if (fullText.indexOf('HDFC BANK') !== -1) bank = 'HDFC';
  else if (fullText.indexOf('AXIS BANK') !== -1) bank = 'AXIS';
  else if (fullText.indexOf('KOTAK') !== -1) bank = 'KOTAK';
  else if (fullText.indexOf('YES BANK') !== -1) bank = 'YES';
  else if (fullText.indexOf('PUNJAB NATIONAL') !== -1) bank = 'PNB';
  else if (fullText.indexOf('UNION BANK') !== -1) bank = 'UNION';
  else if (fullText.indexOf('CANARA') !== -1) bank = 'CANARA';
  else if (fullText.indexOf('BANK OF BARODA') !== -1) bank = 'BOB';

  var am = fullText.match(/ACCOUNT\s*(?:NO|NUMBER|NO\.)[:\s]*(\d{9,18})/);
  if (am) account = am[1];
  if (!account) {
    var am2 = fullText.match(/A\/C\s*(?:NO)?[:\s]*(\d{9,18})/);
    if (am2) account = am2[1];
  }

  var pm = fullText.match(/(\d{2}[\/\-]\d{2}[\/\-]\d{4})\s*(?:TO)\s*(\d{2}[\/\-]\d{2}[\/\-]\d{4})/);
  if (pm) period = pm[1] + ' to ' + pm[2];
  if (!period) {
    var pm2 = fullText.match(/(\d{2}[\/\-][A-Z]{3}[\/\-]\d{4})\s*(?:TO)\s*(\d{2}[\/\-][A-Z]{3}[\/\-]\d{4})/i);
    if (pm2) period = pm2[1] + ' to ' + pm2[2];
  }

  var im = fullText.match(/[A-Z]{4}0[A-Z0-9]{6}/);
  if (im) ifsc = im[0];

  return { bank: bank, account: account, period: period, ifsc: ifsc };
}

function isAmount(t) {
  return /^-?[\d,]+\.\d{2}$/.test(t.replace(/\s/g, ''));
}

function toNum(t) {
  return parseFloat(String(t).replace(/,/g, ''));
}

function categorize(narration) {
  if (!narration) return 'Other';
  var n = narration.toUpperCase();
  if (n.indexOf('CHEQUE') !== -1 || n.indexOf('CHQ') !== -1) return 'Cheque';
  if (n.indexOf('NEFT') !== -1 || n.indexOf('RTGS') !== -1 || n.indexOf('IMPS') !== -1) return 'Transfer';
  if (n.indexOf('ATM') !== -1 || n.indexOf('CASH') !== -1) return 'Cash/ATM';
  if (n.indexOf('UPI') !== -1) return 'UPI';
  if (n.indexOf('EMI') !== -1 || n.indexOf('LOAN') !== -1) return 'Loan/EMI';
  if (n.indexOf('SALARY') !== -1) return 'Salary';
  if (n.indexOf('INTEREST') !== -1) return 'Interest';
  if (n.indexOf('GST') !== -1 || n.indexOf('TAX') !== -1) return 'Tax';
  if (n.indexOf('PROCESSING FEE') !== -1 || n.indexOf('CHARGES') !== -1) return 'Charges';
  return 'Transfer';
}

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

function parseTransactions(lines, bank) {
  if (bank === 'BOM') return parseBOM(lines);
  if (bank === 'INDUSIND') return parseIndusInd(lines);
  if (bank === 'ICICI') return parseICICI(lines);
  if (bank === 'SBI') return parseSBI(lines);
  if (bank === 'HDFC') return parseHDFC(lines);
  return parseGeneric(lines);
}

function parseBOM(lines) {
  var transactions = [];
  var dateReg = /^\d{2}\/\d{2}\/\d{4}$/;

  // Find header line first
  var headerIdx = findHeaderLine(lines, ['DATE', 'DEBIT', 'CREDIT']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['DATE', 'PARTICULARS']);
  var startIdx = headerIdx !== -1 ? headerIdx + 1 : 0;

  for (var i = startIdx; i < lines.length; i++) {
    var line = lines[i];
    var items = line.items;
    if (!items || items.length < 2) continue;

    var first = items[0].text.trim();
    if (!dateReg.test(first)) continue;

    var date = first;
    var textParts = [];
    var nums = [];

    items.forEach(function(it) {
      var t = it.text.trim();
      if (t === date) return;
      if (isAmount(t)) nums.push(toNum(t));
      else textParts.push(t);
    });

    var j = i + 1;
    while (j < lines.length && j <= i + 3) {
      var nextLine = lines[j];
      var nextFirst = nextLine.items && nextLine.items[0] ? nextLine.items[0].text.trim() : '';
      if (dateReg.test(nextFirst)) break;

      nextLine.items.forEach(function(it) {
        var t = it.text.trim();
        if (isAmount(t)) nums.push(toNum(t));
        else if (t.length > 1) textParts.push(t);
      });
      j++;
      if (nums.length >= 2) break;
    }
    i = j - 1;

    var narration = textParts.join(' ').replace(/\s+/g, ' ').trim();
    var debit = '', credit = '', balance = '';

    if (nums.length >= 3) {
      debit = nums[0] || '';
      credit = nums[1] || '';
      balance = nums[2];
    } else if (nums.length === 2) {
      var prevBal = transactions.length > 0 ? transactions[transactions.length-1].balRaw : 0;
      if (nums[1] > prevBal) { credit = nums[0]; }
      else { debit = nums[0]; }
      balance = nums[1];
    } else if (nums.length === 1) {
      balance = nums[0];
    }

    transactions.push({
      date: date, narration: narration, category: categorize(narration),
      refNo: '', chqNo: '', valueDate: '',
      withdrawal: debit || '', deposit: credit || '', balance: balance || '',
      balRaw: balance ? toNum(String(balance)) : 0
    });
  }
  return transactions;
}

function parseIndusInd(lines) {
  var transactions = [];
  var dateReg = /^\d{2}-[A-Za-z]{3}-\d{4}$/;

  var headerIdx = findHeaderLine(lines, ['DATE', 'WITHDRAWAL', 'DEPOSIT']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['DATE', 'PARTICULARS']);
  var startIdx = headerIdx !== -1 ? headerIdx + 1 : 0;

  for (var i = startIdx; i < lines.length; i++) {
    var line = lines[i];
    var items = line.items;
    if (!items || items.length < 2) continue;

    var first = items[0].text.trim();
    if (!dateReg.test(first)) continue;

    var date = first;
    var textParts = [];
    var nums = [];

    items.forEach(function(it) {
      var t = it.text.trim();
      if (t === date) return;
      if (isAmount(t)) nums.push(toNum(t));
      else textParts.push(t);
    });

    var j = i + 1;
    while (j < lines.length && j <= i + 4) {
      var nextLine = lines[j];
      var nextFirst = nextLine.items && nextLine.items[0] ? nextLine.items[0].text.trim() : '';
      if (dateReg.test(nextFirst)) break;

      nextLine.items.forEach(function(it) {
        var t = it.text.trim();
        if (isAmount(t)) nums.push(toNum(t));
        else if (t.length > 1) textParts.push(t);
      });
      j++;
      if (nums.length >= 3) break;
    }
    i = j - 1;

    var narration = textParts.join(' ').replace(/\s+/g, ' ').trim();
    var withdrawal = '', deposit = '', balance = '';

    if (nums.length >= 3) {
      withdrawal = nums[0] || '';
      deposit = nums[1] || '';
      balance = nums[2];
    } else if (nums.length === 2) {
      var prevBal2 = transactions.length > 0 ? transactions[transactions.length-1].balRaw : 0;
      if (nums[1] > prevBal2) { deposit = nums[0]; }
      else { withdrawal = nums[0]; }
      balance = nums[1];
    } else if (nums.length === 1) {
      balance = nums[0];
    }

    transactions.push({
      date: date, narration: narration, category: categorize(narration),
      refNo: '', chqNo: '', valueDate: '',
      withdrawal: withdrawal || '', deposit: deposit || '', balance: balance || '',
      balRaw: balance ? toNum(String(balance)) : 0
    });
  }
  return transactions;
}

function parseICICI(lines) {
  var transactions = [];
  var dateReg = /^\d{2}\/\d{2}\/\d{4}$|^\d{2}-\d{2}-\d{4}$/;

  var headerIdx = findHeaderLine(lines, ['DATE', 'DEBIT', 'CREDIT']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['DATE', 'PARTICULARS']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['DATE', 'WITHDRAWAL']);
  var startIdx = headerIdx !== -1 ? headerIdx + 1 : 0;

  for (var i = startIdx; i < lines.length; i++) {
    var line = lines[i];
    var items = line.items;
    if (!items || items.length < 2) continue;

    var first = items[0].text.trim();
    if (!dateReg.test(first)) continue;

    var date = first;
    var textParts = [];
    var nums = [];

    items.forEach(function(it) {
      var t = it.text.trim();
      if (t === date) return;
      if (isAmount(t)) nums.push(toNum(t));
      else textParts.push(t);
    });

    var j = i + 1;
    while (j < lines.length && j <= i + 3) {
      var nextLine = lines[j];
      var nextFirst = nextLine.items && nextLine.items[0] ? nextLine.items[0].text.trim() : '';
      if (dateReg.test(nextFirst)) break;

      nextLine.items.forEach(function(it) {
        var t = it.text.trim();
        if (isAmount(t)) nums.push(toNum(t));
        else if (t.length > 1) textParts.push(t);
      });
      j++;
      if (nums.length >= 3) break;
    }
    i = j - 1;

    var narration = textParts.join(' ').replace(/\s+/g, ' ').trim();
    var withdrawal = '', deposit = '', balance = '';

    if (nums.length >= 3) {
      withdrawal = nums[0] || '';
      deposit = nums[1] || '';
      balance = nums[2];
    } else if (nums.length === 2) {
      deposit = nums[0];
      balance = nums[1];
    } else if (nums.length === 1) {
      balance = nums[0];
    }

    transactions.push({
      date: date, narration: narration, category: categorize(narration),
      refNo: '', chqNo: '', valueDate: '',
      withdrawal: withdrawal || '', deposit: deposit || '', balance: balance || '',
      balRaw: balance ? toNum(String(balance)) : 0
    });
  }
  return transactions;
}

function parseSBI(lines) {
  var transactions = [];
  var dateReg = /^\d{2}\s[A-Za-z]{3}\s\d{4}$|^\d{2}\/\d{2}\/\d{4}$|^\d{2}-\d{2}-\d{4}$/;

  var headerIdx = findHeaderLine(lines, ['DATE', 'DEBIT', 'CREDIT']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['DATE', 'NARRATION']);
  var startIdx = headerIdx !== -1 ? headerIdx + 1 : 0;

  for (var i = startIdx; i < lines.length; i++) {
    var line = lines[i];
    var items = line.items;
    if (!items || items.length < 2) continue;

    var first = items[0].text.trim();
    if (!dateReg.test(first)) continue;

    var date = first;
    var textParts = [];
    var nums = [];

    items.forEach(function(it) {
      var t = it.text.trim();
      if (t === date) return;
      if (isAmount(t)) nums.push(toNum(t));
      else textParts.push(t);
    });

    var narration = textParts.join(' ').replace(/\s+/g, ' ').trim();
    var withdrawal = '', deposit = '', balance = '';

    if (nums.length >= 3) {
      withdrawal = nums[0] || '';
      deposit = nums[1] || '';
      balance = nums[2];
    } else if (nums.length === 2) {
      deposit = nums[0];
      balance = nums[1];
    } else if (nums.length === 1) {
      balance = nums[0];
    }

    transactions.push({
      date: date, narration: narration, category: categorize(narration),
      refNo: '', chqNo: '', valueDate: '',
      withdrawal: withdrawal || '', deposit: deposit || '', balance: balance || '',
      balRaw: balance ? toNum(String(balance)) : 0
    });
  }
  return transactions;
}

function parseHDFC(lines) {
  var transactions = [];
  var dateReg = /^\d{2}\/\d{2}\/\d{2}$|^\d{2}-\d{2}-\d{4}$|^\d{2}\/\d{2}\/\d{4}$/;

  var headerIdx = findHeaderLine(lines, ['DATE', 'DEBIT', 'CREDIT']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['DATE', 'NARRATION']);
  var startIdx = headerIdx !== -1 ? headerIdx + 1 : 0;

  for (var i = startIdx; i < lines.length; i++) {
    var line = lines[i];
    var items = line.items;
    if (!items || items.length < 2) continue;

    var first = items[0].text.trim();
    if (!dateReg.test(first)) continue;

    var date = first;
    var textParts = [];
    var nums = [];

    items.forEach(function(it) {
      var t = it.text.trim();
      if (t === date) return;
      if (isAmount(t)) nums.push(toNum(t));
      else textParts.push(t);
    });

    var narration = textParts.join(' ').replace(/\s+/g, ' ').trim();
    var withdrawal = '', deposit = '', balance = '';

    if (nums.length >= 3) {
      withdrawal = nums[0] || '';
      deposit = nums[1] || '';
      balance = nums[2];
    } else if (nums.length === 2) {
      deposit = nums[0];
      balance = nums[1];
    } else if (nums.length === 1) {
      balance = nums[0];
    }

    transactions.push({
      date: date, narration: narration, category: categorize(narration),
      refNo: '', chqNo: '', valueDate: '',
      withdrawal: withdrawal || '', deposit: deposit || '', balance: balance || '',
      balRaw: balance ? toNum(String(balance)) : 0
    });
  }
  return transactions;
}

function parseGeneric(lines) {
  var transactions = [];
  var dateReg = /^\d{2}[\/\-]\d{2}[\/\-]\d{2,4}$|^\d{2}[\/\-][A-Za-z]{3}[\/\-]\d{4}$/;

  var headerIdx = findHeaderLine(lines, ['DATE', 'DEBIT', 'CREDIT']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['DATE', 'NARRATION']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['DATE', 'PARTICULARS']);
  var startIdx = headerIdx !== -1 ? headerIdx + 1 : 0;

  for (var i = startIdx; i < lines.length; i++) {
    var line = lines[i];
    var items = line.items;
    if (!items || items.length < 2) continue;

    var first = items[0].text.trim();
    if (!dateReg.test(first)) continue;

    var date = first;
    var textParts = [];
    var nums = [];

    items.forEach(function(it) {
      var t = it.text.trim();
      if (t === date) return;
      if (isAmount(t)) nums.push(toNum(t));
      else textParts.push(t);
    });

    var narration = textParts.join(' ').replace(/\s+/g, ' ').trim();
    var withdrawal = '', deposit = '', balance = '';

    if (nums.length >= 3) {
      withdrawal = nums[0] || '';
      deposit = nums[1] || '';
      balance = nums[2];
    } else if (nums.length === 2) {
      deposit = nums[0];
      balance = nums[1];
    } else if (nums.length === 1) {
      balance = nums[0];
    }

    transactions.push({
      date: date, narration: narration, category: categorize(narration),
      refNo: '', chqNo: '', valueDate: '',
      withdrawal: withdrawal || '', deposit: deposit || '', balance: balance || '',
      balRaw: balance ? toNum(String(balance)) : 0
    });
  }
  return transactions;
}

function buildExcel(transactions, bankInfo) {
  var wb = XLSX.utils.book_new();
  var wsData = [];

  wsData.push(['ACCOUNT STATEMENT', '', '', '', '', '', '', '', '', '']);
  wsData.push(['', '', '', '', '', '', '', '', '', '']);
  wsData.push(['Bank', bankInfo.bank, '', '', '', '', '', '', '', '']);
  wsData.push(['Account', bankInfo.account, '', '', '', '', '', '', '', '']);
  wsData.push(['Period', bankInfo.period, '', '', '', '', '', '', '', '']);
  wsData.push(['IFSC', bankInfo.ifsc, '', '', '', '', '', '', '', '']);
  wsData.push(['', '', '', '', '', '', '', '', '', '']);
  wsData.push(['S.No', 'Date', 'Narration', 'Category', 'Ref No.', 'Chq No.', 'Value Date', 'Withdrawal (₹)', 'Deposit (₹)', 'Balance (₹)']);

  for (var i = 0; i < transactions.length; i++) {
    var t = transactions[i];
    wsData.push([
      i + 1, t.date, t.narration, t.category,
      t.refNo || '', t.chqNo || '', t.valueDate || '',
      t.withdrawal || '', t.deposit || '', t.balance || ''
    ]);
  }

  var ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [
    { wch: 6 }, { wch: 14 }, { wch: 50 }, { wch: 12 },
    { wch: 20 }, { wch: 12 }, { wch: 12 },
    { wch: 16 }, { wch: 14 }, { wch: 14 }
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Statement');
  return wb;
}

var PORT = process.env.PORT || 3000;
app.listen(PORT, function() {
  console.log('BankSync server running on port ' + PORT);
});
