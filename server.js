var express = require('express');
var multer = require('multer');
var pdfParse = require('pdf-parse');
var XLSX = require('xlsx');
var cors = require('cors');

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

  pdfParse(req.file.buffer).then(function(data) {
    var text = data.text;
    var lines = text.split('\n');
    var bankInfo = detectBankInfo(lines, text);
    var transactions = parseTransactions(lines, text, bankInfo.bank);

    if (transactions.length === 0) {
      return res.status(400).json({ error: 'No transactions found. Bank: ' + bankInfo.bank });
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

function detectBankInfo(lines, text) {
  var upper = text.toUpperCase();
  var bank = 'UNKNOWN';
  var account = '';
  var period = '';
  var ifsc = '';
  var micr = '';

  if (upper.indexOf('BANK OF MAHARASHTRA') !== -1 || upper.indexOf('MAHABANK') !== -1 || upper.indexOf('MAHB') !== -1) {
    bank = 'BOM';
  } else if (upper.indexOf('INDUSIND') !== -1) {
    bank = 'INDUSIND';
  } else if (upper.indexOf('STATE BANK OF INDIA') !== -1 || upper.indexOf('SBI') !== -1) {
    bank = 'SBI';
  } else if (upper.indexOf('ICICI BANK') !== -1) {
    bank = 'ICICI';
  } else if (upper.indexOf('HDFC BANK') !== -1) {
    bank = 'HDFC';
  } else if (upper.indexOf('AXIS BANK') !== -1) {
    bank = 'AXIS';
  } else if (upper.indexOf('KOTAK') !== -1) {
    bank = 'KOTAK';
  } else if (upper.indexOf('YES BANK') !== -1) {
    bank = 'YES';
  } else if (upper.indexOf('PUNJAB NATIONAL') !== -1) {
    bank = 'PNB';
  }

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var lu = line.toUpperCase();

    if (lu.indexOf('ACCOUNT NO') !== -1 || lu.indexOf('A/C NO') !== -1 || lu.indexOf('ACC NO') !== -1) {
      var am = line.match(/\d{9,18}/);
      if (am) account = am[0];
    }
    if (!account) {
      var accm = line.match(/Account No[:\s]+(\d{9,18})/i);
      if (accm) account = accm[1];
    }
    if (lu.indexOf('IFSC') !== -1) {
      var im = line.match(/[A-Z]{4}0[A-Z0-9]{6}/);
      if (im) ifsc = im[0];
    }
    if (lu.indexOf('MICR') !== -1) {
      var mm = line.match(/\d{9}/);
      if (mm) micr = mm[0];
    }
    if (lu.indexOf('PERIOD') !== -1 || (lu.indexOf('FROM') !== -1 && lu.indexOf('TO') !== -1)) {
      var pm = line.match(/\d{2}[\/\-]\d{2}[\/\-]\d{4}/g);
      if (pm && pm.length >= 2) period = pm[0] + ' to ' + pm[1];
    }
    if (!period) {
      var pm2 = text.match(/(\d{2}[\/\-]\d{2}[\/\-]\d{4})\s+to\s+(\d{2}[\/\-]\d{2}[\/\-]\d{4})/i);
      if (pm2) period = pm2[1] + ' to ' + pm2[2];
    }
    if (!period) {
      var pm3 = text.match(/(\d{2}[\/\-]\w{3}[\/\-]\d{4})\s+to\s+(\d{2}[\/\-]\w{3}[\/\-]\d{4})/i);
      if (pm3) period = pm3[1] + ' to ' + pm3[2];
    }
  }

  // Extract account from statement line
  if (!account) {
    var stm = text.match(/Account No\s+(\d{9,18})/i);
    if (stm) account = stm[1];
    var stm2 = text.match(/(\d{11,18})\s+from/i);
    if (stm2) account = stm2[1];
  }

  return { bank: bank, account: account, period: period, ifsc: ifsc, micr: micr };
}

function parseTransactions(lines, text, bank) {
  if (bank === 'BOM') {
    return parseBOM(lines, text);
  } else if (bank === 'INDUSIND') {
    return parseIndusInd(lines, text);
  } else {
    return parseGeneric(lines, text);
  }
}

// =================== BOM PARSER ===================
function parseBOM(lines, text) {
  var transactions = [];
  // BOM date format: DD/MM/YYYY
  var datePattern = /^\d{2}\/\d{2}\/\d{4}$/;

  var i = 0;
  while (i < lines.length) {
    var line = lines[i].trim();

    // Check if line starts with a date
    var dateParts = line.split(/\s+/);
    var firstWord = dateParts[0] ? dateParts[0].trim() : '';

    if (!firstWord.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
      i++;
      continue;
    }

    var date = firstWord;
    var rest = line.substring(date.length).trim();

    // Collect narration - may span multiple lines
    var narration = '';
    var chqRef = '';
    var debit = '';
    var credit = '';
    var balance = '';
    var channel = '';

    // Extract numbers from rest
    var nums = extractNumbers(rest);
    var cleanText = rest.replace(/[\d,]+\.\d{2}/g, '').replace(/\s+/g, ' ').trim();

    // Check next lines for continuation
    var j = i + 1;
    while (j < lines.length) {
      var nextLine = lines[j].trim();
      if (!nextLine) { j++; break; }
      // If next line starts with date, stop
      if (nextLine.match(/^\d{2}\/\d{2}\/\d{4}/)) break;
      // If next line has numbers that look like amounts, it might be part of same transaction
      cleanText += ' ' + nextLine;
      var moreNums = extractNumbers(nextLine);
      nums = nums.concat(moreNums);
      j++;
      // Stop after collecting enough
      if (nums.length >= 2) break;
    }

    // Parse narration - text before numbers
    var narrationMatch = cleanText.replace(/[\d,]+\.\d{2}/g, '').replace(/\s+/g, ' ').trim();
    // Remove channel info (branch names like 278-MALEGAON)
    narrationMatch = narrationMatch.replace(/\d{3}-[A-Z\s\(\)]+$/i, '').trim();
    narration = narrationMatch;

    // Extract cheque/ref number
    var chqMatch = rest.match(/\b\d{6,15}\b/);
    if (chqMatch) chqRef = chqMatch[0];

    // Assign amounts
    // BOM: Debit | Credit | Balance
    if (nums.length >= 3) {
      debit = nums[0];
      credit = nums[1];
      balance = nums[2];
    } else if (nums.length === 2) {
      // Could be debit+balance or credit+balance
      var prevBal = transactions.length > 0 ? parseFloat(transactions[transactions.length-1].balance) : 0;
      var diff = Math.abs(parseFloat(nums[1]) - prevBal);
      if (Math.abs(parseFloat(nums[1]) - prevBal + parseFloat(nums[0])) < 1) {
        debit = nums[0];
        balance = nums[1];
      } else {
        credit = nums[0];
        balance = nums[1];
      }
    } else if (nums.length === 1) {
      balance = nums[0];
    }

    if (narration && narration.length > 1) {
      transactions.push({
        date: date,
        narration: narration,
        category: categorize(narration),
        refNo: chqRef,
        chqNo: '',
        valueDate: '',
        withdrawal: debit ? formatNum(debit) : '',
        deposit: credit ? formatNum(credit) : '',
        balance: balance ? formatNum(balance) : ''
      });
    }

    i = j;
  }

  return transactions;
}

// =================== INDUSIND PARSER ===================
function parseIndusInd(lines, text) {
  var transactions = [];
  // IndusInd date: 01-Jul-2026
  var datePattern = /^\d{2}-[A-Za-z]{3}-\d{4}/;

  var i = 0;
  while (i < lines.length) {
    var line = lines[i].trim();

    if (!line.match(datePattern)) {
      i++;
      continue;
    }

    // Skip header row
    if (line.toUpperCase().indexOf('DATE') !== -1 && line.toUpperCase().indexOf('PARTICULARS') !== -1) {
      i++;
      continue;
    }

    var dateMatch = line.match(/^\d{2}-[A-Za-z]{3}-\d{4}/);
    var date = dateMatch[0];
    var rest = line.substring(date.length).trim();

    // Collect full narration (may be multi-line)
    var narrationLines = [rest];
    var j = i + 1;

    while (j < lines.length) {
      var nextLine = lines[j].trim();
      if (!nextLine) { j++; break; }
      if (nextLine.match(datePattern)) break;
      // If line has amounts, it's still part of this transaction
      narrationLines.push(nextLine);
      j++;
      if (narrationLines.join(' ').match(/[\d,]+\.\d{2}.*[\d,]+\.\d{2}/)) break;
    }

    var fullText = narrationLines.join(' ');
    var nums = extractNumbers(fullText);

    // Remove numbers from narration
    var narration = fullText.replace(/[\d,]+\.\d{2}/g, '').replace(/\s+/g, ' ').trim();

    // Extract ref no (long alphanumeric)
    var refMatch = fullText.match(/[A-Z0-9]{10,25}/);
    var refNo = refMatch ? refMatch[0] : '';

    // Extract chq no (6-9 digits standalone)
    var chqMatch = fullText.match(/\b(\d{6,9})\b/);
    var chqNo = chqMatch ? chqMatch[1] : '';

    // IndusInd: Withdrawal | Deposit | Balance
    var withdrawal = '';
    var deposit = '';
    var balance = '';

    if (nums.length >= 3) {
      withdrawal = nums[0];
      deposit = nums[1];
      balance = nums[2];
    } else if (nums.length === 2) {
      var prevBal = transactions.length > 0 ? parseFloat(transactions[transactions.length-1].balance.replace(/,/g,'')) : 0;
      var n0 = parseFloat(nums[0].replace(/,/g,''));
      var n1 = parseFloat(nums[1].replace(/,/g,''));
      if (n1 > prevBal) {
        deposit = nums[0];
        balance = nums[1];
      } else {
        withdrawal = nums[0];
        balance = nums[1];
      }
    } else if (nums.length === 1) {
      balance = nums[0];
    }

    // Clean narration
    narration = narration.replace(/[:\-\/]+$/, '').trim();

    if (narration && date) {
      transactions.push({
        date: date,
        narration: narration,
        category: categorize(narration),
        refNo: refNo,
        chqNo: chqNo,
        valueDate: '',
        withdrawal: withdrawal ? formatNum(withdrawal) : '',
        deposit: deposit ? formatNum(deposit) : '',
        balance: balance ? formatNum(balance) : ''
      });
    }

    i = j;
  }

  return transactions;
}

// =================== GENERIC PARSER ===================
function parseGeneric(lines, text) {
  var transactions = [];
  var datePatterns = [
    /^\d{2}\/\d{2}\/\d{4}/,
    /^\d{2}-[A-Za-z]{3}-\d{4}/,
    /^\d{2}-\d{2}-\d{4}/,
    /^\d{2}\s[A-Za-z]{3}\s\d{4}/
  ];

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;

    var dateMatch = null;
    for (var p = 0; p < datePatterns.length; p++) {
      var m = line.match(datePatterns[p]);
      if (m) { dateMatch = m[0]; break; }
    }
    if (!dateMatch) continue;

    var date = dateMatch;
    var rest = line.substring(date.length).trim();

    // Collect next line if needed
    if (i + 1 < lines.length && !lines[i+1].trim().match(datePatterns[0])) {
      rest += ' ' + lines[i+1].trim();
    }

    var nums = extractNumbers(rest);
    var narration = rest.replace(/[\d,]+\.\d{2}/g, '').replace(/\s+/g, ' ').trim();

    var withdrawal = '';
    var deposit = '';
    var balance = '';

    if (nums.length >= 3) {
      withdrawal = nums[0];
      deposit = nums[1];
      balance = nums[2];
    } else if (nums.length === 2) {
      deposit = nums[0];
      balance = nums[1];
    } else if (nums.length === 1) {
      balance = nums[0];
    }

    if (narration && date) {
      transactions.push({
        date: date,
        narration: narration,
        category: categorize(narration),
        refNo: '',
        chqNo: '',
        valueDate: '',
        withdrawal: withdrawal ? formatNum(withdrawal) : '',
        deposit: deposit ? formatNum(deposit) : '',
        balance: balance ? formatNum(balance) : ''
      });
    }
  }

  return transactions;
}

// =================== HELPERS ===================
function extractNumbers(text) {
  var matches = text.match(/[\d,]+\.\d{2}/g);
  return matches ? matches : [];
}

function formatNum(numStr) {
  if (!numStr) return '';
  var n = parseFloat(String(numStr).replace(/,/g, ''));
  if (isNaN(n)) return '';
  return n;
}

function categorize(narration) {
  if (!narration) return 'Other';
  var n = narration.toUpperCase();
  if (n.indexOf('CHEQUE') !== -1 || n.indexOf('CHQ') !== -1) return 'Cheque';
  if (n.indexOf('NEFT') !== -1 || n.indexOf('RTGS') !== -1 || n.indexOf('IMPS') !== -1) return 'Transfer';
  if (n.indexOf('ATM') !== -1 || n.indexOf('CASH') !== -1) return 'Cash/ATM';
  if (n.indexOf('UPI') !== -1) return 'UPI';
  if (n.indexOf('EMI') !== -1 || n.indexOf('LOAN') !== -1) return 'Loan/EMI';
  if (n.indexOf('SALARY') !== -1 || n.indexOf('SAL/') !== -1) return 'Salary';
  if (n.indexOf('INTEREST') !== -1 || n.indexOf('INT ') !== -1) return 'Interest';
  if (n.indexOf('TAX') !== -1 || n.indexOf('GST') !== -1) return 'Tax';
  if (n.indexOf('INSURANCE') !== -1) return 'Insurance';
  if (n.indexOf('PROCESSING FEE') !== -1 || n.indexOf('CHARGES') !== -1) return 'Charges';
  return 'Transfer';
}

// =================== EXCEL BUILDER ===================
function buildExcel(transactions, bankInfo) {
  var wb = XLSX.utils.book_new();
  var wsData = [];

  wsData.push(['ACCOUNT STATEMENT', '', '', '', '', '', '', '', '', '']);
  wsData.push(['', '', '', '', '', '', '', '', '', '']);
  wsData.push(['Bank', bankInfo.bank, '', '', '', '', '', '', '', '']);
  wsData.push(['Account', bankInfo.account, '', '', '', '', '', '', '', '']);
  wsData.push(['Period', bankInfo.period, '', '', '', '', '', '', '', '']);
  wsData.push(['IFSC', bankInfo.ifsc, '', '', '', '', '', '', '', '']);
  wsData.push(['MICR', bankInfo.micr, '', '', '', '', '', '', '', '']);
  wsData.push(['', '', '', '', '', '', '', '', '', '']);
  wsData.push(['S.No', 'Date', 'Narration', 'Category', 'Ref No.', 'Chq No.', 'Value Date', 'Withdrawal (₹)', 'Deposit (₹)', 'Balance (₹)']);

  for (var i = 0; i < transactions.length; i++) {
    var t = transactions[i];
    wsData.push([
      i + 1,
      t.date,
      t.narration,
      t.category,
      t.refNo || '',
      t.chqNo || '',
      t.valueDate || '',
      t.withdrawal || '',
      t.deposit || '',
      t.balance || ''
    ]);
  }

  var ws = XLSX.utils.aoa_to_sheet(wsData);

  ws['!cols'] = [
    { wch: 6 }, { wch: 14 }, { wch: 50 }, { wch: 12 },
    { wch: 20 }, { wch: 12 }, { wch: 12 }, { wch: 16 },
    { wch: 14 }, { wch: 14 }
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Statement');
  return wb;
}

var PORT = process.env.PORT || 3000;
app.listen(PORT, function() {
  console.log('BankSync server running on port ' + PORT);
});
