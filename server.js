var express = require('express');
var multer = require('multer');
var xlsx = require('xlsx');
var cors = require('cors');
var { PdfReader } = require('pdfreader');

var app = express();
app.use(cors());
var upload = multer({ storage: multer.memoryStorage() });

// Y coordinate ke basis pe rows group karo
function groupByRows(items) {
  var rows = {};
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var y = Math.round(item.y * 10) / 10;
    if (!rows[y]) rows[y] = [];
    rows[y].push(item);
  }
  
  var sortedKeys = Object.keys(rows).sort(function(a, b) {
    return parseFloat(a) - parseFloat(b);
  });
  
  var result = [];
  for (var k = 0; k < sortedKeys.length; k++) {
    var row = rows[sortedKeys[k]];
    row.sort(function(a, b) { return a.x - b.x; });
    result.push(row);
  }
  return result;
}

// Bank detect karo
function detectBank(items) {
  for (var i = 0; i < Math.min(items.length, 100); i++) {
    var t = items[i].text.toUpperCase();
    if (t.indexOf('ICICI') !== -1) return 'ICICI';
    if (t.indexOf('INDUSIND') !== -1 || t.indexOf('INDUS IND') !== -1) return 'INDUSIND';
    if (t.indexOf('BANK OF MAHARASHTRA') !== -1 || t.indexOf('MAHABANK') !== -1) return 'BOM';
    if (t.indexOf('STATE BANK') !== -1 || t.indexOf('SBI') !== -1) return 'SBI';
    if (t.indexOf('HDFC') !== -1) return 'HDFC';
    if (t.indexOf('AXIS BANK') !== -1) return 'AXIS';
    if (t.indexOf('KOTAK') !== -1) return 'KOTAK';
    if (t.indexOf('YES BANK') !== -1) return 'YES';
    if (t.indexOf('PUNJAB NATIONAL') !== -1 || t.indexOf('PNB') !== -1) return 'PNB';
    if (t.indexOf('UNION BANK') !== -1) return 'UNION';
    if (t.indexOf('CANARA') !== -1) return 'CANARA';
    if (t.indexOf('BANK OF BARODA') !== -1 || t.indexOf('BOB') !== -1) return 'BOB';
  }
  return 'UNKNOWN';
}

// Date check karo
function isDate(text) {
  if (!text) return false;
  var t = text.trim();
  // DD/MM/YYYY
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(t)) return true;
  // DD-MM-YYYY
  if (/^\d{2}-\d{2}-\d{4}$/.test(t)) return true;
  // DD-Mon-YYYY (01-May-2026)
  if (/^\d{2}-[A-Za-z]{3}-\d{4}$/.test(t)) return true;
  // DD Mon YYYY (01 May 2026)
  if (/^\d{2}\s[A-Za-z]{3}\s\d{4}$/.test(t)) return true;
  // DD/MM/YY
  if (/^\d{2}\/\d{2}\/\d{2}$/.test(t)) return true;
  // YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return true;
  return false;
}

// Amount check karo
function isAmount(text) {
  if (!text) return false;
  var t = text.trim().replace(/,/g, '');
  return /^\d+\.\d{2}$/.test(t);
}

// Header row dhundho
function findHeaderRow(rows) {
  var keywords = ['DATE', 'NARRATION', 'PARTICULARS', 'DESCRIPTION', 
                  'DEBIT', 'CREDIT', 'WITHDRAWAL', 'DEPOSIT', 
                  'BALANCE', 'AMOUNT', 'TRANSACTION'];
  
  for (var i = 0; i < Math.min(rows.length, 50); i++) {
    var rowText = '';
    for (var j = 0; j < rows[i].length; j++) {
      rowText += rows[i][j].text.toUpperCase() + ' ';
    }
    
    var matchCount = 0;
    for (var k = 0; k < keywords.length; k++) {
      if (rowText.indexOf(keywords[k]) !== -1) matchCount++;
    }
    
    if (matchCount >= 2) return i;
  }
  return -1;
}

// Column positions map karo header se
function mapColumns(headerRow) {
  var colMap = {
    date: -1,
    narration: -1,
    debit: -1,
    credit: -1,
    balance: -1,
    refNo: -1,
    cheque: -1,
    valueDate: -1
  };
  
  for (var i = 0; i < headerRow.length; i++) {
    var t = headerRow[i].text.toUpperCase();
    var x = headerRow[i].x;
    
    if (t.indexOf('DATE') !== -1 && colMap.date === -1) colMap.date = x;
    if ((t.indexOf('NARRATION') !== -1 || t.indexOf('PARTICULARS') !== -1 || 
         t.indexOf('DESCRIPTION') !== -1) && colMap.narration === -1) colMap.narration = x;
    if ((t.indexOf('DEBIT') !== -1 || t.indexOf('WITHDRAWAL') !== -1 || 
         t.indexOf('DR') !== -1) && colMap.debit === -1) colMap.debit = x;
    if ((t.indexOf('CREDIT') !== -1 || t.indexOf('DEPOSIT') !== -1 || 
         t.indexOf('CR') !== -1) && colMap.credit === -1) colMap.credit = x;
    if (t.indexOf('BALANCE') !== -1) colMap.balance = x;
    if (t.indexOf('REF') !== -1 || t.indexOf('TRANSACTION ID') !== -1 || 
        t.indexOf('CHEQUENO') !== -1) colMap.refNo = x;
    if (t.indexOf('CHEQUE') !== -1 || t.indexOf('CHQ') !== -1) colMap.cheque = x;
    if (t.indexOf('VALUE') !== -1) colMap.valueDate = x;
  }
  
  return colMap;
}

// X position se column identify karo
function getColumnByX(x, colMap, tolerance) {
  tolerance = tolerance || 5;
  
  if (colMap.balance !== -1 && Math.abs(x - colMap.balance) < tolerance) return 'balance';
  if (colMap.credit !== -1 && Math.abs(x - colMap.credit) < tolerance) return 'credit';
  if (colMap.debit !== -1 && Math.abs(x - colMap.debit) < tolerance) return 'debit';
  if (colMap.narration !== -1 && Math.abs(x - colMap.narration) < tolerance) return 'narration';
  if (colMap.date !== -1 && Math.abs(x - colMap.date) < tolerance) return 'date';
  if (colMap.refNo !== -1 && Math.abs(x - colMap.refNo) < tolerance) return 'refNo';
  if (colMap.cheque !== -1 && Math.abs(x - colMap.cheque) < tolerance) return 'cheque';
  if (colMap.valueDate !== -1 && Math.abs(x - colMap.valueDate) < tolerance) return 'valueDate';
  
  // Agar exact match nahi toh nearest column
  var minDist = 999;
  var nearest = 'narration';
  var cols = Object.keys(colMap);
  for (var c = 0; c < cols.length; c++) {
    if (colMap[cols[c]] !== -1) {
      var dist = Math.abs(x - colMap[cols[c]]);
      if (dist < minDist) {
        minDist = dist;
        nearest = cols[c];
      }
    }
  }
  return nearest;
}

// Universal parse function
function universalParse(items) {
  var rows = groupByRows(items);
  var headerIdx = findHeaderRow(rows);
  
  if (headerIdx === -1) {
    // Fallback - date se dhundho
    return fallbackParse(items);
  }
  
  var colMap = mapColumns(rows[headerIdx]);
  var transactions = [];
  var currentTx = null;
  
  for (var i = headerIdx + 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row || row.length === 0) continue;
    
    // Row ka pehla item check karo
    var firstText = row[0].text.trim();
    
    // Naya transaction start hota hai jab date milti hai
    if (isDate(firstText) || (row.length > 1 && isDate(row[1].text.trim()))) {
      if (currentTx && currentTx.balance) {
        transactions.push(currentTx);
      }
      
      currentTx = {
        date: '',
        narration: '',
        refNo: '',
        chequeNo: '',
        valueDate: '',
        withdrawal: '',
        deposit: '',
        balance: ''
      };
      
      // Row ke items ko columns mein daalo
      for (var j = 0; j < row.length; j++) {
        var item = row[j];
        var t = item.text.trim();
        var col = getColumnByX(item.x, colMap, 8);
        
        if (col === 'date' && isDate(t)) {
          currentTx.date = t;
        } else if (col === 'date' && isAmount(t)) {
          // Date column mein amount aa gaya
          assignAmount(currentTx, t);
        } else if (col === 'narration' || col === 'date') {
          if (!isDate(t) && !isAmount(t) && t.length > 1) {
            currentTx.narration = currentTx.narration ? 
              currentTx.narration + ' ' + t : t;
          }
        } else if (col === 'debit') {
          if (isAmount(t)) currentTx.withdrawal = t;
        } else if (col === 'credit') {
          if (isAmount(t)) currentTx.deposit = t;
        } else if (col === 'balance') {
          if (isAmount(t)) currentTx.balance = t;
        } else if (col === 'refNo') {
          currentTx.refNo = t;
        } else if (col === 'cheque') {
          currentTx.chequeNo = t;
        } else if (col === 'valueDate') {
          if (isDate(t)) currentTx.valueDate = t;
        } else {
          // Unknown column - amount hai toh assign karo
          if (isAmount(t)) assignAmount(currentTx, t);
          else if (!isDate(t) && t.length > 1) {
            currentTx.narration = currentTx.narration ? 
              currentTx.narration + ' ' + t : t;
          }
        }
      }
      
      if (!currentTx.date) currentTx.date = firstText;
      
    } else if (currentTx) {
      // Continuation row — narration ya amount
      for (var jj = 0; jj < row.length; jj++) {
        var itm = row[jj];
        var txt = itm.text.trim();
        var column = getColumnByX(itm.x, colMap, 8);
        
        if (isAmount(txt)) {
          if (column === 'debit') currentTx.withdrawal = txt;
          else if (column === 'credit') currentTx.deposit = txt;
          else if (column === 'balance') currentTx.balance = txt;
          else assignAmount(currentTx, txt);
        } else if (isDate(txt) && !currentTx.valueDate) {
          currentTx.valueDate = txt;
        } else if (txt.length > 1 && txt !== 'Dr' && txt !== 'Cr' && 
                   txt !== 'DR' && txt !== 'CR') {
          currentTx.narration = currentTx.narration ? 
            currentTx.narration + ' ' + txt : txt;
        }
      }
    }
  }
  
  if (currentTx && currentTx.balance) {
    transactions.push(currentTx);
  }
  
  return transactions;
}

// Amount assign karo — balance pehle, phir deposit, phir withdrawal
function assignAmount(tx, amount) {
  if (!tx.balance) {
    tx.balance = amount;
  } else if (!tx.deposit && !tx.withdrawal) {
    tx.deposit = tx.balance;
    tx.balance = amount;
  } else if (!tx.withdrawal) {
    tx.withdrawal = tx.deposit;
    tx.deposit = tx.balance;
    tx.balance = amount;
  }
}

// Fallback parser — sirf date aur amounts se
function fallbackParse(items) {
  var transactions = [];
  var currentTx = null;
  
  for (var i = 0; i < items.length; i++) {
    var t = items[i].text.trim();
    
    if (isDate(t)) {
      if (currentTx && currentTx.balance) {
        transactions.push(currentTx);
      }
      currentTx = {
        date: t,
        narration: '',
        refNo: '',
        chequeNo: '',
        valueDate: t,
        withdrawal: '',
        deposit: '',
        balance: ''
      };
    } else if (currentTx) {
      if (isAmount(t)) {
        assignAmount(currentTx, t);
      } else if (t.length > 2 && t !== 'Dr' && t !== 'Cr' && 
                 t !== 'DR' && t !== 'CR') {
        currentTx.narration = currentTx.narration ? 
          currentTx.narration + ' ' + t : t;
      }
    }
  }
  
  if (currentTx && currentTx.balance) {
    transactions.push(currentTx);
  }
  
  return transactions;
}

function buildExcel(transactions, bankName) {
  var wb = xlsx.utils.book_new();
  var wsData = [];

  wsData.push(['ACCOUNT STATEMENT - ' + bankName]);
  wsData.push([]);
  wsData.push([
    'S.No', 'Date', 'Narration', 'Category', 
    'Ref No.', 'Chq No.', 'Value Date', 
    'Withdrawal (₹)', 'Deposit (₹)', 'Balance (₹)'
  ]);

  for (var i = 0; i < transactions.length; i++) {
    var t = transactions[i];
    wsData.push([
      i + 1,
      t.date,
      t.narration,
      '',
      t.refNo || '',
      t.chequeNo || '',
      t.valueDate || t.date,
      t.withdrawal || '',
      t.deposit || '',
      t.balance || ''
    ]);
  }

  var ws = xlsx.utils.aoa_to_sheet(wsData);
  
  // Column widths
  ws['!cols'] = [
    { wch: 5 },  // S.No
    { wch: 12 }, // Date
    { wch: 40 }, // Narration
    { wch: 12 }, // Category
    { wch: 15 }, // Ref No
    { wch: 12 }, // Chq No
    { wch: 12 }, // Value Date
    { wch: 15 }, // Withdrawal
    { wch: 15 }, // Deposit
    { wch: 15 }  // Balance
  ];
  
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

  var items = [];
  var reader = new PdfReader();
  var responded = false;

  reader.parseBuffer(req.file.buffer, function(err, item) {
    if (responded) return;
    
    if (err) {
      responded = true;
      return res.status(500).json({ error: 'PDF read error: ' + err.message });
    }

    if (!item) {
      responded = true;
      var bank = detectBank(items);
      var transactions = universalParse(items);

      if (transactions.length === 0) {
        return res.status(400).json({ 
          error: 'No transactions found. Bank: ' + bank + 
                 '. Please check if PDF is text-based (not scanned image).' 
        });
      }

      var excelBuffer = buildExcel(transactions, bank);
      res.setHeader('Content-Type', 
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 
        'attachment; filename="statement.xlsx"');
      return res.send(excelBuffer);
    }

    if (item.text) {
      items.push({ 
        text: item.text, 
        x: item.x, 
        y: item.y, 
        page: item.page 
      });
    }
  });
});

var PORT = process.env.PORT || 3000;
app.listen(PORT, function() {
  console.log('BankSync Pro running on port ' + PORT);
});
