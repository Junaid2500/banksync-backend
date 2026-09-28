var express = require('express');
var multer = require('multer');
var PDFParser = require('pdf2json');
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

  var pdfParser = new PDFParser();

  pdfParser.on('pdfParser_dataError', function(err) {
    res.status(500).json({ error: 'PDF parse error: ' + err.parserError });
  });

  pdfParser.on('pdfParser_dataReady', function(pdfData) {
    try {
      var allLines = extractLines(pdfData);
      var bankInfo = detectBank(allLines);
      var transactions = parseTransactions(allLines, bankInfo.bank);

      if (transactions.length === 0) {
        return res.status(400).json({ 
          error: 'No transactions found. Bank: ' + bankInfo.bank + '. Please try another PDF.' 
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
  });

  pdfParser.parseBuffer(req.file.buffer);
});

// Extract text lines from pdf2json data
function extractLines(pdfData) {
  var lines = [];
  
  if (!pdfData.Pages) return lines;

  pdfData.Pages.forEach(function(page) {
    if (!page.Texts) return;

    // Group texts by Y position (same Y = same line)
    var yGroups = {};

    page.Texts.forEach(function(textItem) {
      var y = Math.round(textItem.y * 10) / 10;
      var text = decodeURIComponent(textItem.R[0].T);
      
      if (!yGroups[y]) yGroups[y] = [];
      yGroups[y].push({ x: textItem.x, text: text });
    });

    // Sort by Y, then by X within each line
    var yKeys = Object.keys(yGroups).sort(function(a, b) { return a - b; });

    yKeys.forEach(function(y) {
      var items = yGroups[y].sort(function(a, b) { return a.x - b.x; });
      var lineText = items.map(function(i) { return i.text; }).join(' ');
      var lineItems = items;
      lines.push({ text: lineText.trim(), items: lineItems, y: parseFloat(y) });
    });
  });

  return lines;
}

// Detect bank from text
function detectBank(lines) {
  var fullText = lines.map(function(l) { return l.text; }).join(' ').toUpperCase();
  
  var bank = 'UNKNOWN';
  var account = '';
  var period = '';
  var ifsc = '';
  var micr = '';

  if (fullText.indexOf('BANK OF MAHARASHTRA') !== -1 || fullText.indexOf('MAHABANK') !== -1) bank = 'BOM';
  else if (fullText.indexOf('INDUSIND') !== -1) bank = 'INDUSIND';
  else if (fullText.indexOf('ICICI BANK') !== -1 || fullText.indexOf('ICICI') !== -1) bank = 'ICICI';
  else if (fullText.indexOf('STATE BANK OF INDIA') !== -1) bank = 'SBI';
  else if (fullText.indexOf('HDFC BANK') !== -1) bank = 'HDFC';
  else if (fullText.indexOf('AXIS BANK') !== -1) bank = 'AXIS';
  else if (fullText.indexOf('KOTAK') !== -1) bank = 'KOTAK';
  else if (fullText.indexOf('YES BANK') !== -1) bank = 'YES';
  else if (fullText.indexOf('PUNJAB NATIONAL') !== -1) bank = 'PNB';

  // Extract account number
  var accMatch = fullText.match(/ACCOUNT\s*(?:NO|NUMBER|NO\.)[:\s]*(\d{9,18})/);
  if (accMatch) account = accMatch[1];

  var accMatch2 = fullText.match(/A\/C\s*(?:NO|NUMBER)?[:\s]*(\d{9,18})/);
  if (!account && accMatch2) account = accMatch2[1];

  // Extract period
  var periodMatch = fullText.match(/(\d{2}[\/\-]\d{2}[\/\-]\d{4})\s*(?:TO|to)\s*(\d{2}[\/\-]\d{2}[\/\-]\d{4})/);
  if (periodMatch) period = periodMatch[1] + ' to ' + periodMatch[2];

  var periodMatch2 = fullText.match(/(\d{2}[\/\-][A-Z]{3}[\/\-]\d{4})\s*(?:TO|to)\s*(\d{2}[\/\-][A-Z]{3}[\/\-]\d{4})/i);
  if (!period && periodMatch2) period = periodMatch2[1] + ' to ' + periodMatch2[2];

  // Extract IFSC
  var ifscMatch = fullText.match(/[A-Z]{4}0[A-Z0-9]{6}/);
  if (ifscMatch) ifsc = ifscMatch[0];

  return { bank: bank, account: account, period: period, ifsc: ifsc, micr: micr };
}

// Parse transactions based on bank
function parseTransactions(lines, bank) {
  if (bank === 'BOM') return parseBOM(lines);
  if (bank === 'INDUSIND') return parseIndusInd(lines);
  if (bank === 'ICICI') return parseICICI(lines);
  if (bank === 'SBI') return parseSBI(lines);
  return parseGeneric(lines);
}

// BOM: Date DD/MM/YYYY
function parseBOM(lines) {
  var transactions = [];
  var dateReg = /^\d{2}\/\d{2}\/\d{4}$/;

  for (var i = 0; i < lines.length; i++) {
    var items = lines[i].items;
    if (!items || items.length < 2) continue;

    var firstText = items[0].text.trim();
    if (!firstText.match(dateReg)) continue;

    // Skip header
    if (lines[i].text.toUpperCase().indexOf('DATE') !== -1 && 
        lines[i].text.toUpperCase().indexOf('DEBIT') !== -1) continue;

    var date = firstText;
    var narration = '';
    var chqRef = '';
    var debit = '';
    var credit = '';
    var balance = '';

    // Get all text items on this line
    var allTexts = items.map(function(it) { return it.text.trim(); });
    var nums = [];
    var textParts = [];

    allTexts.forEach(function(t) {
      if (t.match(/^[\d,]+\.\d{2}$/) || t.match(/^-[\d,]+\.\d{2}$/)) {
        nums.push(parseFloat(t.replace(/,/g, '')));
      } else if (t !== date) {
        textParts.push(t);
      }
    });

    // Check next line for continuation narration
    if (i + 1 < lines.length) {
      var nextItems = lines[i+1].items;
      if (nextItems) {
        var nextFirst = nextItems[0] ? nextItems[0].text.trim() : '';
        if (!nextFirst.match(dateReg) && !nextFirst.match(/^\d+$/) && nextFirst.length > 2) {
          var nextNums = [];
          var nextTexts = [];
          nextItems.forEach(function(it) {
            var t = it.text.trim();
            if (t.match(/^[\d,]+\.\d{2}$/)) {
              nextNums.push(parseFloat(t.replace(/,/g, '')));
            } else {
              nextTexts.push(t);
            }
          });
          if (nextNums.length > 0) {
            nums = nums.concat(nextNums);
            textParts = textParts.concat(nextTexts);
            i++;
          }
        }
      }
    }

    narration = textParts.join(' ').trim();

    // BOM columns: Debit | Credit | Balance
    if (nums.length >= 3) {
      debit = nums[0] !== 0 ? nums[0] : '';
      credit = nums[1] !== 0 ? nums[1] : '';
      balance = nums[2];
    } else if (nums.length === 2) {
      var 
