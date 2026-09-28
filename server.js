function parseICICI(lines) {
  var headerIdx = findHeaderLine(lines, ['TRANSACTION', 'DESCRIPTION', 'WITHDRAWAL']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['TRANSACTION', 'DEPOSIT', 'BALANCE']);
  if (headerIdx === -1) headerIdx = findHeaderLine(lines, ['DATE', 'DESCRIPTION', 'BALANCE']);
  if (headerIdx === -1) headerIdx = 0;

  var transactions = [];
  
  // DD-Mon-YYYY format: 01-May-2026
  var dateRegex = /^\d{2}-[A-Za-z]{3}-\d{4}$/;
  var amountRegex = /^[\d,]+\.\d{2}$/;

  var i = headerIdx + 1;
  while (i < lines.length) {
    var text = lines[i].text.trim();
    
    if (dateRegex.test(text)) {
      var date = text;
      var narration = '';
      var refNo = '';
      var chequeNo = '';
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
        } else if (/^[A-Z]\d{6,}$/.test(t)) {
          // Transaction ID like S8471070, M337293
          refNo = t;
        } else if (t.length > 2 && t !== 'Dr' && t !== 'Cr') {
          narration = narration ? narration + ' ' + t : t;
        }
        i++;
      }

      if (balance) {
        transactions.push({
          date: date,
          narration: narration,
          refNo: refNo,
          chequeNo: chequeNo,
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
