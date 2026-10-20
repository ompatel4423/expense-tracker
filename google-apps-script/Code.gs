// ==========================================
// CONFIGURATION
// ==========================================
const CONFIG = {
  DEFAULT_DAILY_LIMIT: 100,      // Fallback for cell G2
  DEFAULT_MONTHLY_LIMIT: 5000,   // Fallback for cell H2
  // Broad query matching Axis Bank debit alerts, credit card transactions, and AutoPay alerts
  GMAIL_SEARCH_QUERY: '(from:axis.bank.in OR from:axisbank.com) ("Amount Debited" OR "AutoPay" OR "Transaction Amount" OR "Credit Card" OR "spent on" OR "Transaction alert" OR "debited for") -label:ExpensesProcessed',
  PROCESSED_LABEL: 'ExpensesProcessed',
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbz4LphVWeHvjVxweBjqJ7zP4ISgzxeGiuT4zNH_vkeAYizlQG8MpRCChwfrpLEaH4yRoQ/exec'
};

// ==========================================
// WEB APP ROUTER (Dual-Mode: HTML & REST API)
// ==========================================
function doGet(e) {
  // REST API router (bypasses Google Apps Script iframe warning banner)
  if (e && e.parameter && e.parameter.action) {
    const action = e.parameter.action;
    if (action === 'getData') {
      const data = getDashboardData();
      return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (action === 'updateDailyLimit') {
      const success = updateDailyLimit(e.parameter.val);
      return ContentService.createTextOutput(JSON.stringify({ success: success, dailyLimit: e.parameter.val }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (action === 'updateMonthlyLimit') {
      const success = updateMonthlyLimit(e.parameter.val);
      return ContentService.createTextOutput(JSON.stringify({ success: success, monthlyLimit: e.parameter.val }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Native Apps Script HTML delivery
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Expense Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function getDashboardData() {
  const cache = CacheService.getScriptCache();
  const cachedJson = cache.get("dashboard_data_v1");
  if (cachedJson) {
    try {
      return JSON.parse(cachedJson);
    } catch (e) {}
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    initializeSheetHeaders(sheet);
  }
  
  // Read Daily Limit and Monthly Budget in a single batch read
  let dailyLimit = CONFIG.DEFAULT_DAILY_LIMIT;
  let monthlyLimit = CONFIG.DEFAULT_MONTHLY_LIMIT;
  try {
    const limits = sheet.getRange("G2:H2").getValues()[0];
    if (limits && limits[0] && !isNaN(limits[0])) dailyLimit = Number(limits[0]);
    if (limits && limits[1] && !isNaN(limits[1])) monthlyLimit = Number(limits[1]);
  } catch(e) {}
  
  const TIMEZONE = 'Asia/Kolkata';
  const todayStr = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  
  let todayTotal = 0;
  let transactions = [];

  if (lastRow > 1) {
    // Read data up to 6 columns (Date, Time, Amount, Merchant, Card/Mode, Subject)
    const colCount = Math.min(sheet.getLastColumn(), 6);
    const data = sheet.getRange(2, 1, lastRow - 1, Math.max(colCount, 4)).getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      let dateStr = "";
      if (row[0] instanceof Date) {
        dateStr = Utilities.formatDate(row[0], TIMEZONE, 'yyyy-MM-dd');
      } else {
        const str = String(row[0]).trim();
        const dmy = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
        if (dmy) {
          dateStr = `${dmy[3]}-${dmy[2].padStart(2, '0')}-${dmy[1].padStart(2, '0')}`;
        } else {
          dateStr = str.substring(0, 10);
        }
      }

      let timeStr = "";
      if (row[1] instanceof Date) {
        timeStr = Utilities.formatDate(row[1], TIMEZONE, 'hh:mm a');
      } else if (row[1]) {
        const str = String(row[1]);
        const m = str.match(/(\d{1,2}):(\d{2})(?::\d{2})?/);
        if (m) {
          let h = parseInt(m[1], 10);
          const min = m[2];
          const ampm = h >= 12 ? 'PM' : 'AM';
          h = h % 12 || 12;
          timeStr = `${h}:${min} ${ampm}`;
        } else {
          timeStr = str;
        }
      }

      const amount = Number(row[2]);
      const merchant = row[3] || 'Axis Debit';

      // Determine card / payment channel from column E
      let card = 'Axis Direct';
      if (row[4]) {
        const val = String(row[4]).trim();
        if (val.length > 0 && val.length < 35 && !val.toLowerCase().includes('subject:') && !val.toLowerCase().includes('alert for')) {
          card = val;
        } else if (/XX\d{4}/i.test(val)) {
          const m = val.match(/XX\d{4}/i);
          card = 'Axis ' + m[0];
        }
      }

      transactions.unshift({
        date: dateStr,
        time: timeStr,
        amount: amount,
        merchant: merchant,
        card: card
      });
    }
  }

  // Deduplicate transactions (e.g. YouTube ₹89, Apple ₹249, Equal AI duplicate bank notifications)
  const cleanTransactions = deduplicateTransactionsList(transactions);

  // If duplicate rows exist in the sheet, automatically purge them in the background
  if (cleanTransactions.length < transactions.length) {
    try {
      cleanDuplicateTransactions(sheet, false);
    } catch(e) {
      console.warn("Auto background cleanup:", e);
    }
  }

  for (let i = 0; i < cleanTransactions.length; i++) {
    const t = cleanTransactions[i];
    if (t.date === todayStr && !isNaN(t.amount)) {
      todayTotal += t.amount;
    }
  }

  const result = {
    todayTotal: todayTotal,
    dailyLimit: dailyLimit,
    monthlyLimit: monthlyLimit,
    todayStr: todayStr,
    transactions: cleanTransactions
  };

  try {
    const jsonStr = JSON.stringify(result);
    if (jsonStr.length < 100000) {
      cache.put("dashboard_data_v1", jsonStr, 60);
    }
  } catch(e) {}

  return result;
}

// Update Daily Limit (Cell G2)
function updateDailyLimit(newLimit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const val = Number(newLimit);
  if (!isNaN(val) && val > 0) {
    sheet.getRange("G2").setValue(val);
    try { CacheService.getScriptCache().remove("dashboard_data_v1"); } catch(e) {}
    return true;
  }
  return false;
}

// Update Monthly Budget (Cell H2)
function updateMonthlyLimit(newLimit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const val = Number(newLimit);
  if (!isNaN(val) && val > 0) {
    sheet.getRange("H2").setValue(val);
    try { CacheService.getScriptCache().remove("dashboard_data_v1"); } catch(e) {}
    return true;
  }
  return false;
}

// ==========================================
// MAIN WORKFLOW TRIGGER
// ==========================================
function trackExpensesAndCheckLimit() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  initializeSheetHeaders(sheet);
  
  const TIMEZONE = 'Asia/Kolkata';
  const processedLabel = getOrCreateLabel(CONFIG.PROCESSED_LABEL);
  const threads = GmailApp.search(CONFIG.GMAIL_SEARCH_QUERY, 0, 25);
  let newTransactionsCount = 0;

  // Load recent rows from the sheet to detect and prevent duplicate bank alerts
  const recentRows = [];
  const totalRows = sheet.getLastRow();
  if (totalRows > 1) {
    const startRow = Math.max(2, totalRows - 120);
    const numRows = totalRows - startRow + 1;
    const colCount = Math.min(sheet.getLastColumn(), 6);
    const sheetVals = sheet.getRange(startRow, 1, numRows, Math.max(colCount, 6)).getValues();
    for (let r = 0; r < sheetVals.length; r++) {
      const row = sheetVals[r];
      if (!row[0]) continue;
      let dStr = "";
      if (row[0] instanceof Date) {
        dStr = Utilities.formatDate(row[0], TIMEZONE, 'yyyy-MM-dd');
      } else {
        const str = String(row[0]).trim();
        const dmy = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
        if (dmy) {
          dStr = `${dmy[3]}-${dmy[2].padStart(2, '0')}-${dmy[1].padStart(2, '0')}`;
        } else {
          dStr = str.substring(0, 10);
        }
      }
      let tStr = row[1] instanceof Date ? Utilities.formatDate(row[1], TIMEZONE, 'HH:mm:ss') : String(row[1] || '');
      recentRows.push({
        rowIndex: startRow + r,
        date: dStr,
        time: tStr,
        amount: Number(row[2]),
        merchant: String(row[3] || ''),
        card: String(row[4] || ''),
        subject: String(row[5] || '')
      });
    }
  }

  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const message of messages) {
      const body = message.getPlainBody();
      const subject = message.getSubject() || '';
      const date = message.getDate();
      const transaction = parseAxisEmail(body, subject);
      
      if (transaction && transaction.amount > 0) {
        let dateVal = Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd');
        let timeVal = Utilities.formatDate(date, TIMEZONE, 'HH:mm:ss');
        
        // Check if email body contains explicit Date & Time (e.g. "17-10-2026, 08:47:13 IST")
        const dtMatch = body.match(/Date & Time:?[\s\r\n\t]*(\d{2})[-/](\d{2})[-/](\d{4}),?\s*(\d{2}):(\d{2}):(\d{2})/i);
        if (dtMatch) {
          dateVal = `${dtMatch[3]}-${dtMatch[2]}-${dtMatch[1]}`;
          timeVal = `${dtMatch[4]}:${dtMatch[5]}:${dtMatch[6]}`;
        } else {
          // Check top-level date like "19-10-2026" or "17-10-2026"
          const dMatch = body.match(/^(\d{2})[-/](\d{2})[-/](\d{4})/m);
          if (dMatch) {
            dateVal = `${dMatch[3]}-${dMatch[2]}-${dMatch[1]}`;
          }
        }

        const candidateTx = {
          date: dateVal,
          time: timeVal,
          amount: transaction.amount,
          merchant: transaction.merchant,
          card: transaction.card || 'Axis Direct',
          subject: subject
        };

        // Check for duplicates among recent rows or rows added in this batch
        let existingMatch = null;
        for (let r = 0; r < recentRows.length; r++) {
          if (areDuplicateTransactions(recentRows[r], candidateTx)) {
            existingMatch = recentRows[r];
            break;
          }
        }

        if (existingMatch) {
          // Duplicate alert detected! Check if new candidateTx provides better details
          const better = pickBetterTransaction(existingMatch, candidateTx);
          if (better === candidateTx) {
            // Update the existing row with richer merchant name, card digits, etc.
            sheet.getRange(existingMatch.rowIndex, 1, 1, 6).setValues([[
              dateVal,
              timeVal,
              candidateTx.amount,
              candidateTx.merchant,
              candidateTx.card,
              candidateTx.subject
            ]]);
            existingMatch.merchant = candidateTx.merchant;
            existingMatch.card = candidateTx.card;
            existingMatch.subject = candidateTx.subject;
            existingMatch.time = candidateTx.time;
            newTransactionsCount++;
          }
          // Else: existing row is already equal or higher quality; ignore redundant notification
        } else {
          // No duplicate found: append as a new transaction row
          sheet.appendRow([
            dateVal,
            timeVal,
            candidateTx.amount,
            candidateTx.merchant,
            candidateTx.card,
            candidateTx.subject
          ]);
          recentRows.push({
            rowIndex: sheet.getLastRow(),
            date: dateVal,
            time: timeVal,
            amount: candidateTx.amount,
            merchant: candidateTx.merchant,
            card: candidateTx.card,
            subject: candidateTx.subject
          });
          newTransactionsCount++;
        }
      }
    }
    thread.addLabel(processedLabel);
  }

  // Automatically sweep and purge duplicate notifications from sheet
  try {
    cleanDuplicateTransactions(sheet, false);
  } catch(e) {
    console.warn("Auto cleanup in sync:", e);
  }

  if (newTransactionsCount > 0) {
    try { CacheService.getScriptCache().remove("dashboard_data_v1"); } catch(e) {}
  }

  checkDailyBudgetLimit(sheet);
}

// ==========================================
// MULTI-CHANNEL BANK & CARD EMAIL PARSER
// ==========================================
function parseAxisEmail(body, subject) {
  if (!body) return null;

  // 1. Structured Summary Email (Axis Bank Credit Card OR AutoPay transactions)
  // Handles:
  // - "Here's the summary of your Axis Bank Credit Card Transaction:"
  // - "Here's the summary of your successful AutoPay transaction:"
  // - Any email with "Transaction Amount:" and "Merchant Name:"
  if (/summary of your.*(?:Credit Card|AutoPay|transaction)/i.test(body) || (/Transaction Amount:/i.test(body) && /Merchant Name:/i.test(body))) {
    const amtMatch = body.match(/Transaction Amount:?[\s\r\n\t]*(?:INR|Rs\.?)\s*([\d,]+(?:\.\d{1,2})?)/i) ||
                     body.match(/(?:INR|Rs\.?)\s*([\d,]+(?:\.\d{1,2})?)/i);
    const merchMatch = body.match(/Merchant Name:?[\s\r\n\t]*([^\r\n]+)/i);
    const cardMatch = body.match(/(?:Axis Bank Credit Card No\.?|Axis Bank Card No\.?|Card No\.?|Card ending)[\s\r\n\t]*([A-Za-z0-9]+)/i);

    if (amtMatch) {
      const amount = parseFloat(amtMatch[1].replace(/,/g, ''));
      if (!isNaN(amount) && amount > 0) {
        const merchant = merchMatch ? merchMatch[1].trim() : 'Axis Transaction';
        const cardNum = cardMatch ? cardMatch[1].trim() : '';
        const isCC = /Credit Card/i.test(body) || /Credit Card/i.test(subject);
        const cardPrefix = isCC ? 'Axis CC' : 'Axis Card';
        const card = cardNum ? (cardNum.startsWith('XX') ? `${cardPrefix} ${cardNum}` : `${cardPrefix} XX${cardNum}`) : cardPrefix;
        return { amount: amount, merchant: merchant, card: card, type: isCC ? 'Credit Card' : 'AutoPay' };
      }
    }
  }

  // 2. Inline Credit Card Transaction Alert (Swipes, Online POS, E-commerce)
  if (/Credit Card|Card no\.|ending with/i.test(body) || /Credit Card/i.test(subject)) {
    let amtMatch = body.match(/(?:INR|Rs\.?)\s*([\d,]+(?:\.\d{1,2})?)\s*(?:has been spent|was spent|spent|debited)/i) ||
                   body.match(/(?:spent|debited)\s*(?:for|with|by)?\s*(?:INR|Rs\.?)\s*([\d,]+(?:\.\d{1,2})?)/i) ||
                   body.match(/Transaction Amount:?[\s\r\n\t]*(?:INR|Rs\.?)\s*([\d,]+(?:\.\d{1,2})?)/i);

    let merchMatch = body.match(/at\s+([A-Za-z0-9\s.,&'-]+?)(?:\s+on\s+|\s*\.|\s*Available|\s*Avbl|\s*Total|\s*Ref|\s*$)/i) ||
                     body.match(/Merchant Name:?[\s\r\n\t]*([^\r\n]+)/i);

    let cardMatch = body.match(/(?:Credit Card|Card)\s*(?:no\.?|ending with|ending)?\s*([X\d]{4,16})/i);

    if (amtMatch) {
      const amount = parseFloat(amtMatch[1].replace(/,/g, ''));
      if (!isNaN(amount) && amount > 0) {
        let merchant = merchMatch ? merchMatch[1].trim() : 'Credit Card';
        merchant = merchant.replace(/\s+/g, ' ');
        const cardNum = cardMatch ? cardMatch[1].trim() : '';
        const card = cardNum ? (cardNum.startsWith('XX') ? `Axis CC ${cardNum}` : `Axis CC XX${cardNum}`) : 'Axis CC';
        return { amount: amount, merchant: merchant, card: card, type: 'Credit Card' };
      }
    }
  }

  // 3. Bank Account / UPI / Debit Alert
  let amtMatch = body.match(/Amount Debited:?[\s\r\n\t]*(?:INR|Rs\.?)\s*([\d,]+\.?\d{0,2})/i) ||
                 body.match(/(?:INR|Rs\.?)\s*([\d,]+\.?\d{0,2})\s*(?:has been debited|is debited|debited)/i) ||
                 body.match(/debited\s*(?:for|with|by)\s*(?:INR|Rs\.?)\s*([\d,]+\.?\d{0,2})/i) ||
                 body.match(/transaction of\s*(?:INR|Rs\.?)\s*([\d,]+\.?\d{0,2})/i);

  if (amtMatch) {
    const amount = parseFloat(amtMatch[1].replace(/,/g, ''));
    if (!isNaN(amount) && amount > 0) {
      let merchant = 'Axis Debit';
      const infoMatch = body.match(/Transaction Info:?[\s\r\n\t]*([^\r\n]+)/i);
      if (infoMatch) {
        let rawInfo = infoMatch[1].trim();
        rawInfo = rawInfo.replace(/\.?\s*(?:Available\s*)?Balance:?.*$/i, '').trim();
        const parts = rawInfo.split('/');
        merchant = (parts.length > 1 && parts[parts.length - 1].trim() !== '') 
          ? parts[parts.length - 1].trim() 
          : rawInfo;
      }
      return { amount: amount, merchant: merchant, card: 'Axis A/C', type: 'Bank Debit' };
    }
  }

  return null;
}

// ==========================================
// DAILY LIMIT CHECKER (With Accurate Receipt)
// ==========================================
function checkDailyBudgetLimit(sheet) {
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const data = sheet.getDataRange().getValues();
  
  const cellLimit = sheet.getRange("G2").getValue();
  const dailyLimit = (cellLimit && !isNaN(cellLimit)) ? Number(cellLimit) : CONFIG.DEFAULT_DAILY_LIMIT;

  let todayTotal = 0;
  let todayTx = [];

  for (let i = 1; i < data.length; i++) {
    const rawDate = data[i][0];
    if (!rawDate) continue;

    const rowDate = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const amount = Number(data[i][2]);
    const merchant = data[i][3] || 'Axis Debit';
    const time = data[i][1] ? data[i][1].toString() : '';
    const card = data[i][4] ? String(data[i][4]).trim() : 'Axis Direct';

    if (rowDate === todayStr && !isNaN(amount)) {
      todayTotal += amount;
      todayTx.push({ amount: amount, merchant: merchant, time: time, card: card });
    }
  }

  const userProperties = PropertiesService.getUserProperties();
  const lastAlertDate = userProperties.getProperty('LAST_ALERT_DATE');

  if (todayTotal > dailyLimit && lastAlertDate !== todayStr) {
    dispatchKageEmailAlert(todayTotal, dailyLimit, todayTx);
    userProperties.setProperty('LAST_ALERT_DATE', todayStr);
  }
}

function dispatchKageEmailAlert(spent, limit, transactions) {
  const recipient = Session.getActiveUser().getEmail();
  const overAmount = (spent - limit) > 0 ? (spent - limit) : 0;
  const subject = `[BUDGET ALERT] Limit Exceeded: Spent ₹${spent.toFixed(2)} (Cap: ₹${limit.toFixed(2)})`;

  let txHtml = '';
  if (transactions && transactions.length > 0) {
    txHtml = transactions.map(t => `
      <tr>
        <td style="padding: 10px 0; border-bottom: 1px solid #1a1e27; color: #f1f5f9; font-size: 14px; font-weight: 500;">
          ${t.merchant}
          <div style="font-size: 11px; color: #717d96; margin-top: 2px;">${t.card ? t.card + ' • ' : ''}${t.time}</div>
        </td>
        <td style="padding: 10px 0; border-bottom: 1px solid #1a1e27; text-align: right; color: #ff6b64; font-size: 14px; font-weight: 600;">
          -₹${Number(t.amount).toFixed(2)}
        </td>
      </tr>
    `).join('');
  } else {
    txHtml = `<tr><td colspan="2" style="padding: 12px 0; color: #717d96; font-size: 13px;">Transactions recorded in dashboard.</td></tr>`;
  }

  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head><meta charset="utf-8"></head>
    <body style="margin: 0; padding: 24px 0; background-color: #050608; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;">
      <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 520px; background-color: #0d0f15; border: 1px solid #1e2330; border-radius: 20px; overflow: hidden; box-shadow: 0 20px 40px rgba(0,0,0,0.8);">
        <tr>
          <td style="padding: 32px 32px 20px; border-bottom: 1px solid #171b24;">
            <div style="display: inline-block; padding: 5px 12px; border-radius: 999px; background-color: rgba(224, 35, 28, 0.12); border: 1px solid rgba(224, 35, 28, 0.3); color: #ff5247; font-size: 11px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase;">
              ● BUDGET BREACH
            </div>
            <h1 style="color: #ffffff; font-size: 24px; font-weight: 600; margin: 16px 0 6px; letter-spacing: -0.02em;">Daily Limit Exceeded</h1>
            <p style="color: #8b95a8; font-size: 14px; margin: 0; line-height: 1.5;">
              Your expenses today have surpassed your configured ceiling of <strong style="color: #ffffff;">₹${limit.toFixed(2)}</strong>.
            </p>
          </td>
        </tr>
        <tr>
          <td style="padding: 24px 32px;">
            <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #12151e; border: 1px solid #232838; border-radius: 14px; padding: 18px 20px;">
              <tr>
                <td style="vertical-align: top;">
                  <div style="font-size: 11px; text-transform: uppercase; color: #717d96; font-weight: 600;">Spent Today</div>
                  <div style="font-size: 28px; font-weight: 700; color: #ff4757; margin-top: 4px;">₹${spent.toFixed(2)}</div>
                </td>
                <td style="vertical-align: top; text-align: right;">
                  <div style="font-size: 11px; text-transform: uppercase; color: #717d96; font-weight: 600;">Daily Limit</div>
                  <div style="font-size: 16px; font-weight: 600; color: #ffffff; margin-top: 4px;">₹${limit.toFixed(2)}</div>
                  <div style="font-size: 12px; color: #ff6b64; margin-top: 2px;">+₹${overAmount.toFixed(2)} over</div>
                </td>
              </tr>
            </table>
            <div style="margin-top: 24px;">
              <div style="font-size: 12px; font-weight: 600; text-transform: uppercase; color: #717d96; margin-bottom: 8px;">Today's Activity</div>
              <table width="100%" cellpadding="0" cellspacing="0">${txHtml}</table>
            </div>
            <div style="margin-top: 32px; text-align: center;">
              <a href="${CONFIG.WEB_APP_URL}" style="display: block; background: #e0231c; color: #ffffff; text-decoration: none; padding: 14px 28px; border-radius: 12px; font-size: 15px; font-weight: 600;">Open Expense Dashboard &rarr;</a>
            </div>
          </td>
        </tr>
      </table>
    </body>
    </html>
  `;

  GmailApp.sendEmail(recipient, subject, "", { htmlBody: htmlBody });
}

function initializeSheetHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date', 'Time', 'Amount (INR)', 'Merchant/Info', 'Card/Mode', 'Email Subject']);
    sheet.getRange("A1:F1").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  
  // Set up Daily Limit in G1:G2
  if (sheet.getRange("G1").getValue() === "") {
    sheet.getRange("G1").setValue("Daily Limit (₹)").setFontWeight("bold");
    sheet.getRange("G2").setValue(CONFIG.DEFAULT_DAILY_LIMIT).setFontColor("#2b6cb0").setFontWeight("bold");
  }

  // Set up Monthly Budget in H1:H2
  if (sheet.getRange("H1").getValue() === "") {
    sheet.getRange("H1").setValue("Monthly Budget (₹)").setFontWeight("bold");
    sheet.getRange("H2").setValue(CONFIG.DEFAULT_MONTHLY_LIMIT).setFontColor("#2b6cb0").setFontWeight("bold");
  }
}

function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

// ==========================================
// GOOGLE SHEETS CUSTOM MENU
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💳 Expense Tracker')
    .addItem('⚡ Sync Latest Emails', 'trackExpensesAndCheckLimit')
    .addItem('🧹 Clean Up Duplicates', 'menuCleanDuplicates')
    .addSeparator()
    .addItem('📦 Archive Previous Years', 'autoArchivePreviousYears')
    .addToUi();
}

// ==========================================
// AUTO-ARCHIVE ENGINE (Maintains Speed Forever)
// ==========================================
function autoArchivePreviousYears() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const lastRow = activeSheet.getLastRow();
  if (lastRow <= 1) return;

  const currentYear = new Date().getFullYear();
  const colCount = Math.min(activeSheet.getLastColumn(), 6);
  const data = activeSheet.getRange(2, 1, lastRow - 1, Math.max(colCount, 4)).getValues();

  const toArchive = {}; // { '2025': [ [row], [row] ] }
  const rowsToDelete = [];

  for (let i = 0; i < data.length; i++) {
    const rawDate = data[i][0];
    if (!rawDate) continue;

    let rowYear = null;
    if (rawDate instanceof Date) {
      rowYear = rawDate.getFullYear();
    } else {
      const str = String(rawDate).trim();
      const m = str.match(/^(\d{4})/);
      if (m) rowYear = parseInt(m[1], 10);
    }

    // Move any past year to its archive tab
    if (rowYear && rowYear < currentYear) {
      if (!toArchive[rowYear]) toArchive[rowYear] = [];
      toArchive[rowYear].push(data[i]);
      rowsToDelete.push(i + 2); // 1-indexed (row 1 is header)
    }
  }

  // Populate year archive sheets
  for (const year in toArchive) {
    const sheetName = `Archive_${year}`;
    let archiveSheet = ss.getSheetByName(sheetName);
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet(sheetName);
      archiveSheet.appendRow(['Date', 'Time', 'Amount (INR)', 'Merchant/Info', 'Card/Mode', 'Email Subject']);
      archiveSheet.getRange("A1:F1").setFontWeight("bold");
      archiveSheet.setFrozenRows(1);
    }
    const rows = toArchive[year];
    if (rows.length > 0) {
      archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
  }

  // Delete archived rows from active sheet in reverse order to preserve row indices
  for (let r = rowsToDelete.length - 1; r >= 0; r--) {
    activeSheet.deleteRow(rowsToDelete[r]);
  }

  if (rowsToDelete.length > 0) {
    try { CacheService.getScriptCache().remove("dashboard_data_v1"); } catch(e) {}
  }
}

// ==========================================
// DEDUPLICATION & MULTI-ALERT MERGE ENGINE
// ==========================================

const KNOWN_SUBSCRIPTION_KEYWORDS = [
  'apple', 'appleservic', 'youtube', 'google', 'equalai', 'equal ai',
  'netflix', 'spotify', 'amazon prime', 'hotstar', 'patreon', 'github',
  'openai', 'chatgpt', 'claude', 'midjourney', 'cursor', 'vercel'
];

function normalizeMerchant(name) {
  if (!name) return '';
  return String(name)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function parseTimeToMinutes(timeStr) {
  if (!timeStr) return null;
  const str = String(timeStr).trim();
  const m12 = str.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)/i);
  if (m12) {
    let h = parseInt(m12[1], 10);
    const m = parseInt(m12[2], 10);
    const isPm = m12[4].toLowerCase() === 'pm';
    if (isPm && h < 12) h += 12;
    if (!isPm && h === 12) h = 0;
    return h * 60 + m;
  }
  const m24 = str.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (m24) {
    let h = parseInt(m24[1], 10);
    const m = parseInt(m24[2], 10);
    return h * 60 + m;
  }
  return null;
}

function isSubscriptionMerchant(merchant) {
  const norm = normalizeMerchant(merchant);
  for (let i = 0; i < KNOWN_SUBSCRIPTION_KEYWORDS.length; i++) {
    const kwNorm = normalizeMerchant(KNOWN_SUBSCRIPTION_KEYWORDS[i]);
    if (norm.includes(kwNorm) || kwNorm.includes(norm)) {
      return true;
    }
  }
  return false;
}

function areDuplicateTransactions(tx1, tx2) {
  // 1. Amount must be identical (within 0.01)
  const amt1 = Number(tx1.amount);
  const amt2 = Number(tx2.amount);
  if (isNaN(amt1) || isNaN(amt2) || Math.abs(amt1 - amt2) > 0.01) {
    return false;
  }

  // 2. Dates must match
  if (tx1.date !== tx2.date) {
    return false;
  }

  // 3. Merchant normalization & comparison
  const m1 = normalizeMerchant(tx1.merchant);
  const m2 = normalizeMerchant(tx2.merchant);

  let merchantMatches = false;
  if (m1 === m2) {
    merchantMatches = true;
  } else if (m1.length >= 4 && m2.length >= 4) {
    if (m1.includes(m2) || m2.includes(m1)) {
      merchantMatches = true;
    } else {
      let prefixLen = 0;
      const minLen = Math.min(m1.length, m2.length);
      while (prefixLen < minLen && m1[prefixLen] === m2[prefixLen]) {
        prefixLen++;
      }
      if (prefixLen >= 5) {
        merchantMatches = true;
      }
    }
  }

  if (!merchantMatches) {
    for (let i = 0; i < KNOWN_SUBSCRIPTION_KEYWORDS.length; i++) {
      const kwNorm = normalizeMerchant(KNOWN_SUBSCRIPTION_KEYWORDS[i]);
      if (m1.includes(kwNorm) && m2.includes(kwNorm)) {
        merchantMatches = true;
        break;
      }
    }
  }

  const isGeneric1 = /^(axis|axisdebit|creditcard|axistransaction)$/i.test(m1);
  const isGeneric2 = /^(axis|axisdebit|creditcard|axistransaction)$/i.test(m2);
  if ((isGeneric1 && !isGeneric2) || (isGeneric2 && !isGeneric1)) {
    const t1 = parseTimeToMinutes(tx1.time);
    const t2 = parseTimeToMinutes(tx2.time);
    if (t1 !== null && t2 !== null && Math.abs(t1 - t2) <= 15) {
      merchantMatches = true;
    }
  }

  if (!merchantMatches) {
    return false;
  }

  // 4. Timing check
  const t1 = parseTimeToMinutes(tx1.time);
  const t2 = parseTimeToMinutes(tx2.time);

  if (t1 !== null && t2 !== null) {
    const diff = Math.abs(t1 - t2);
    if (diff <= 20) {
      return true;
    }
    if (isSubscriptionMerchant(tx1.merchant) || isSubscriptionMerchant(tx2.merchant)) {
      return true;
    }
    return false;
  }

  if (isSubscriptionMerchant(tx1.merchant) || isSubscriptionMerchant(tx2.merchant)) {
    return true;
  }

  return true;
}

function pickBetterTransaction(tx1, tx2) {
  let score1 = 0;
  let score2 = 0;

  const m1 = tx1.merchant || '';
  const m2 = tx2.merchant || '';

  if (m1.length > m2.length) score1 += 2;
  if (m2.length > m1.length) score2 += 2;

  if (/SERVIC$/i.test(m1)) score1 -= 3;
  if (/SERVIC$/i.test(m2)) score2 -= 3;

  if (/XX\d{4}/i.test(tx1.card || '')) score1 += 3;
  if (/XX\d{4}/i.test(tx2.card || '')) score2 += 3;

  if (tx1.time && tx1.time.length >= 5) score1 += 1;
  if (tx2.time && tx2.time.length >= 5) score2 += 1;

  return score2 > score1 ? tx2 : tx1;
}

function deduplicateTransactionsList(txList) {
  if (!txList || txList.length <= 1) return txList || [];
  const toSkip = new Set();
  const clean = [];

  for (let i = 0; i < txList.length; i++) {
    if (toSkip.has(i)) continue;
    let currentBest = txList[i];

    for (let j = i + 1; j < txList.length; j++) {
      if (toSkip.has(j)) continue;
      if (areDuplicateTransactions(currentBest, txList[j])) {
        toSkip.add(j);
        currentBest = pickBetterTransaction(currentBest, txList[j]);
      }
    }
    clean.push(currentBest);
  }

  return clean;
}

function menuCleanDuplicates() {
  cleanDuplicateTransactions(null, true);
}

function cleanDuplicateTransactions(targetSheet, isInteractive) {
  isInteractive = (isInteractive === true);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = targetSheet || (ss ? ss.getActiveSheet() : null);
  if (!sheet) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    if (isInteractive && ss) ss.toast("No transactions found in sheet.", "Deduplication Complete", 5);
    return 0;
  }

  const TIMEZONE = 'Asia/Kolkata';
  const colCount = Math.min(sheet.getLastColumn(), 6);
  const data = sheet.getRange(2, 1, lastRow - 1, Math.max(colCount, 6)).getValues();

  const rows = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    let dStr = "";
    if (r[0] instanceof Date) {
      dStr = Utilities.formatDate(r[0], TIMEZONE, 'yyyy-MM-dd');
    } else {
      const str = String(r[0]).trim();
      const dmy = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
      if (dmy) {
        dStr = `${dmy[3]}-${dmy[2].padStart(2, '0')}-${dmy[1].padStart(2, '0')}`;
      } else {
        dStr = str.substring(0, 10);
      }
    }
    let tStr = r[1] instanceof Date ? Utilities.formatDate(r[1], TIMEZONE, 'HH:mm:ss') : String(r[1] || '');
    rows.push({
      rowIndex: i + 2,
      date: dStr,
      time: tStr,
      amount: Number(r[2]),
      merchant: String(r[3] || ''),
      card: String(r[4] || ''),
      subject: String(r[5] || '')
    });
  }

  const toDeleteRowIndices = new Set();
  let mergedCount = 0;

  for (let i = 0; i < rows.length; i++) {
    if (toDeleteRowIndices.has(rows[i].rowIndex)) continue;
    for (let j = i + 1; j < rows.length; j++) {
      if (toDeleteRowIndices.has(rows[j].rowIndex)) continue;

      if (areDuplicateTransactions(rows[i], rows[j])) {
        const better = pickBetterTransaction(rows[i], rows[j]);
        if (better === rows[j]) {
          sheet.getRange(rows[i].rowIndex, 1, 1, 6).setValues([[
            rows[j].date,
            rows[j].time,
            rows[j].amount,
            rows[j].merchant,
            rows[j].card,
            rows[j].subject
          ]]);
          rows[i].merchant = rows[j].merchant;
          rows[i].card = rows[j].card;
          rows[i].subject = rows[j].subject;
          rows[i].time = rows[j].time;
        }
        toDeleteRowIndices.add(rows[j].rowIndex);
        mergedCount++;
      }
    }
  }

  const sortedDelete = Array.from(toDeleteRowIndices).sort((a, b) => b - a);
  for (let d = 0; d < sortedDelete.length; d++) {
    sheet.deleteRow(sortedDelete[d]);
  }

  if (mergedCount > 0) {
    try {
      CacheService.getScriptCache().remove("dashboard_data_v1");
    } catch(e) {}
  }

  if (isInteractive && ss) {
    const msg = mergedCount > 0 
      ? `Successfully merged and removed ${mergedCount} duplicate transaction(s).`
      : `No duplicate transactions found. Ledger is already clean!`;
    ss.toast(msg, "Deduplication Complete", 6);
    try {
      SpreadsheetApp.getUi().alert("Deduplication Complete", msg, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch(e) {}
  }

  return mergedCount;
}