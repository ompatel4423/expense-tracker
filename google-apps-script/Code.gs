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
  
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
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
        dateStr = Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        dateStr = String(row[0]).substring(0, 10);
      }

      const timeStr = row[1] ? row[1].toString() : '';
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

      if (dateStr === todayStr && !isNaN(amount)) {
        todayTotal += amount;
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

  const result = {
    todayTotal: todayTotal,
    dailyLimit: dailyLimit,
    monthlyLimit: monthlyLimit,
    todayStr: todayStr,
    transactions: transactions
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
  
  const processedLabel = getOrCreateLabel(CONFIG.PROCESSED_LABEL);
  const threads = GmailApp.search(CONFIG.GMAIL_SEARCH_QUERY, 0, 25);
  let newTransactionsCount = 0;

  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const message of messages) {
      const body = message.getPlainBody();
      const subject = message.getSubject() || '';
      const date = message.getDate();
      const transaction = parseAxisEmail(body, subject);
      
      if (transaction && transaction.amount > 0) {
        sheet.appendRow([
          Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          Utilities.formatDate(date, Session.getScriptTimeZone(), 'HH:mm:ss'),
          transaction.amount,
          transaction.merchant,
          transaction.card || 'Axis Direct',
          subject
        ]);
        newTransactionsCount++;
      }
    }
    thread.addLabel(processedLabel);
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

  // 1. AutoPay Transaction (Subscriptions, Recurring debits e.g. Apple, Netflix, SIHub)
  if (/AutoPay/i.test(body) || /AutoPay/i.test(subject)) {
    const amtMatch = body.match(/Transaction Amount:?[\s\r\n\t]*(?:INR|Rs\.?)\s*([\d,]+\.?\d{0,2})/i) ||
                     body.match(/(?:INR|Rs\.?)\s*([\d,]+\.?\d{0,2})/i);
    const merchMatch = body.match(/Merchant Name:?[\s\r\n\t]*([^\r\n]+)/i);
    const cardMatch = body.match(/(?:Axis Bank Card No\.?|Card No\.?|Card ending)[\s\r\n\t]*([A-Za-z0-9]+)/i);

    if (amtMatch) {
      const amount = parseFloat(amtMatch[1].replace(/,/g, ''));
      if (!isNaN(amount) && amount > 0) {
        const merchant = merchMatch ? merchMatch[1].trim() : 'AutoPay Debit';
        const cardNum = cardMatch ? cardMatch[1].trim() : '';
        const card = cardNum ? (cardNum.startsWith('XX') ? `Axis Card ${cardNum}` : `Axis Card XX${cardNum}`) : 'Axis Card';
        return { amount: amount, merchant: merchant, card: card, type: 'AutoPay' };
      }
    }
  }

  // 2. Credit Card Transaction Alert (Swipes, Online POS, E-commerce)
  if (/Credit Card|Card no\.|ending with/i.test(body) || /Credit Card/i.test(subject)) {
    let amtMatch = body.match(/(?:INR|Rs\.?)\s*([\d,]+\.?\d{0,2})\s*(?:has been spent|was spent|spent|debited)/i) ||
                   body.match(/(?:spent|debited)\s*(?:for|with|by)?\s*(?:INR|Rs\.?)\s*([\d,]+\.?\d{0,2})/i) ||
                   body.match(/Transaction Amount:?[\s\r\n\t]*(?:INR|Rs\.?)\s*([\d,]+\.?\d{0,2})/i);

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