# Google Apps Script Backend (Axis Bank Expense Tracker)

This directory contains the Google Apps Script backend code that powers the Expense Tracker.

## Links
- **Google Sheet**: [View Sheet](https://docs.google.com/spreadsheets/d/11d93LcZgoc1fOm1TLDsjT19PK4uw7zSf4juzFaLf_ck/edit?usp=drive_link)
- **Web App API URL**: `https://script.google.com/macros/s/AKfycbz4LphVWeHvjVxweBjqJ7zP4ISgzxeGiuT4zNH_vkeAYizlQG8MpRCChwfrpLEaH4yRoQ/exec`

## Endpoints
- `GET ?action=getData` - Returns JSON payload containing:
  - `todayTotal`: Total spent today (₹)
  - `dailyLimit`: Daily spending threshold (₹)
  - `monthlyLimit`: Monthly budget threshold (₹)
  - `todayStr`: Current ISO date string
  - `transactions`: Array of recorded expenses with `date`, `time`, `amount`, `merchant`, and `card`
- `GET ?action=updateDailyLimit&val=<AMOUNT>` - Updates daily limit cell in the Google Sheet.
- `GET ?action=updateMonthlyLimit&val=<AMOUNT>` - Updates monthly limit cell in the Google Sheet.

## Gmail Automation
The script filters incoming emails across `axis.bank.in` and `axisbank.com`:
- **AutoPay recurring transactions**: Apple, Netflix, Spotify, SIHub subscriptions.
- **Credit Card transactions**: Online purchases (Amazon, YouTube, Flipkart), swiped POS payments.
- **Bank debits**: Standard UPI & NetBanking debits.

## Auto-Archiving
Includes `autoArchivePreviousYears()` to keep the main sheet fast forever:
- Automatically groups transactions from previous calendar years into tabs named `Archive_<YEAR>`.
- Preserves headers and cleans up the active tracking sheet.
- Can be run on demand from the custom menu: **💳 Expense Tracker** → **📦 Archive Previous Years**.
