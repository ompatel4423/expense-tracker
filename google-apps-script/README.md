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
  - `transactions`: Array of recorded expenses with `date`, `time`, `amount`, and `merchant`
- `GET ?action=updateDailyLimit&val=<AMOUNT>` - Updates daily limit cell in the Google Sheet.
- `GET ?action=updateMonthlyLimit&val=<AMOUNT>` - Updates monthly limit cell in the Google Sheet.

## Gmail Automation
The script filters incoming emails from Axis Bank:
`from:alerts@axis.bank.in "Amount Debited" -label:ExpensesProcessed`
When run via a time-driven trigger (e.g., every 5-10 minutes), it automatically extracts the transaction amount and merchant name, appends it to the Google Sheet, and labels the email as processed.
