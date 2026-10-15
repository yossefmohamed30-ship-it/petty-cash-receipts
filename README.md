# 🧾 Petty Cash Receipt Uploader

Google Apps Script sidebar / web app for Google Sheets:

1. Upload a receipt photo (or PDF / transfer screenshot).
2. It's saved to Google Drive (`Petty Cash Receipts` folder).
3. Gemini reads the image and extracts the amount, date and type (out / in).
4. You type what it was for → **Done** → a new row is added:
   `Date | Amount (negative = spent) | Description | link to receipt`

## Setup
1. In the spreadsheet: **Extensions → Apps Script**.
2. Add `ReceiptUploader.gs` (script) and `ReceiptForm.html` (HTML file named `ReceiptForm`).
3. Set `SPREADSHEET_ID` in `CONFIG` to your sheet's ID.
4. Save, reload the sheet → menu **🧾 المصروفات**.
5. Get a free Gemini API key from https://aistudio.google.com/apikey and set it via **🔑 إعداد مفتاح Gemini** (stored in Script Properties, never in code).
6. Optional: **Deploy → Web app** to open it in its own tab / on mobile.

Target sheet tabs need headers in row 1: `Date | Amount | Description | link`.
