# Billing System - Google Sheets Setup Guide

## Step 1: Create the Google Sheet

1. Go to https://sheets.google.com and create a new spreadsheet
2. Rename it to "Billing System"
3. Create 3 sheet tabs (click the + at the bottom) named EXACTLY:
   - Clients
   - Invoices
   - Payments

4. In the **Clients** tab, put these headers in Row 1 (one per column A-F):
   id | name | address | phone | email | createdAt

5. In the **Invoices** tab, put these headers in Row 1 (A-J):
   id | invoiceNumber | clientId | date | serveeName | serveeAddress | amount | reference | notes | createdAt

6. In the **Payments** tab, put these headers in Row 1 (A-H):
   id | clientId | invoiceId | date | amount | method | reference | createdAt

## Step 2: Deploy the Apps Script

1. In your Google Sheet, click **Extensions > Apps Script**
2. Delete any existing code in the editor
3. Copy/paste the ENTIRE contents of `google-apps-script.js` from this folder
4. Click the floppy disk icon (Save), or Ctrl+S
5. Click **Deploy > New deployment**
6. Click the gear next to "Select type" and pick **Web app**
7. Set these options:
   - Description: Billing System API
   - Execute as: Me
   - Who has access: Anyone
8. Click **Deploy**
9. Click **Authorize access** and follow the prompts
   - If you see "Google hasn't verified this app", click "Advanced" then "Go to... (unsafe)"
   - This is normal for personal scripts
10. Copy the **Web app URL** it shows you (starts with https://script.google.com/macros/s/...)

## Step 3: Open the Billing System

1. Open `index.html` in any web browser (double-click it)
2. Paste your Web app URL when prompted
3. Click Connect
4. Done! You're connected to your Google Sheet database.

## Sharing with your team

- Give everyone a copy of `index.html` (or put it on a shared drive)
- Everyone enters the SAME Apps Script URL
- All users read/write to the same Google Sheet
- You can also open the Google Sheet directly anytime to see raw data

## Backup

- The Google Sheet IS your backup - it's just a spreadsheet
- You can do File > Make a copy anytime
- Google also keeps version history (File > Version history)

## Troubleshooting

- If you get connection errors, check that the URL is correct
- If you redeploy the script, you'll get a NEW URL - update it in Settings
- Make sure sheet tab names are exactly: Clients, Invoices, Payments (case-sensitive)
