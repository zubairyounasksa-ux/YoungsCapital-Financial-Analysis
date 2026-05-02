# FinSight — Google Sheets Setup Guide

A complete step-by-step guide to deploy your financial analysis system.

---

## What you'll get

| Component | Purpose |
|-----------|---------|
| `Code.gs` | Apps Script backend — receives data, writes to Sheets |
| `index.html` | Web frontend — served by Apps Script |
| **Financial Data** sheet | Every analyzed report, all 40 columns |
| **Dashboard** sheet | Auto-updating summary with formulas |
| **Config** sheet | Settings reference |

---

## Step 1 — Create a new Google Spreadsheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a **blank spreadsheet**
2. Name it: `FinSight Financial Analysis`
3. Keep it open

---

## Step 2 — Open Apps Script

1. In your spreadsheet, click **Extensions → Apps Script**
2. A new browser tab opens with the Apps Script editor

---

## Step 3 — Add Code.gs

1. Delete any existing code in the editor (the default `function myFunction() {}`)
2. Copy the **entire contents** of `Code.gs` and paste it in
3. Click **Save** (Ctrl+S / Cmd+S)

---

## Step 4 — Add index.html

1. In the Apps Script editor, click the **+** button next to "Files" in the left sidebar
2. Choose **HTML**
3. Name the file exactly: `index` (without `.html` — Apps Script adds it)
4. Delete any existing content
5. Copy the **entire contents** of `index.html` and paste it in
6. Click **Save**

---

## Step 5 — Initialize the spreadsheet

1. In the Apps Script editor, at the top, make sure the function dropdown shows **`initialize`**
2. Click **Run**
3. First run: Google will ask for permissions — click **Review permissions → Allow**
4. You should see an alert: *"FinSight sheets initialized successfully!"*
5. Switch back to your spreadsheet — you'll see 3 new sheets: **Dashboard**, **Financial Data**, **Config**

---

## Step 6 — Deploy as a Web App

1. In Apps Script editor, click **Deploy → New deployment**
2. Click the gear icon ⚙ next to "Type" → select **Web app**
3. Fill in:
   - **Description:** `FinSight v1`
   - **Execute as:** `Me`
   - **Who has access:** `Anyone` *(so the frontend can call it)*
4. Click **Deploy**
5. Copy the **Web app URL** — it looks like:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```
6. Keep this URL — you'll need it in the next step

---

## Step 7 — Get your Anthropic API Key

1. Go to [console.anthropic.com](https://console.anthropic.com)
2. Click **API Keys → Create Key**
3. Copy the key (starts with `sk-ant-…`)

---

## Step 8 — Open the Web App and configure it

1. Open the **Web App URL** from Step 6 in your browser
2. The FinSight dashboard loads
3. Click **Settings** (top right)
4. Paste:
   - **Web App URL** — the URL from Step 6
   - **Anthropic API Key** — the key from Step 7
5. Click **Test connection** — you should see "Connection successful!"
6. Click **Save settings**

---

## Step 9 — Analyze your first report

1. Click **Analyze Screenshots**
2. Enter the **Company name** and **Report period**
3. Upload one or more screenshots (income statement, balance sheet, cash flow)
4. Click **Analyze**
5. Wait ~15–30 seconds for AI extraction
6. Result appears in the dashboard and is saved to your Google Sheet

---

## How data flows

```
Your screenshots
      │
      ▼
Anthropic Claude Vision API (in your browser)
      │  extracts structured JSON
      ▼
Apps Script Web App (Code.gs)
      │  writes 40 columns per row
      ▼
Google Sheets — Financial Data tab
      │
      ▼
Dashboard tab (live formula-linked summary)
```

---

## Re-deploying after code changes

If you edit `Code.gs` or `index.html`:

1. In Apps Script, click **Deploy → Manage deployments**
2. Click the pencil ✏ icon on your deployment
3. Change **Version** to `New version`
4. Click **Deploy**

> ⚠ The Web App URL stays the same — no need to update Settings.

---

## GitHub hosting (optional)

If you want `index.html` to live on GitHub Pages instead of being served by Apps Script:

1. Remove the `doGet()` function from `Code.gs` (it's only needed if Apps Script serves the HTML)
2. Host `index.html` on GitHub Pages
3. In `index.html`, the **Web App URL** in Settings still points to your Apps Script deployment
4. Add your GitHub Pages URL to the Apps Script project's CORS settings if needed

> The simplest approach: let Apps Script serve the HTML (the default setup above).

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "Connection failed" | Make sure "Who has access" is set to "Anyone" in Deploy settings |
| "API key error" | Check your Anthropic key in Settings — must start with `sk-ant-` |
| Blank dashboard | Run `initialize` function again from Apps Script editor |
| Data not showing | Click **Refresh** button on the dashboard |
| "Script not found" | Re-deploy and copy the new URL |
| Permission errors | In Apps Script, click Run on any function and re-authorize |

---

## Sheet structure — Financial Data columns

| Col | Field | Col | Field |
|-----|-------|-----|-------|
| A | Row ID | B | Company |
| C | Period | D | Analyzed At |
| E | Net Sales (Cur) | F | Net Sales (Prev) |
| G | Revenue Growth % | H | Cost of Sales (Cur) |
| I | Cost of Sales (Prev) | J | Gross Profit (Cur) |
| K | Gross Profit (Prev) | L | Admin & Dist Exp (Cur) |
| M | Admin & Dist Exp (Prev) | N | Operating Profit (Cur) |
| O | Operating Profit (Prev) | P | Other Income (Cur) |
| Q | Other Income (Prev) | R | Finance Cost (Cur) |
| S | Finance Cost (Prev) | T | PBT (Cur) |
| U | PBT (Prev) | V | PAT (Cur) |
| W | PAT (Prev) | X | EPS (Cur) |
| Y | EPS (Prev) | Z | Total Assets |
| AA | Equity | AB | Cash & Bank |
| AC | LT Investments | AD | LT Debt |
| AE | ST Borrowings | AF | Operating CF |
| AG | Investing CF | AH | Financing CF |
| AI | Ending Cash | AJ | Interim Dividend |
| AK | Final Dividend | AL | Bonus/Rights |
| AM | Quick Rationale | AN | Outlook |
