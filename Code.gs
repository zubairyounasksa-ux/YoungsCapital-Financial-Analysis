// ============================================================
// FinSight — Google Apps Script Backend
// File: Code.gs
// Deploy as: Web App → Execute as Me → Anyone can access
// ============================================================

const SHEET_NAME_DATA    = "Financial Data";
const SHEET_NAME_DASH    = "Dashboard";
const SHEET_NAME_CONFIG  = "Config";

// ─────────────────────────────────────────────────────────────
// HTTP entry points
// ─────────────────────────────────────────────────────────────

function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("FinSight — Financial Analysis")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;

    if (action === "saveEntry")   return respond(saveEntry(payload.data));
    if (action === "getEntries")  return respond(getEntries());
    if (action === "deleteEntry") return respond(deleteEntry(payload.rowId));

    return respond({ ok: false, error: "Unknown action" });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────
// Sheet bootstrap
// ─────────────────────────────────────────────────────────────

function getOrCreateSheet(name) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── DATA sheet ────────────────────────────────────────────
  const ds = getOrCreateSheet(SHEET_NAME_DATA);
  if (ds.getLastRow() === 0) {
    const headers = [
      "Row ID", "Company", "Period", "Analyzed At",
      // Income Statement
      "Net Sales (Cur) PKR bn", "Net Sales (Prev) PKR bn", "Revenue Growth %",
      "Cost of Sales (Cur)", "Cost of Sales (Prev)",
      "Gross Profit (Cur)", "Gross Profit (Prev)",
      "Admin & Dist Exp (Cur)", "Admin & Dist Exp (Prev)",
      "Operating Profit (Cur)", "Operating Profit (Prev)",
      "Other Income (Cur)", "Other Income (Prev)",
      "Finance Cost (Cur)", "Finance Cost (Prev)",
      "PBT (Cur)", "PBT (Prev)",
      "PAT (Cur)", "PAT (Prev)",
      "EPS (Cur) PKR", "EPS (Prev) PKR",
      // Balance Sheet
      "Total Assets PKR bn", "Equity PKR bn", "Cash & Bank PKR bn",
      "LT Investments PKR bn", "LT Debt PKR bn", "ST Borrowings PKR bn",
      // Cash Flow
      "Operating CF PKR bn", "Investing CF PKR bn",
      "Financing CF PKR bn", "Ending Cash PKR bn",
      // Dividends
      "Interim Dividend", "Final Dividend", "Bonus / Rights",
      // Narrative
      "Quick Rationale", "Outlook"
    ];
    ds.appendRow(headers);

    // Style header row
    const headerRange = ds.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#1D9E75")
               .setFontColor("#FFFFFF")
               .setFontWeight("bold")
               .setFontSize(10)
               .setWrap(false);
    ds.setFrozenRows(1);
    ds.setColumnWidth(1, 80);
    ds.setColumnWidth(2, 160);
    ds.setColumnWidth(3, 90);
    ds.setColumnWidth(4, 100);
    // Widen narrative columns
    ds.setColumnWidth(39, 350);
    ds.setColumnWidth(40, 350);
  }

  // ── DASHBOARD sheet ───────────────────────────────────────
  const dash = getOrCreateSheet(SHEET_NAME_DASH);
  buildDashboard(dash);

  // ── CONFIG sheet ──────────────────────────────────────────
  const cfg = getOrCreateSheet(SHEET_NAME_CONFIG);
  if (cfg.getLastRow() === 0) {
    cfg.appendRow(["Setting", "Value"]);
    cfg.appendRow(["Currency", "PKR"]);
    cfg.appendRow(["Units", "bn"]);
    cfg.appendRow(["Created", new Date().toLocaleString()]);
    cfg.getRange(1,1,1,2).setBackground("#1D9E75").setFontColor("#FFFFFF").setFontWeight("bold");
  }

  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(dash);
  return { ok: true };
}

// ─────────────────────────────────────────────────────────────
// Dashboard builder
// ─────────────────────────────────────────────────────────────

function buildDashboard(dash) {
  dash.clearContents();
  dash.clearFormats();

  // Title block
  dash.getRange("A1").setValue("FinSight — Financial Analysis Dashboard");
  dash.getRange("A1").setFontSize(18).setFontWeight("bold").setFontColor("#0F6E56");
  dash.getRange("A2").setValue("Auto-updated from Financial Data sheet");
  dash.getRange("A2").setFontSize(10).setFontColor("#888780");
  dash.getRange("A3").setValue("Last refreshed: " + new Date().toLocaleString());
  dash.getRange("A3").setFontSize(10).setFontColor("#888780");

  // KPI labels row (row 5)
  const kpiLabels = ["Total Reports", "Latest Company", "Avg Revenue Growth", "Avg PAT Growth"];
  kpiLabels.forEach((lbl, i) => {
    const col = i * 2 + 1;
    dash.getRange(5, col).setValue(lbl);
    dash.getRange(5, col).setBackground("#E1F5EE").setFontColor("#0F6E56")
        .setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center");
    // Merge 2 cols per KPI for visual width
    dash.getRange(5, col, 1, 2).merge().setBackground("#E1F5EE");
  });

  // KPI formulas row (row 6)
  const dataSheet = `'${SHEET_NAME_DATA}'`;
  dash.getRange("A6").setFormula(`=COUNTA(${dataSheet}!B:B)-1`);
  dash.getRange("C6").setFormula(`=IFERROR(INDEX(${dataSheet}!B:B,COUNTA(${dataSheet}!B:B)),"-")`);
  dash.getRange("E6").setFormula(`=IFERROR(TEXT(AVERAGE(${dataSheet}!G2:G),"0.0")&"%","-")`);
  dash.getRange("G6").setFormula(`=IFERROR(TEXT(AVERAGEIF(${dataSheet}!W2:W,"<>"&"",(${dataSheet}!W2:W-${dataSheet}!X2:X)/ABS(${dataSheet}!X2:X))*100,"0.0")&"%","-")`);

  ["A6","C6","E6","G6"].forEach(cell => {
    dash.getRange(cell, 1, 1, 2);
    dash.getRange(cell).setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  });

  // Section header — Summary table
  dash.getRange("A8").setValue("All Reports Summary");
  dash.getRange("A8").setFontSize(12).setFontWeight("bold").setFontColor("#0F6E56");
  dash.getRange("A8:H8").setBackground("#EAF3DE");

  // Summary table headers row 9
  const sumHeaders = ["Company", "Period", "Net Sales (Cur)", "PAT (Cur)", "EPS (Cur)", "Rev Growth %", "Analyzed At", "Status"];
  sumHeaders.forEach((h, i) => {
    dash.getRange(9, i+1).setValue(h);
    dash.getRange(9, i+1).setBackground("#1D9E75").setFontColor("#FFFFFF")
        .setFontWeight("bold").setFontSize(10);
  });

  // Dynamic rows from data sheet (rows 10-60, up to 50 reports)
  for (let r = 10; r <= 60; r++) {
    const dataRow = r - 8; // maps to data row (row 2 of data = row 10 of dash)
    dash.getRange(r, 1).setFormula(`=IFERROR(${dataSheet}!B${dataRow},"")`);
    dash.getRange(r, 2).setFormula(`=IFERROR(${dataSheet}!C${dataRow},"")`);
    dash.getRange(r, 3).setFormula(`=IFERROR(IF(${dataSheet}!E${dataRow}<>"",TEXT(${dataSheet}!E${dataRow},"0.00")&" bn",""),"")`);
    dash.getRange(r, 4).setFormula(`=IFERROR(IF(${dataSheet}!W${dataRow}<>"",TEXT(${dataSheet}!W${dataRow},"0.00")&" bn",""),"")`);
    dash.getRange(r, 5).setFormula(`=IFERROR(IF(${dataSheet}!Y${dataRow}<>"","PKR "&TEXT(${dataSheet}!Y${dataRow},"0.00"),""),"")`);
    dash.getRange(r, 6).setFormula(`=IFERROR(IF(${dataSheet}!G${dataRow}<>"",TEXT(${dataSheet}!G${dataRow},"0.0")&"%",""),"")`);
    dash.getRange(r, 7).setFormula(`=IFERROR(${dataSheet}!D${dataRow},"")`);
    dash.getRange(r, 8).setFormula(`=IFERROR(IF(${dataSheet}!B${dataRow}<>"","Completed",""),"")`);

    // Alternating row colors
    if (r % 2 === 0) dash.getRange(r, 1, 1, 8).setBackground("#F1EFE8");
  }

  // Column widths on dashboard
  [160, 90, 120, 120, 100, 100, 100, 90].forEach((w, i) => dash.setColumnWidth(i+1, w));
  dash.setFrozenRows(9);
}

// ─────────────────────────────────────────────────────────────
// CRUD operations
// ─────────────────────────────────────────────────────────────

function saveEntry(data) {
  try {
    const ds  = getOrCreateSheet(SHEET_NAME_DATA);
    if (ds.getLastRow() === 0) setupSheets();

    const is  = data.incomeStatement  || {};
    const bs  = data.balanceSheet     || {};
    const cf  = data.cashFlow         || {};
    const div = data.dividends        || {};

    const ns = is.netSales || {};
    const growth = (ns.current && ns.previous)
      ? (((ns.current - ns.previous) / Math.abs(ns.previous)) * 100).toFixed(2)
      : "";

    const rowId = Date.now().toString();

    const row = [
      rowId,
      data.company,
      data.date,
      new Date().toLocaleString(),
      // Income Statement
      ns.current ?? "",        ns.previous ?? "",     growth,
      is.costOfSales?.current  ?? "", is.costOfSales?.previous  ?? "",
      is.grossProfit?.current  ?? "", is.grossProfit?.previous  ?? "",
      is.adminDistExp?.current ?? "", is.adminDistExp?.previous ?? "",
      is.operatingProfit?.current ?? "", is.operatingProfit?.previous ?? "",
      is.otherIncome?.current  ?? "", is.otherIncome?.previous  ?? "",
      is.financeCost?.current  ?? "", is.financeCost?.previous  ?? "",
      is.profitBeforeTax?.current ?? "", is.profitBeforeTax?.previous ?? "",
      is.profitAfterTax?.current  ?? "", is.profitAfterTax?.previous  ?? "",
      is.eps?.current ?? "",   is.eps?.previous ?? "",
      // Balance Sheet
      bs.totalAssets ?? "",    bs.equity ?? "",        bs.cashAndBank ?? "",
      bs.longTermInvestments ?? "", bs.longTermDebt ?? "", bs.shortTermBorrowings ?? "",
      // Cash Flow
      cf.operatingCashFlow ?? "", cf.investingCashFlow ?? "",
      cf.financingCashFlow ?? "", cf.endingCash ?? "",
      // Dividends
      div.interim ?? "",       div.final ?? "",        div.bonusRights ?? "",
      // Narrative
      data.quickRationale ?? "", data.outlook ?? ""
    ];

    ds.appendRow(row);

    // Format the new data row
    const lastRow = ds.getLastRow();
    // Number format for financial columns (E to AH = cols 5 to 35 minus text cols)
    const numCols = ds.getRange(lastRow, 5, 1, 31);
    numCols.setNumberFormat("0.00");
    // Growth % column (col 7)
    ds.getRange(lastRow, 7).setNumberFormat("0.00");
    // Alternate row shading
    if (lastRow % 2 === 0) ds.getRange(lastRow, 1, 1, row.length).setBackground("#F1EFE8");

    // Refresh dashboard timestamp
    refreshDashboardTimestamp();

    return { ok: true, rowId, rowNum: lastRow };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function getEntries() {
  try {
    const ds = getOrCreateSheet(SHEET_NAME_DATA);
    if (ds.getLastRow() <= 1) return { ok: true, entries: [] };

    const rows = ds.getRange(2, 1, ds.getLastRow() - 1, ds.getLastColumn()).getValues();
    const entries = rows
      .filter(r => r[1] !== "")
      .map(r => ({
        rowId:   r[0]?.toString(),
        company: r[1],
        date:    r[2],
        analyzedAt: r[3],
        netSalesCur:  r[4],
        netSalesPrev: r[5],
        revenueGrowth: r[6],
        patCur:   r[21],
        patPrev:  r[22],
        epsCur:   r[23],
        epsPrev:  r[24],
        quickRationale: r[38],
        outlook:  r[39]
      }))
      .reverse();

    return { ok: true, entries };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function deleteEntry(rowId) {
  try {
    const ds   = getOrCreateSheet(SHEET_NAME_DATA);
    const data = ds.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]?.toString() === rowId?.toString()) {
        ds.deleteRow(i + 1);
        refreshDashboardTimestamp();
        return { ok: true };
      }
    }
    return { ok: false, error: "Row not found" };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function refreshDashboardTimestamp() {
  try {
    const dash = getOrCreateSheet(SHEET_NAME_DASH);
    dash.getRange("A3").setValue("Last refreshed: " + new Date().toLocaleString());
  } catch (_) {}
}

// Run once manually after deployment to initialize sheets
function initialize() {
  setupSheets();
  SpreadsheetApp.getUi().alert("FinSight sheets initialized successfully!");
}
