// ============================================================
//  D2C MARKETING INTELLIGENCE DASHBOARD
//  Main Controller — Code.gs
//  Google Apps Script
// ============================================================

// ─── CONSTANTS ────────────────────────────────────────────
var META_SHEET   = 'META_DATA';
var GOOGLE_SHEET = 'GOOGLE_DATA';
var CONFIG_SHEET = 'CONFIG';
var PRODUCT_SHEET = 'PRODUCT_MASTER';
var SALES_SHEET   = 'SALES_DATA';
var DAILY_SHEET   = 'DAILY_STATS';

// ─── ON OPEN: Custom Menu ─────────────────────────────────
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 D2C Intelligence')
    .addItem('🌐 Launch Live Dashboard', 'openWebAppLink')
    .addItem('📊 Open HTML Dialog (Legacy)', 'openDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('📥 Data & Sync')
      .addItem('⚙️ Setup All Tabs (run first time)', 'createAllTabs')
      .addItem('🔑 Setup API Keys in CONFIG', 'setupConfigApiKeys')
      .addItem('🔄 Run Daily Sync Now (manual)', 'dailyAutoSync')
      .addItem('📅 Sync for a Specific Date', 'manualSyncForDate')
      .addItem('⏰ Setup Daily Auto-Sync Trigger (run once)', 'setupDailySyncTrigger'))
    .addSeparator()
    .addItem('🌱 Insert Sample Data', 'insertSampleData')
    .addItem('🤖 Auto-Tag Blank Rows (new only)', 'autoTagBlankRows')
    .addItem('🔁 Force Re-Tag ALL Rows', 'forceReTagAllRows')
    .addItem('⚙️ Setup Sheets (legacy)', 'setupSheets')
    .addSeparator()
    .addItem('📖 Create / Refresh README Sheet', 'createReadmeSheet')
    .addItem('🔗 Get Live Dashboard URL', 'showLiveDashboardUrl')
    .addSeparator()
    .addItem('🛠️ Setup Native Dashboard', 'setupNativeDashboard')
    .addItem('🔄 Update Dashboard (Native)', 'refreshNativeIntelligence')
    .addItem('🔧 Run Diagnostics', 'runDiagnostics')
    .addToUi();
}


// ─── SHOW LIVE DASHBOARD URL ──────────────────────────────
function showLiveDashboardUrl() {
  var url = ScriptApp.getService().getUrl();
  var ui = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert('⚠️ Web App Not Deployed\n\nTo get a live URL, go to:\nApps Script → Deploy → New Deployment → Web App\n\nSet "Execute as: Me" and "Who has access: Anyone".\nThen copy the URL and bookmark it.');
  } else {
    ui.alert('✅ Your Live Dashboard URL:\n\n' + url + '\n\nBookmark this link. It always shows the latest data automatically.');
  }
}

// ─── DAILY AUTO-TAG TRIGGER SETUP ─────────────────────────
function setupDailyAutoTagTrigger() {
  var ui = SpreadsheetApp.getUi();
  // Remove old triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyAutoTag' || t.getHandlerFunction() === 'forceReTagAllRows') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Create fresh daily trigger at 6 AM IST (UTC+5:30 → ~00:30 UTC)
  ScriptApp.newTrigger('dailyAutoTag')
    .timeBased()
    .everyDays(1)
    .atHour(1)   // 1 AM UTC ≈ 6:30 AM IST
    .create();
  ui.alert('✅ Daily Auto-Tag Trigger Set!\n\nEvery day at ~6 AM IST the system will automatically:\n1. Detect any new rows added to META_DATA or GOOGLE_DATA\n2. Auto-tag them with Product, Narrative, Language, Ad Type, Influencer\n3. Data will be ready for the live dashboard\n\nYou only need to run this setup once.');
}

// ─── DAILY AUTO-TAG (called by trigger) ───────────────────
function dailyAutoTag() {
  logDebug('Daily auto-tag triggered at ' + new Date().toISOString());
  try {
    autoTagBlankRows(); // Tags only untagged rows first (fast)
  } catch(e) {
    logDebug('Auto-tag error: ' + e.message);
  }
}

// ─── ON EDIT TRIGGER: auto-tag new rows as they're pasted ─
function onEditAutoTag(e) {
  try {
    var sheet = e.range.getSheet();
    var name = sheet.getName();
    if (name !== META_SHEET && name !== GOOGLE_SHEET) return;
    // Only trigger if more than 10 rows changed (bulk paste)
    if (e.range.getNumRows() > 5) {
      autoTagBlankRows();
    }
  } catch(ex) {}
}

// ─── SERVE DATA FOR LIVE DASHBOARD ───────────────────────
function getSheetDataForDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = { meta: [], google: [], lastUpdated: new Date().toISOString(), sheetId: ss.getId() };
  
  var metaSheet = ss.getSheetByName(META_SHEET);
  if (metaSheet && metaSheet.getLastRow() > 1) {
    var data = metaSheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().trim(); });
    for (var i = 1; i < data.length; i++) {
      var row = {};
      headers.forEach(function(h, j) { row[h] = data[i][j]; });
      result.meta.push(row);
    }
  }
  
  var gSheet = ss.getSheetByName(GOOGLE_SHEET);
  if (gSheet && gSheet.getLastRow() > 1) {
    var gData = gSheet.getDataRange().getValues();
    var gHeaders = gData[0].map(function(h) { return h.toString().trim(); });
    for (var i = 1; i < gData.length; i++) {
      var row = {};
      gHeaders.forEach(function(h, j) { row[h] = gData[i][j]; });
      result.google.push(row);
    }
  }
  
  return JSON.stringify(result);
}

// ─── CREATE README SHEET ──────────────────────────────────
function createReadmeSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var README_SHEET = 'README';
  
  // Delete existing to refresh
  var existing = ss.getSheetByName(README_SHEET);
  if (existing) ss.deleteSheet(existing);
  
  var sheet = ss.insertSheet(README_SHEET, 0); // Insert as first tab
  
  // Styling setup
  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 700);
  sheet.setColumnWidth(3, 250);
  
  var content = [
    ['', '📋 README — Be.Bodywise D2C Marketing Intelligence System', ''],
    ['', 'This sheet auto-tags and analyses all your Meta & Google ad performance data. Follow the steps below every day.', ''],
    ['', '', ''],
    ['', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', ''],
    ['', 'SECTION 1: HOW TO PUSH META ADS DATA (Daily)', ''],
    ['', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', ''],
    ['', '', ''],
    ['', 'STEP 1 — Log into Meta Ads Manager', 'https://adsmanager.facebook.com'],
    ['', 'STEP 2 — Go to "Ads" level (not Campaign or Ad Set)', ''],
    ['', 'STEP 3 — Set date range to "Yesterday" (or custom date range you want to add)', ''],
    ['', 'STEP 4 — Click "Columns" → Select "Customise Columns"', ''],
    ['', 'STEP 5 — Add these exact columns IN THIS ORDER:', ''],
    ['', '   Date, Campaign name, Ad set name, Ad name, Amount spent (INR), Impressions, Clicks (all), Link clicks, Landing page views, Purchases, Purchase conversion value, Add to Cart, CTR (link click-through rate), CPC (cost per link click), ROAS (Return on Ad Spend), CPM (cost per 1,000 impressions), Frequency, 3-second video plays, Video plays at 25%, Video plays at 50%, Video plays at 75%, Video plays at 100% (ThruPlay)', ''],
    ['', 'STEP 6 — Click "Export" → "Export Table Data (CSV)"', ''],
    ['', 'STEP 7 — Open the CSV in Excel/Google Sheets. Select ALL data (Ctrl+A)', ''],
    ['', 'STEP 8 — Come to this Google Sheet → Click the META_DATA tab', ''],
    ['', 'STEP 9 — Go to the LAST empty row (below existing data)', ''],
    ['', 'STEP 10 — Paste (Ctrl+V). Do NOT paste headers again — only paste data rows', ''],
    ['', 'STEP 11 — Go to 🚀 D2C Intelligence menu → 🤖 Auto-Tag Blank Rows', ''],
    ['', '         This will tag all the new rows with Product, Narrative, Language, Influencer etc.', ''],
    ['', 'STEP 12 — Open the Live Dashboard URL to see updated data', ''],
    ['', '', ''],
    ['', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', ''],
    ['', 'SECTION 2: HOW TO PUSH GOOGLE ADS DATA (Daily)', ''],
    ['', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', ''],
    ['', '', ''],
    ['', 'STEP 1 — Log into Google Ads', 'https://ads.google.com'],
    ['', 'STEP 2 — Go to "Ads & Assets" → "Ads" view', ''],
    ['', 'STEP 3 — Set date range to "Yesterday" (or custom date range)', ''],
    ['', 'STEP 4 — Click "Columns" → "Modify Columns" → Add these:', ''],
    ['', '   Date, Campaign, Ad group, Ad name, Cost (INR), Impressions, Clicks, Conversions, Conversion value, View-through conv., Avg. CPC, CTR, Conv. rate, ROAS', ''],
    ['', 'STEP 5 — Click the download icon (↓) → Download as CSV', ''],
    ['', 'STEP 6 — Open CSV → Select all data rows (NOT the header)', ''],
    ['', 'STEP 7 — Come to this Google Sheet → Click GOOGLE_DATA tab', ''],
    ['', 'STEP 8 — Paste at the BOTTOM of existing data (do NOT overwrite or add headers again)', ''],
    ['', 'STEP 9 — Go to 🚀 D2C Intelligence menu → 🤖 Auto-Tag Blank Rows', ''],
    ['', 'STEP 10 — Open the Live Dashboard URL to see updated Google data', ''],
    ['', '', ''],
    ['', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', ''],
    ['', 'SECTION 3: IMPORTANT RULES — READ CAREFULLY', ''],
    ['', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', ''],
    ['', '', ''],
    ['', '⚠️  NEVER delete existing rows in META_DATA or GOOGLE_DATA. Only ADD new rows at the bottom.', ''],
    ['', '⚠️  NEVER paste column headers again when adding new data. Only paste data rows.', ''],
    ['', '⚠️  NEVER modify the auto-tagged columns (Product, Ad Type, Type of Ad, Influencer Name, Narrative, Language).', ''],
    ['', '⚠️  If you re-run a date that already exists in the sheet, first filter and delete those date rows, then paste fresh.', ''],
    ['', '✅  You can add data for multiple days at once — just paste all rows together at the bottom.', ''],
    ['', '✅  If a daily trigger is set up (via menu), tagging happens AUTOMATICALLY every morning at 6 AM IST.', ''],
    ['', '✅  The Live Dashboard always reads the latest tagged data from this sheet.', ''],
    ['', '', ''],
    ['', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', ''],
    ['', 'SECTION 4: SHEET STRUCTURE', ''],
    ['', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', ''],
    ['', '', ''],
    ['', 'META_DATA    — All Meta (Facebook/Instagram) ad-level daily performance data', ''],
    ['', 'GOOGLE_DATA  — All Google Ads campaign/ad group level daily data', ''],
    ['', 'SALES_DATA   — Optional: daily revenue/orders from your store backend', ''],
    ['', 'PRODUCT_MASTER — Product names, categories, target ROAS, cost per unit', ''],
    ['', 'CONFIG       — Dashboard settings: brand name, currency, ROAS targets, platforms', ''],
    ['', 'README       — This sheet (instructions)', ''],
    ['', '', ''],
    ['', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', ''],
    ['', 'SECTION 5: AUTO-TAGGING EXPLAINED', ''],
    ['', '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', ''],
    ['', '', ''],
    ['', 'The system reads the "Ad name" column and extracts:', ''],
    ['', '  • Product      — which product the ad is for', ''],
    ['', '  • Ad Type      — Internal or Influencer', ''],
    ['', '  • Type of Ad   — Reel, Static Creative, UGC', ''],
    ['', '  • Influencer   — name of the influencer if present', ''],
    ['', '  • Narrative    — the creative angle (Problem/Solution, Before/After, etc.)', ''],
    ['', '  • Language     — language of the ad creative', ''],
    ['', '', ''],
    ['', 'For auto-tagging to work BEST, ad names must follow this structure:', ''],
    ['', '  [source]_[format]_[product]_[creator]_[language]_[narrative]_[targeting]_[serial]_[date]', ''],
    ['', '  Example: influencer_music_reel_hair_gummies_vaishnavi_bhavsar_english_before_after_pan_india_0053_210126', ''],
    ['', '', ''],
    ['', 'Refer to the Master Ad Naming System document shared by your team for the full guide.', ''],
  ];
  
  sheet.getRange(1, 1, content.length, 3).setValues(content);
  
  // Format the title row
  var titleRange = sheet.getRange(1, 2, 1, 2);
  titleRange.setFontSize(16).setFontWeight('bold').setFontColor('#1a56db');
  
  // Format section headers
  var sectionRows = [5, 19, 32, 45, 56];
  sectionRows.forEach(function(r) {
    sheet.getRange(r, 2).setFontSize(13).setFontWeight('bold').setFontColor('#1e3a5f');
    sheet.getRange(r, 2).setBackground('#e8f0fe');
  });
  
  // Format step rows
  sheet.getRange(1, 1, content.length, 3).setFontFamily('Arial').setFontSize(11).setVerticalAlignment('top').setWrap(true);
  sheet.setFrozenRows(1);
  
  SpreadsheetApp.getUi().alert('✅ README sheet created!\n\nIt has been added as the FIRST tab in your spreadsheet.\nShare it with your media buying team so they know exactly how to push data.');
}


var debugLogs = [];

function logDebug(msg) {
  debugLogs.push([new Date(), msg]);
  console.log(msg);
}

function runDiagnostics() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('DEBUG_LOG');
  if (logSheet) {
    if (logSheet.getLastRow() > 1) {
      logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 2).clearContent();
    }
  } else {
    logSheet = ss.insertSheet('DEBUG_LOG');
    logSheet.appendRow(['Timestamp', 'Message']);
    logSheet.getRange('A1:B1').setBackground('#1f2937').setFontColor('#ffffff').setFontWeight('bold');
  }

  debugLogs = []; // Reset logs
  logDebug("Diagnostics started manually");
  try {
    var start = new Date().getTime();
    
    // Proactively log raw META_DATA headers & rows for troubleshooting
    var metaSheet = ss.getSheetByName(META_SHEET);
    if (metaSheet && metaSheet.getLastRow() > 0) {
      var rawHeaders = metaSheet.getRange(1, 1, 1, metaSheet.getLastColumn()).getValues()[0];
      logDebug("META_DATA Headers count: " + rawHeaders.length + ", values: " + JSON.stringify(rawHeaders));
      var sampleRowsCount = Math.min(metaSheet.getLastRow() - 1, 3);
      if (sampleRowsCount > 0) {
        var rawRows = metaSheet.getRange(2, 1, sampleRowsCount, metaSheet.getLastColumn()).getValues();
        rawRows.forEach(function(rowValues, index) {
          logDebug("META_DATA Raw Row " + (index + 1) + ": " + JSON.stringify(rowValues));
        });
      }
    } else {
      logDebug("META_DATA sheet not found or empty");
    }

    var data = getDashboardData({
      startDate: formatDateJS(getDefaultStartDate()),
      endDate: formatDateJS(new Date()),
      platform: 'all',
      product: 'all',
      adType: 'all'
    });
    var elapsed = new Date().getTime() - start;
    logDebug("Diagnostics finished successfully in " + elapsed + " ms");
    
    // Batch write all accumulated logs
    if (debugLogs.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, debugLogs.length, 2).setValues(debugLogs);
    }
    SpreadsheetApp.flush();
    
    ui.alert("✅ Diagnostics Succeeded!\n\nExecution Time: " + elapsed + " ms\n\nCheck the 'DEBUG_LOG' sheet tab for step-by-step logs.");
  } catch (err) {
    logDebug("Diagnostics failed: " + err.toString() + "\nStack: " + err.stack);
    
    // Batch write all accumulated logs
    if (debugLogs.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, debugLogs.length, 2).setValues(debugLogs);
    }
    SpreadsheetApp.flush();
    
    ui.alert("❌ Diagnostics Failed!\n\nError: " + err.message + "\n\nCheck the 'DEBUG_LOG' sheet tab for details.");
  }
}

// ─── GOOGLE WEB APP GET HANDLER ───────────────────────────
function doGet() {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('📊 D2C Marketing Intelligence Command Center')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function openWebAppLink() {
  var url = ScriptApp.getService().getUrl();
  var ui = SpreadsheetApp.getUi();
  
  // If URL is empty or matches default editor test URL format, give instruction
  if (!url || url.indexOf('/dev') !== -1 || url.indexOf('script.google.com') === -1) {
    var htmlOutput = HtmlService.createHtmlOutput(
      "<div style='font-family:\"Segoe UI\",Roboto,Helvetica,Arial,sans-serif;padding:18px;color:#1e293b;line-height:1.5;'>" +
      "<h3 style='margin-top:0;color:#0f172a;font-size:15px;display:flex;align-items:center;gap:6px;'>🌐 Deploy Your Free Web App Dashboard</h3>" +
      "<p style='font-size:12px;color:#475569;'>To open your beautiful dashboard in a full-screen standalone tab, please deploy it once:</p>" +
      "<ol style='font-size:12px;color:#334155;padding-left:20px;margin-bottom:15px;'>" +
      "<li style='margin-bottom:6px;'>Click <b>Extensions</b> -> <b>Apps Script</b> in your Google Sheet menu bar.</li>" +
      "<li style='margin-bottom:6px;'>Click the blue <b>Deploy</b> button (top right) and choose <b>New deployment</b>.</li>" +
      "<li style='margin-bottom:6px;'>Click the gear icon (Select type) and choose <b>Web app</b>.</li>" +
      "<li style='margin-bottom:6px;'>Set 'Execute as' to <b>Me</b> and 'Who has access' to <b>Anyone</b> or <b>Only myself</b>.</li>" +
      "<li style='margin-bottom:6px;'>Click <b>Deploy</b> (authorize access if prompted), then copy and open the Web App URL!</li>" +
      "</ol>" +
      "<p style='font-size:11px;color:#64748b;font-style:italic;'>Once deployed, you can bookmark the link to open your dashboard instantly from any device!</p>" +
      "</div>"
    ).setWidth(500).setHeight(320);
    ui.showModalDialog(htmlOutput, '🌐 Deploy Web App');
  } else {
    var htmlOutput = HtmlService.createHtmlOutput(
      "<div style='font-family:\"Segoe UI\",Roboto,Helvetica,Arial,sans-serif;padding:20px;color:#1e293b;line-height:1.4;'>" +
      "<div style='text-align:center;margin-bottom:20px;'>" +
      "<h2 style='margin-top:0;font-size:16px;color:#0f172a;'>🚀 Launch Standalone Command Center</h2>" +
      "<p style='margin:8px 0 16px;font-size:12px;color:#64748b;'>Your dashboard will open in a new full-screen browser tab.</p>" +
      "<a href='" + url + "' target='_blank' style='display:inline-block;padding:10px 24px;background:#2563eb;color:#ffffff;font-weight:700;font-size:12.5px;text-decoration:none;border-radius:6px;box-shadow:0 4px 10px rgba(37,99,235,0.25);'>👉 Open Dashboard in New Tab</a>" +
      "</div>" +
      "<div style='border-top:1px solid #e2e8f0;padding-top:12px;'>" +
      "<p style='font-size:11px;color:#ef4444;font-weight:600;margin-top:0;'>⚠️ Getting a 'Page Not Found' or 'Sorry, unable to open...' error?</p>" +
      "<p style='font-size:11px;color:#475569;margin-bottom:8px;'>This means the Web App has not been deployed yet. Follow these quick steps to enable it:</p>" +
      "<ol style='font-size:11px;color:#334155;padding-left:18px;margin-bottom:0;line-height:1.4;'>" +
      "<li style='margin-bottom:4px;'>In the menu, go to <b>Extensions ➔ Apps Script</b>.</li>" +
      "<li style='margin-bottom:4px;'>Click the blue <b>Deploy</b> button (top right) ➔ select <b>New deployment</b>.</li>" +
      "<li style='margin-bottom:4px;'>Click the gear icon (Select type) ➔ choose <b>Web app</b>.</li>" +
      "<li style='margin-bottom:4px;'>Set <i>Execute as</i> to <b>Me</b>, and <i>Who has access</i> to <b>Anyone</b>.</li>" +
      "<li style='margin-bottom:4px;'>Click <b>Deploy</b>, authorize permissions, and try opening the link again!</li>" +
      "</ol>" +
      "</div>" +
      "</div>"
    ).setWidth(500).setHeight(360);
    ui.showModalDialog(htmlOutput, '🌐 Standalone Dashboard');
  }
}

// ─── OPEN DASHBOARD ───────────────────────────────────────
function openDashboard() {
  var html = HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setWidth(1400)
    .setHeight(900)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  SpreadsheetApp.getUi().showModalDialog(html, '📊 D2C Marketing Intelligence Dashboard');
}

// ─── INCLUDE HTML PARTIALS ────────────────────────────────
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── SETUP SHEETS ─────────────────────────────────────────
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // META_DATA — includes full funnel, custom video columns, calculated rates, and custom tags
  ensureSheet(ss, META_SHEET, [
    'Day', 'Account Name', 'Campaign name', 'Ad set name', 'Ad name', 'Currency', 'Amount spent (INR)', 
    'Impressions', 'Link clicks', 'Landing page views', 'Add To Cart', 'Purchases', 'Purchases conversion value', 
    '3-second video plays', 'Video plays at 25%', 'Video plays at 50%', 'Video plays at 75%', 'Video plays at 95%', 'Video plays at 100%', 
    'Website URL', 'Ad set ID', 'Ad ID', 
    'Product', 'Ad Type', 'Type of Ad', 'Influencer Name', 'Narrative', 'Language',
    'ROAS', 'Hook Rate %', 'Hold Rate %'
  ]);

  // GOOGLE_DATA — includes LPV and ATC equivalents
  ensureSheet(ss, GOOGLE_SHEET, [
    'Date','Campaign','Ad Group','Ad Name',
    'Spend','Impressions','Clicks','Link Clicks','Landing Page Views',
    'Conversions','Revenue','Add to Cart','CTR %','CPC','ROAS',
    'Quality Score','Conv Rate %','Hook Rate %','Hold Rate %',
    'Narrative','Type of Ad (Static/Reel)','Language of the Ad','Name of the Influencer','Product Name','Category'
  ]);

  // Apply number formatting to META_DATA
  var metaSheet = ss.getSheetByName(META_SHEET);
  if (metaSheet) {
    metaSheet.getRange('G2:G10000').setNumberFormat('₹#,##0.00'); // Amount spent (INR)
    metaSheet.getRange('H2:L10000').setNumberFormat('#,##0');    // Imp, Link clicks, LP views, ATC, Purchases
    metaSheet.getRange('M2:M10000').setNumberFormat('₹#,##0.00'); // Purchases conversion value
    metaSheet.getRange('N2:S10000').setNumberFormat('#,##0');    // Video plays
    
    // Clear columns AC, AD, AE below row 2 to allow the ArrayFormula to expand without #REF! blockages
    metaSheet.getRange('AC3:AE10000').clearContent();
    
    // Auto-calculate ROAS, Hook Rate %, and Hold Rate % on the sheet itself using ArrayFormulas
    metaSheet.getRange('AC2').setFormula('=ARRAYFORMULA(IF(LEN(A2:A), IFERROR(M2:M / G2:G, 0), ""))');
    metaSheet.getRange('AD2').setFormula('=ARRAYFORMULA(IF(LEN(A2:A), IFERROR((N2:N / H2:H) * 100, 0), ""))');
    metaSheet.getRange('AE2').setFormula('=ARRAYFORMULA(IF(LEN(A2:A), IFERROR((O2:O / N2:N) * 100, 0), ""))');

    metaSheet.getRange('AC2:AC10000').setNumberFormat('0.00"x"');   // ROAS
    metaSheet.getRange('AD2:AE10000').setNumberFormat('0.00"%"');   // Hook & Hold Rate %
  }

  // Apply number formatting to GOOGLE_DATA
  var googleSheet = ss.getSheetByName(GOOGLE_SHEET);
  if (googleSheet) {
    googleSheet.getRange('E2:E10000').setNumberFormat('₹#,##0.00'); // Spend
    googleSheet.getRange('F2:J10000').setNumberFormat('#,##0');    // Imp, Clicks, Link Clicks, LPV, Conversions
    googleSheet.getRange('K2:L10000').setNumberFormat('₹#,##0.00'); // Revenue, ATC
    googleSheet.getRange('M2:M10000').setNumberFormat('0.00"%"');  // CTR %
    googleSheet.getRange('N2:N10000').setNumberFormat('₹#,##0.00'); // CPC
    googleSheet.getRange('O2:O10000').setNumberFormat('0.00"x"');   // ROAS
    googleSheet.getRange('P2:P10000').setNumberFormat('#,##0');    // Quality Score
    googleSheet.getRange('Q2:S10000').setNumberFormat('0.00"%"');  // Conv Rate %, Hook Rate %, Hold Rate %
  }

  // CONFIG
  var cfg = ensureSheet(ss, CONFIG_SHEET, [
    'Metric','Target','Scale Threshold','Stop Threshold','Unit'
  ]);
  var cfgSheet = ss.getSheetByName(CONFIG_SHEET);
  if (cfgSheet.getLastRow() <= 1) {
    cfgSheet.getRange('A2:E11').setValues([
      ['ROAS',         3.0,  4.0,  1.5,  'x'],
      ['CTR',          1.5,  2.5,  0.5,  '%'],
      ['CPC',          50,   30,   100,  '₹'],
      ['Conv Rate',    2.0,  3.5,  0.5,  '%'],
      ['CPM',          200,  150,  500,  '₹'],
      ['Frequency',    2.5,  1.5,  4.0,  'x'],
      ['Min Spend',    200,  500,  0,    '₹'],
      ['Hook Rate',    25,   35,   10,   '%'],
      ['Hold Rate',    20,   30,   8,    '%'],
      ['ATC Rate',     5,    8,    1,    '%']
    ]);
  }

  // PRODUCT_MASTER
  var pm = ensureSheet(ss, PRODUCT_SHEET, ['Product', 'Price', 'Cost of Goods', 'Margin', 'Target CPA']);

  // SALES_DATA
  ensureSheet(ss, SALES_SHEET, [
    'Order ID', 'Customer ID', 'Order Date', 'Revenue', 'Product', 'Platform', 'Cohort Month'
  ]);

  // DAILY_STATS
  ensureSheet(ss, DAILY_SHEET, [
    'Date','Spend','Conversions','CAC','AOV','ROAS'
  ]);

  setupInstructionsTab(ss);
  SpreadsheetApp.getUi().alert('✅ Sheets setup complete! The "INSTRUCTIONS" tab has been created, and your data tabs are ready. Now use "Insert Sample Data" to load test data, then open the Dashboard.');
}

function setupInstructionsTab(ss) {
  var name = 'INSTRUCTIONS';
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(name, 0); // Put it as the first tab
  }
  
  sheet.setHiddenGridlines(false);
  
  var titleBg = '#0f172a'; // slate 900
  var headerBg = '#1e293b'; // slate 800
  var textCol = '#f8fafc';
  
  var rows = [
    ['D2C MARKETING DASHBOARD - USER MANUAL', '', '', ''],
    ['How to copy your marketing data from Google Sheets into the Dashboard', '', '', ''],
    ['', '', '', ''],
    ['⏰ DAILY WORKFLOW STEPS', '', '', ''],
    ['Step', 'Action Required', 'Detailed Workflow Description', 'Tips & Format'],
    ['1', 'Export Ad Data', 'Export Meta Ads & Google Ads daily performance reports from your ads managers.', 'Select Date Range first'],
    ['2', 'Paste to META_DATA', 'Paste your Meta Ads daily dump into the META_DATA sheet tab starting on row 2.', 'Date, Campaign, Spend, Clicks, conversions, etc.'],
    ['3', 'Paste to GOOGLE_DATA', 'Paste your Google Ads daily dump into the GOOGLE_DATA sheet tab starting on row 2.', 'Date, Campaign, Spend, Clicks, conversions, etc.'],
    ['4', 'Paste to SALES_DATA', 'Paste your daily orders log (for LTV cohort calculations) into the SALES_DATA tab starting on row 2.', 'Order ID, Customer ID, Date, Rev, Cohort Month'],
    ['4b', 'Auto-Tag Blank Rows (Optional)', 'Click "🚀 D2C Intelligence" -> "🤖 Auto-Tag Blank Rows" to automatically parse & map Products and Influencer names.', 'Runs naming convention scanner'],
    ['5', 'Copy Sheets Tab Data', 'Go to the META_DATA tab. Press Cmd+A (Mac) or Ctrl+A (Windows) to select all cells, then Cmd+C (Copy).', 'Do NOT exclude header row'],
    ['6', 'Paste into Local Web App', 'Open index.html on your computer, click the Data Import tab, and paste (Cmd+V) into the META_DATA textarea.', 'The parser handles tab columns automatically'],
    ['7', 'Repeat & Sync', 'Repeat the copy-paste for GOOGLE_DATA, SALES_DATA, PRODUCT_MASTER, and CONFIG. Click "Save & Update".', 'Cached instantly in localStorage'],
    ['', '', '', ''],
    ['📋 SPREADSHEET TABS STRUCTURE REFERENCE', '', '', ''],
    ['Sheet Tab Name', 'Required Columns List (Headers must match exactly in row 1)', 'Purpose', 'Notes'],
    ['META_DATA', 'Day, Account Name, Campaign name, Ad set name, Ad name, Currency, Amount spent (INR), Impressions, Link clicks, Landing page views, Add To Cart, Purchases, Purchases conversion value, 3-second video plays, Video plays at 25%, Video plays at 50%, Video plays at 75%, Video plays at 95%, Video plays at 100%, Website URL, Ad set ID, Ad ID, Product, Ad Type, Type of Ad, Influencer Name, Narrative, Language, ROAS, Hook Rate %, Hold Rate %', 'Meta campaign data & creative hook metrics', '28 columns pasted (Product, Ad Type, Type of Ad, Influencer Name, Narrative, Language last). ROAS, Hook Rate %, and Hold Rate % calculated automatically in columns 29-31.'],
    ['GOOGLE_DATA', 'Date, Campaign, Ad Group, Ad Name, Spend, Impressions, Clicks, Link Clicks, Landing Page Views, Conversions, Revenue, Add to Cart, CTR, CPC, ROAS, Quality Score, Conv Rate, Hook Rate, Hold Rate, Narrative, Type of Ad (Static/Reel), Language of the Ad, Name of the Influencer, Product Name, Category', 'Google campaign performance metrics', '25 columns required (Narrative, Type of Ad, Language, Influencer, Product & Category last)'],
    ['SALES_DATA', 'Order ID, Customer ID, Order Date, Revenue, Product, Platform, Cohort Month', 'Transaction ledger for repeat buying cohorts', 'Cohort Month format: YYYY-MM'],
    ['PRODUCT_MASTER', 'Product, Price, Cost of Goods, Margin, Target CPA', 'Stores profit margins per product SKU', 'Margin format: e.g. 0.65 for 65%'],
    ['CONFIG', 'Setting Name, Value, Description', 'Threshold parameters for AI optimizations', 'Target thresholds']
  ];
  
  sheet.getRange(1, 1, rows.length, 4).setValues(rows);
  
  sheet.getRange('A1:D1').merge().setBackground(titleBg).setFontColor(textCol).setFontSize(16).setFontWeight('bold').setVerticalAlignment('middle').setHorizontalAlignment('center');
  sheet.getRange('A2:D2').merge().setBackground(titleBg).setFontColor('#94a3b8').setFontSize(11).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sheet.setRowHeight(1, 40);
  sheet.setRowHeight(2, 25);
  
  sheet.getRange('A4:D4').setBackground('#f8fafc').setFontWeight('bold').setFontSize(12).setFontColor('#0f172a');
  sheet.getRange('A14:D14').setBackground('#f8fafc').setFontWeight('bold').setFontSize(12).setFontColor('#0f172a');
  
  sheet.getRange('A5:D5').setBackground(headerBg).setFontColor(textCol).setFontWeight('bold').setFontSize(10);
  sheet.getRange('A15:D15').setBackground(headerBg).setFontColor(textCol).setFontWeight('bold').setFontSize(10);
  
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 480);
  sheet.setColumnWidth(4, 250);
  
  sheet.getRange('A6:A12').setHorizontalAlignment('center');
  sheet.getRange('A1:D20').setWrap(true);
}

function ensureSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  // Always write/overwrite headers in Row 1 to ensure order and consistency
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  if (sheet.getFrozenRows() === 0) {
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ─── GET DASHBOARD DATA (called from HTML) ────────────────
function getDashboardData(filters) {
  logDebug("getDashboardData execution started");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  filters = filters || {};

  var dateExtent = null;
  if (!filters.startDate || !filters.endDate) {
    dateExtent = detectDateExtent(ss);
  }

  var startDate = filters.startDate ? new Date(filters.startDate) : (dateExtent ? dateExtent.min : getDefaultStartDate());
  var endDate   = filters.endDate   ? new Date(filters.endDate)   : (dateExtent ? dateExtent.max : new Date());
  var platform  = filters.platform  || 'all';
  var product   = filters.product   || 'all';
  var adType    = filters.adType    || 'all';

  // Normalize end date to end of day
  endDate.setHours(23, 59, 59, 999);
  startDate.setHours(0, 0, 0, 0);

  logDebug("Parsed date filter range: " + formatDateJS(startDate) + " to " + formatDateJS(endDate));

  var metaRows   = [];
  var googleRows = [];

  if (platform === 'all' || platform === 'meta') {
    logDebug("Reading rows from META_DATA...");
    metaRows = readSheetRows(ss, META_SHEET, startDate, endDate, product, adType);
    logDebug("Meta rows loaded: " + metaRows.length);
  }
  if (platform === 'all' || platform === 'google') {
    logDebug("Reading rows from GOOGLE_DATA...");
    googleRows = readSheetRows(ss, GOOGLE_SHEET, startDate, endDate, product, adType);
    logDebug("Google rows loaded: " + googleRows.length);
  }

  var allRows = metaRows.concat(googleRows);
  logDebug("Combined Meta & Google rows count: " + allRows.length);

  logDebug("Aggregating overall, Meta and Google KPIs...");
  var kpis = aggregateKPIs(allRows);
  var metaKPIs = aggregateKPIs(metaRows);
  var googleKPIs = aggregateKPIs(googleRows);

  logDebug("Building time series datasets...");
  var timeSeries = buildTimeSeries(allRows, startDate, endDate);
  var metaTimeSeries = buildTimeSeries(metaRows, startDate, endDate);
  var googleTimeSeries = buildTimeSeries(googleRows, startDate, endDate);

  logDebug("Reading CONFIG sheet thresholds...");
  var config = getConfig();

  logDebug("Aggregating by campaign and creatives...");
  var campaignData = aggregateByCampaign(allRows, config);
  var adSetData = aggregateByAdSet(allRows, config);
  var adData = aggregateByAd(allRows, config);
  var productData = aggregateByProduct(allRows);
  var adTypeData = aggregateByAdType(allRows);

  logDebug("Generating intelligence signals...");
  var intelligence = generateIntelligence(campaignData, adData, productData, adTypeData, kpis, config, allRows);

  // Extract unique products and ad types in-memory
  var productSet = {};
  var adTypeSet = {};
  allRows.forEach(function(r) {
    if (r.product) productSet[r.product] = 1;
    if (r.adType) adTypeSet[r.adType] = 1;
  });
  var products = Object.keys(productSet).sort();
  var adTypes  = Object.keys(adTypeSet).sort();

  // Advanced features calculation with error boundaries
  var cohortData = [];
  logDebug("Calculating Cohort Data...");
  try {
    cohortData = buildCohortData(ss);
    logDebug("Cohort Data calculated successfully: " + cohortData.length + " cohorts found");
  } catch(e) {
    logDebug("⚠️ Error during Cohort calculation: " + e.toString());
  }

  var budgetRecommendations = { actions: [], totalWastedSpend: 0, potentialOptimizedRevenue: 0 };
  logDebug("Calculating Budget Recommendations...");
  try {
    budgetRecommendations = buildRecommendationsOptimized(campaignData, config);
    logDebug("Budget recommendations loaded successfully");
  } catch(e) {
    logDebug("⚠️ Error during Budget Optimizer calculation: " + e.toString());
  }

  var fatigueData = [];
  logDebug("Calculating Ad Fatigue metrics...");
  try {
    fatigueData = buildFatigueData(allRows, adData);
    logDebug("Fatigue data calculated successfully: " + fatigueData.length + " ads processed");
  } catch(e) {
    logDebug("⚠️ Error during Fatigue calculation: " + e.toString());
  }

  // Extract product details from PRODUCT_MASTER sheet
  var productDetails = {};
  var prodSheet = ss.getSheetByName(PRODUCT_SHEET);
  if (prodSheet && prodSheet.getLastRow() > 1) {
    var prodData = prodSheet.getDataRange().getValues();
    for (var i = 1; i < prodData.length; i++) {
      var r = prodData[i];
      if (r.length >= 3) {
        var pName = String(r[0]).trim();
        productDetails[pName] = {
          price: cleanParseFloat(r[1]) || 0,
          cogs: cleanParseFloat(r[2]) || 0,
          margin: cleanParseFloat(r[3]) || 0,
          targetCpa: cleanParseFloat(r[4]) || 0
        };
      }
    }
  }

  // Calculate Shopify revenue & order count (if sales log is loaded)
  var totalShopifyRevenue = 0;
  var uniqueShopifyOrders = {};
  var dailyShopifyRevenue = {};
  var dailyCOGS = {};
  var totalCOGS = 0;

  var salesSheet = ss.getSheetByName(SALES_SHEET);
  var salesGridLoaded = false;
  if (salesSheet && salesSheet.getLastRow() > 1) {
    var salesData = salesSheet.getDataRange().getValues();
    var hasData = false;
    for (var i = 1; i < salesData.length; i++) {
      if (salesData[i] && salesData[i].length > 0 && salesData[i][0]) {
        hasData = true;
        break;
      }
    }
    if (hasData) {
      salesGridLoaded = true;
      for (var i = 1; i < salesData.length; i++) {
        var row = salesData[i];
        if (row.length === 0 || !row[0]) continue;
        var orderId = String(row[0]).trim();
        var oDateRaw = row[2];
        var orderRev = cleanParseFloat(row[3]) || 0;
        var pName = String(row[4]).trim();
        
        var oDate = (oDateRaw instanceof Date) ? oDateRaw : new Date(oDateRaw);
        var oDs = formatDateJS(oDate);

        totalShopifyRevenue += orderRev;
        uniqueShopifyOrders[orderId] = true;

        if (!dailyShopifyRevenue[oDs]) dailyShopifyRevenue[oDs] = 0;
        dailyShopifyRevenue[oDs] += orderRev;

        var oCogs = (productDetails[pName] ? productDetails[pName].cogs : 0);
        if (!dailyCOGS[oDs]) dailyCOGS[oDs] = 0;
        dailyCOGS[oDs] += oCogs;
        totalCOGS += oCogs;
      }
    }
  }
  var totalShopifyOrders = Object.keys(uniqueShopifyOrders).length;

  if (!salesGridLoaded) {
    // Estimate daily and total COGS from ad conversions
    allRows.forEach(function(r) {
      var ds = r.dateStr;
      var orderCogs = (productDetails[r.product] ? productDetails[r.product].cogs : 0);
      var rowCogs = r.conversions * orderCogs;
      if (!dailyCOGS[ds]) dailyCOGS[ds] = 0;
      dailyCOGS[ds] += rowCogs;
      totalCOGS += rowCogs;
    });
  }

  // Lookback window optimizer & creator leaderboard (Backend calculations)
  var maxTs = 0;
  allRows.forEach(function(r) {
    var t = r.date.getTime();
    if (t > maxTs) maxTs = t;
  });
  var maxDate = maxTs > 0 ? new Date(maxTs) : new Date();
  maxDate.setHours(23, 59, 59, 999);

  var ts3d = new Date(maxDate); ts3d.setDate(maxDate.getDate() - 3); ts3d.setHours(0,0,0,0);
  var ts7d = new Date(maxDate); ts7d.setDate(maxDate.getDate() - 7); ts7d.setHours(0,0,0,0);
  var ts10d = new Date(maxDate); ts10d.setDate(maxDate.getDate() - 10); ts10d.setHours(0,0,0,0);

  var adSetsMap = {};
  var adsMap = {};
  var creatorsMap = {};

  allRows.forEach(function(r) {
    var rTime = r.date.getTime();
    
    // Ad Set
    var adSetKey = r.platform + '||' + r.adSet;
    if (!adSetsMap[adSetKey]) {
      adSetsMap[adSetKey] = {
        platform: r.platform,
        adSet: r.adSet,
        spend_3d: 0, rev_3d: 0, conv_3d: 0,
        spend_7d: 0, rev_7d: 0, conv_7d: 0,
        spend_10d: 0, rev_10d: 0, conv_10d: 0
      };
    }
    var as = adSetsMap[adSetKey];
    if (rTime >= ts3d.getTime()) {
      as.spend_3d += r.spend; as.rev_3d += r.revenue; as.conv_3d += r.conversions;
    }
    if (rTime >= ts7d.getTime()) {
      as.spend_7d += r.spend; as.rev_7d += r.revenue; as.conv_7d += r.conversions;
    }
    if (rTime >= ts10d.getTime()) {
      as.spend_10d += r.spend; as.rev_10d += r.revenue; as.conv_10d += r.conversions;
    }

    // Ad Creative
    var adKey = r.platform + '||' + r.adName;
    if (!adsMap[adKey]) {
      adsMap[adKey] = {
        platform: r.platform,
        adName: r.adName,
        product: r.product,
        adType: r.adType,
        typeOfAd: r.typeOfAd,
        influencerName: r.influencerName,
        spend_3d: 0, rev_3d: 0, conv_3d: 0, imp_3d: 0, clicks_3d: 0, vViews_3d: 0, thruPlay_3d: 0, hookCount_3d: 0,
        spend_7d: 0, rev_7d: 0, conv_7d: 0, imp_7d: 0, clicks_7d: 0, vViews_7d: 0, thruPlay_7d: 0, hookCount_7d: 0,
        spend_10d: 0, rev_10d: 0, conv_10d: 0, vViews_10d: 0, thruPlay_10d: 0, hookCount_10d: 0
      };
    }
    var ad = adsMap[adKey];
    if (rTime >= ts3d.getTime()) {
      ad.spend_3d += r.spend; ad.rev_3d += r.revenue; ad.conv_3d += r.conversions;
      ad.imp_3d += r.impressions; ad.clicks_3d += r.clicks;
      ad.vViews_3d += r.videoViews;
      ad.thruPlay_3d += r.videoViews * (r.holdRate / 100);
      ad.hookCount_3d += r.videoViews * (r.hookRate / 100);
    }
    if (rTime >= ts7d.getTime()) {
      ad.spend_7d += r.spend; ad.rev_7d += r.revenue; ad.conv_7d += r.conversions;
      ad.imp_7d += r.impressions; ad.clicks_7d += r.clicks;
      ad.vViews_7d += r.videoViews;
      ad.thruPlay_7d += r.videoViews * (r.holdRate / 100);
      ad.hookCount_7d += r.videoViews * (r.hookRate / 100);
    }
    if (rTime >= ts10d.getTime()) {
      ad.spend_10d += r.spend; ad.rev_10d += r.revenue; ad.conv_10d += r.conversions;
      ad.vViews_10d += r.videoViews;
      ad.thruPlay_10d += r.videoViews * (r.holdRate / 100);
      ad.hookCount_10d += r.videoViews * (r.hookRate / 100);
    }

    // Creator Leaderboard
    var creator = r.influencerName || 'In-House';
    var creatorKey = creator;
    if (!creatorsMap[creatorKey]) {
      creatorsMap[creatorKey] = {
        name: creator,
        adType: r.adType,
        spend: 0, revenue: 0,
        vViews: 0, thruPlay: 0, hookCount: 0,
        adsList: {}
      };
    }
    var cr = creatorsMap[creatorKey];
    cr.spend += r.spend;
    cr.revenue += r.revenue;
    cr.vViews += r.videoViews;
    cr.thruPlay += r.videoViews * (r.holdRate / 100);
    cr.hookCount += r.videoViews * (r.hookRate / 100);
    cr.adsList[r.adName] = true;
  });

  var minTargetRoas = config.stopRoas || 1.5;
  var scaleTargetRoas = config.scaleRoas || 4.0;
  var targetHookRate = 30.0;
  var targetHoldRate = 20.0;

  var optimizerAdSets = Object.keys(adSetsMap).map(function(k) {
    var as = adSetsMap[k];
    as.roas_3d = as.spend_3d > 0 ? round2(as.rev_3d / as.spend_3d) : 0;
    as.roas_7d = as.spend_7d > 0 ? round2(as.rev_7d / as.spend_7d) : 0;

    var trend = 'Stable ➡️';
    if (as.spend_3d > 0 && as.spend_7d > 0) {
      if (as.roas_3d > as.roas_7d * 1.05) trend = 'Rising 📈';
      else if (as.roas_3d < as.roas_7d * 0.95) trend = 'Falling 📉';
    } else if (as.spend_3d > 0) {
      trend = 'New 🆕';
    }

    var rec = 'Hold & Watch ⏸️';
    if (as.spend_3d <= 10) {
      rec = 'Inactive 💤';
    } else if (as.roas_3d >= scaleTargetRoas && trend !== 'Falling 📉') {
      rec = 'Scale Spend 🚀';
    } else if (as.roas_3d < minTargetRoas && as.roas_7d < minTargetRoas) {
      rec = 'Kill / Pause 🛑';
    } else if (as.roas_3d < minTargetRoas) {
      rec = 'Soft Kill ⚠️';
    }

    return {
      platform: as.platform,
      name: as.adSet,
      spend: [round2(as.spend_3d), round2(as.spend_7d)],
      roas: [as.roas_3d, as.roas_7d],
      trend: trend,
      recommendation: rec
    };
  }).sort(function(a, b) { return b.spend[1] - a.spend[1]; });

  var optimizerAds = Object.keys(adsMap).map(function(k) {
    var ad = adsMap[k];
    ad.roas_3d = ad.spend_3d > 0 ? round2(ad.rev_3d / ad.spend_3d) : 0;
    ad.roas_7d = ad.spend_7d > 0 ? round2(ad.rev_7d / ad.spend_7d) : 0;
    
    var ctr_3d = ad.imp_3d > 0 ? round2((ad.clicks_3d / ad.imp_3d) * 100) : 0;
    var hook_3d = ad.vViews_3d > 0 ? round2((ad.hookCount_3d / ad.vViews_3d) * 100) : 0;
    var hold_3d = ad.vViews_3d > 0 ? round2((ad.thruPlay_3d / ad.vViews_3d) * 100) : 0;
    
    var ctr_7d = ad.imp_7d > 0 ? round2((ad.clicks_7d / ad.imp_7d) * 100) : 0;
    var hook_7d = ad.vViews_7d > 0 ? round2((ad.hookCount_7d / ad.vViews_7d) * 100) : 0;
    var hold_7d = ad.vViews_7d > 0 ? round2((ad.thruPlay_7d / ad.vViews_7d) * 100) : 0;

    ad.hook_10d = ad.vViews_10d > 0 ? round2((ad.hookCount_10d / ad.vViews_10d) * 100) : 0;
    ad.hold_10d = ad.vViews_10d > 0 ? round2((ad.thruPlay_10d / ad.vViews_10d) * 100) : 0;

    var trend = 'Stable ➡️';
    if (ad.spend_3d > 0 && ad.spend_7d > 0) {
      if (ad.roas_3d > ad.roas_7d * 1.05) trend = 'Rising 📈';
      else if (ad.roas_3d < ad.roas_7d * 0.95) trend = 'Falling 📉';
    } else if (ad.spend_3d > 0) {
      trend = 'New 🆕';
    }

    var rec = 'Hold & Watch ⏸️';
    if (ad.spend_3d <= 10) {
      rec = 'Inactive 💤';
    } else if (ad.roas_3d >= scaleTargetRoas && trend !== 'Falling 📉') {
      rec = 'Scale Budget 🚀';
    } else if (ad.roas_3d < minTargetRoas && ad.roas_7d < minTargetRoas) {
      rec = 'Kill / Pause 🛑';
    } else if (ad.roas_3d < minTargetRoas) {
      if (ad.vViews_10d > 50) {
        if (ad.hook_10d < targetHookRate) {
          rec = 'Fix Hook 🪝';
        } else if (ad.hold_10d < targetHoldRate) {
          rec = 'Fix Body 🔄';
        } else {
          rec = 'Kill / Pause 🛑';
        }
      } else {
        rec = 'Soft Kill ⚠️';
      }
    }

    return {
      platform: ad.platform,
      name: ad.adName,
      product: ad.product,
      adType: ad.adType,
      typeOfAd: ad.typeOfAd,
      influencerName: ad.influencerName,
      spend: [round2(ad.spend_3d), round2(ad.spend_7d)],
      roas: [ad.roas_3d, ad.roas_7d],
      revenue: [round2(ad.rev_3d), round2(ad.rev_7d)],
      conversions: [Math.round(ad.conv_3d), Math.round(ad.conv_7d)],
      impressions: [Math.round(ad.imp_3d), Math.round(ad.imp_7d)],
      clicks: [Math.round(ad.clicks_3d), Math.round(ad.clicks_7d)],
      ctr: [ctr_3d, ctr_7d],
      hook: [hook_3d, hook_7d],
      hold: [hold_3d, hold_7d],
      trend: trend,
      recommendation: rec
    };
  }).sort(function(a, b) { return b.spend[1] - a.spend[1]; });

  var creatorLeaderboard = Object.keys(creatorsMap).map(function(k) {
    var cr = creatorsMap[k];
    cr.roas = cr.spend > 0 ? round2(cr.revenue / cr.spend) : 0;
    cr.hook = cr.vViews > 0 ? round2((cr.hookCount / cr.vViews) * 100) : 0;
    cr.hold = cr.vViews > 0 ? round2((cr.thruPlay / cr.vViews) * 100) : 0;
    
    var type = 'Internal UGC';
    if (cr.name.indexOf('@') === 0 || cr.adType === 'Influencer') {
      type = 'Influencer';
    }
    
    return {
      name: cr.name,
      type: type,
      spend: round2(cr.spend),
      revenue: round2(cr.revenue),
      roas: cr.roas,
      hook: cr.hook,
      hold: cr.hold,
      activeAds: Object.keys(cr.adsList).length
    };
  }).sort(function(a, b) { return b.revenue - a.revenue; });

  // ─── TOP / BOTTOM AD SETS, NARRATIVES, LANGUAGES, AND PRODUCT TRENDS ───
  var adSetPerfMap = {};
  allRows.forEach(function(r) {
    if (!r.adSet) return;
    var key = r.platform + '||' + r.adSet;
    if (!adSetPerfMap[key]) {
      adSetPerfMap[key] = {
        platform: r.platform,
        adSetName: r.adSet,
        spend: 0,
        revenue: 0,
        conversions: 0
      };
    }
    adSetPerfMap[key].spend += r.spend;
    adSetPerfMap[key].revenue += r.revenue;
    adSetPerfMap[key].conversions += r.conversions;
  });

  var adSetPerfList = Object.keys(adSetPerfMap).map(function(k) {
    var item = adSetPerfMap[k];
    item.roas = item.spend > 0 ? item.revenue / item.spend : 0;
    return item;
  });

  var filteredAdSets = adSetPerfList.filter(function(item) {
    return item.spend > 500;
  });

  var topAdSets = filteredAdSets.slice().sort(function(a, b) {
    return b.roas - a.roas;
  }).slice(0, 5).map(function(item) {
    return {
      platform: item.platform,
      adSetName: item.adSetName,
      spend: round2(item.spend),
      revenue: round2(item.revenue),
      roas: round2(item.roas),
      conversions: Math.round(item.conversions)
    };
  });

  var bottomAdSets = filteredAdSets.slice().sort(function(a, b) {
    return a.roas - b.roas;
  }).slice(0, 5).map(function(item) {
    return {
      platform: item.platform,
      adSetName: item.adSetName,
      spend: round2(item.spend),
      revenue: round2(item.revenue),
      roas: round2(item.roas),
      conversions: Math.round(item.conversions)
    };
  });

  // Lookback Leaderboards for Ad Sets
  var topAdSets3d = optimizerAdSets.slice().filter(function(as) { return as.spend[0] > 10; }).sort(function(a, b) { return b.roas[0] - a.roas[0]; }).slice(0, 5);
  var bottomAdSets3d = optimizerAdSets.slice().filter(function(as) { return as.spend[0] > 10; }).sort(function(a, b) { return a.roas[0] - b.roas[0]; }).slice(0, 5);
  
  var topAdSets7d = optimizerAdSets.slice().filter(function(as) { return as.spend[1] > 10; }).sort(function(a, b) { return b.roas[1] - a.roas[1]; }).slice(0, 5);
  var bottomAdSets7d = optimizerAdSets.slice().filter(function(as) { return as.spend[1] > 10; }).sort(function(a, b) { return a.roas[1] - b.roas[1]; }).slice(0, 5);

  // Lookback Leaderboards for Ads
  var topAds3d = optimizerAds.slice().filter(function(ad) { return ad.spend[0] > 10; }).sort(function(a, b) { return b.roas[0] - a.roas[0]; }).slice(0, 5);
  var bottomAds3d = optimizerAds.slice().filter(function(ad) { return ad.spend[0] > 10; }).sort(function(a, b) { return a.roas[0] - b.roas[0]; }).slice(0, 5);

  var topAds7d = optimizerAds.slice().filter(function(ad) { return ad.spend[1] > 10; }).sort(function(a, b) { return b.roas[1] - a.roas[1]; }).slice(0, 5);
  var bottomAds7d = optimizerAds.slice().filter(function(ad) { return ad.spend[1] > 10; }).sort(function(a, b) { return a.roas[1] - b.roas[1]; }).slice(0, 5);

  var narrativePerfMap = {};
  allRows.forEach(function(r) {
    var narrative = r.narrative || 'Brand Story';
    if (!narrativePerfMap[narrative]) {
      narrativePerfMap[narrative] = {
        narrative: narrative,
        spend: 0,
        revenue: 0,
        conversions: 0
      };
    }
    narrativePerfMap[narrative].spend += r.spend;
    narrativePerfMap[narrative].revenue += r.revenue;
    narrativePerfMap[narrative].conversions += r.conversions;
  });

  var targetRoas = config.roas || 3.0;

  var narrativeStats = Object.keys(narrativePerfMap).map(function(n) {
    var item = narrativePerfMap[n];
    var roas = item.spend > 0 ? item.revenue / item.spend : 0;
    var status = roas >= targetRoas ? 'Working ✅' : 'Not Working ❌';
    if (item.spend === 0) status = 'No Spend 💤';
    return {
      narrative: n,
      spend: round2(item.spend),
      revenue: round2(item.revenue),
      roas: round2(roas),
      conversions: Math.round(item.conversions),
      status: status
    };
  }).sort(function(a, b) { return b.spend - a.spend; });

  var languagePerfMap = {};
  allRows.forEach(function(r) {
    var language = r.language || 'English';
    if (!languagePerfMap[language]) {
      languagePerfMap[language] = {
        language: language,
        spend: 0,
        revenue: 0,
        conversions: 0
      };
    }
    languagePerfMap[language].spend += r.spend;
    languagePerfMap[language].revenue += r.revenue;
    languagePerfMap[language].conversions += r.conversions;
  });

  var languageStats = Object.keys(languagePerfMap).map(function(l) {
    var item = languagePerfMap[l];
    var roas = item.spend > 0 ? item.revenue / item.spend : 0;
    var status = roas >= targetRoas ? 'Working ✅' : 'Not Working ❌';
    if (item.spend === 0) status = 'No Spend 💤';
    return {
      language: l,
      spend: round2(item.spend),
      revenue: round2(item.revenue),
      roas: round2(roas),
      conversions: Math.round(item.conversions),
      status: status
    };
  }).sort(function(a, b) { return b.spend - a.spend; });

  var prodDailyMap = {};
  allRows.forEach(function(r) {
    var dStr = r.dateStr;
    var prod = r.product || 'Other Product';
    var key = dStr + '||' + prod;
    if (!prodDailyMap[key]) {
      prodDailyMap[key] = {
        date: dStr,
        product: prod,
        spend: 0,
        revenue: 0,
        conversions: 0
      };
    }
    prodDailyMap[key].spend += r.spend;
    prodDailyMap[key].revenue += r.revenue;
    prodDailyMap[key].conversions += r.conversions;
  });

  var productDailyTrend = Object.keys(prodDailyMap).map(function(k) {
    var item = prodDailyMap[k];
    var roas = item.spend > 0 ? item.revenue / item.spend : 0;
    return {
      date: item.date,
      product: item.product,
      spend: round2(item.spend),
      revenue: round2(item.revenue),
      roas: round2(roas),
      conversions: Math.round(item.conversions)
    };
  });
  
  productDailyTrend.sort(function(a, b) {
    var dComp = b.date.localeCompare(a.date);
    if (dComp !== 0) return dComp;
    return a.product.localeCompare(b.product);
  });

  logDebug("Populating DAILY_STATS sheet...");
  try {
    writeDailyStats(ss, allRows);
  } catch(e) {
    logDebug("Error populating DAILY_STATS: " + e.toString());
  }

  logDebug("getDashboardData execution successfully completed");

  return {
    kpis: kpis,
    metaKPIs: metaKPIs,
    googleKPIs: googleKPIs,
    timeSeries: timeSeries,
    metaTimeSeries: metaTimeSeries,
    googleTimeSeries: googleTimeSeries,
    campaignData: campaignData,
    adSetData: adSetData,
    adData: adData,
    productData: productData,
    adTypeData: adTypeData,
    intelligence: intelligence,
    cohortData: cohortData,
    budgetRecommendations: budgetRecommendations,
    fatigueData: fatigueData,
    optimizerAdSets: optimizerAdSets,
    optimizerAds: optimizerAds,
    creatorLeaderboard: creatorLeaderboard,
    filterOptions: { products: products, adTypes: adTypes },
    dateRange: {
      start: formatDateJS(startDate),
      end:   formatDateJS(endDate)
    },
    totalShopifyRevenue: totalShopifyRevenue,
    totalShopifyOrders: totalShopifyOrders,
    totalCOGS: totalCOGS,
    dailyCOGS: dailyCOGS,
    dailyShopifyRevenue: dailyShopifyRevenue,
    topAdSets: topAdSets,
    bottomAdSets: bottomAdSets,
    topAdSets3d: topAdSets3d,
    bottomAdSets3d: bottomAdSets3d,
    topAdSets7d: topAdSets7d,
    bottomAdSets7d: bottomAdSets7d,
    topAds3d: topAds3d,
    bottomAds3d: bottomAds3d,
    topAds7d: topAds7d,
    bottomAds7d: bottomAds7d,
    narrativeStats: narrativeStats,
    languageStats: languageStats,
    productDailyTrend: productDailyTrend
  };
}

function detectDateExtent(ss) {
  var minTs = null;
  var maxTs = null;
  var sheets = [ss.getSheetByName(META_SHEET), ss.getSheetByName(GOOGLE_SHEET)];
  
  sheets.forEach(function(sheet) {
    if (!sheet || sheet.getLastRow() <= 1) return;
    
    // Find the 'Date' column index
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var dateIdx = -1;
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim() === 'Date') {
        dateIdx = i + 1;
        break;
      }
    }
    
    if (dateIdx !== -1) {
      var data = sheet.getRange(2, dateIdx, sheet.getLastRow() - 1, 1).getValues();
      for (var j = 0; j < data.length; j++) {
        var val = data[j][0];
        if (val) {
          var d = new Date(val);
          var ts = d.getTime();
          if (!isNaN(ts)) {
            if (minTs === null || ts < minTs) minTs = ts;
            if (maxTs === null || ts > maxTs) maxTs = ts;
          }
        }
      }
    }
  });
  
  if (minTs === null || maxTs === null) return null;
  
  var minDate = new Date(minTs);
  var maxDate = new Date(maxTs);
  
  minDate.setHours(0, 0, 0, 0);
  maxDate.setHours(23, 59, 59, 999);
  
  return { min: minDate, max: maxDate };
}

function writeDailyStats(ss, allRows) {
  var sheet = ss.getSheetByName(DAILY_SHEET);
  if (!sheet) return;
  
  var dailyMap = {};
  allRows.forEach(function(r) {
    var dStr = r.dateStr;
    if (!dailyMap[dStr]) {
      dailyMap[dStr] = {
        spend: 0,
        conversions: 0,
        revenue: 0
      };
    }
    dailyMap[dStr].spend += r.spend;
    dailyMap[dStr].conversions += r.conversions;
    dailyMap[dStr].revenue += r.revenue;
  });
  
  var dailyRows = Object.keys(dailyMap).map(function(dStr) {
    var data = dailyMap[dStr];
    var cac = data.conversions > 0 ? data.spend / data.conversions : 0;
    var aov = data.conversions > 0 ? data.revenue / data.conversions : 0;
    var roas = data.spend > 0 ? data.revenue / data.spend : 0;
    return [
      dStr,
      round2(data.spend),
      Math.round(data.conversions),
      round2(cac),
      round2(aov),
      round2(roas)
    ];
  });
  
  dailyRows.sort(function(a, b) {
    return b[0].localeCompare(a[0]);
  });
  
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clearContent();
  }
  
  if (dailyRows.length > 0) {
    sheet.getRange(2, 1, dailyRows.length, 6).setValues(dailyRows);
    sheet.getRange('B2:B' + (dailyRows.length + 1)).setNumberFormat('₹#,##0.00'); // Spend
    sheet.getRange('C2:C' + (dailyRows.length + 1)).setNumberFormat('#,##0');    // Conversions
    sheet.getRange('D2:E' + (dailyRows.length + 1)).setNumberFormat('₹#,##0.00'); // CAC, AOV
    sheet.getRange('F2:F' + (dailyRows.length + 1)).setNumberFormat('0.00"x"');   // ROAS
  }
}

function buildRecommendationsOptimized(campaignData, config) {
  var actions = [];
  var totalWastedSpend = 0;
  var scaleCandidates = [];

  campaignData.forEach(function(c) {
    if (c.signal === 'STOP') {
      totalWastedSpend += c.spend;
      actions.push({
        type: 'STOP',
        campaign: c.campaign,
        platform: c.platform,
        currentSpend: c.spend,
        currentRoas: c.roas,
        recommendation: 'Pause Campaign',
        message: 'Pause "' + c.campaign + '" (' + c.platform + ') to cut wastage of ₹' + formatNum(c.spend) + '. Current ROAS is ' + c.roas + 'x against 1.5x min target.',
        actionItem: 'Deallocate ₹' + formatNum(c.spend) + ' immediately.'
      });
    } else if (c.signal === 'SCALE') {
      scaleCandidates.push(c);
    }
  });

  scaleCandidates.forEach(function(c) {
    var suggestedIncrease = Math.round(c.spend * 0.25);
    if (suggestedIncrease < 100) suggestedIncrease = 500;
    var projectedRevenue = Math.round(suggestedIncrease * c.roas);
    
    actions.push({
      type: 'SCALE',
      campaign: c.campaign,
      platform: c.platform,
      currentSpend: c.spend,
      currentRoas: c.roas,
      recommendation: 'Increase Budget by 25%',
      message: 'Increase budget of "' + c.campaign + '" (' + c.platform + ') by ₹' + formatNum(suggestedIncrease) + '. Current ROAS is ' + c.roas + 'x and trending ' + c.trend + '.',
      actionItem: 'Allocate ₹' + formatNum(suggestedIncrease) + ' (Est. Gain: +₹' + formatNum(projectedRevenue) + ' revenue).'
    });
  });

  return {
    actions: actions,
    totalWastedSpend: round2(totalWastedSpend),
    potentialOptimizedRevenue: round2(scaleCandidates.reduce(function(sum, c) {
      return sum + (c.spend * 0.25 * c.roas);
    }, 0))
  };
}

// ─── READ SHEET ROWS WITH FILTERS ─────────────────────────
function readSheetRows(ss, sheetName, startDate, endDate, product, adType) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var rows = [];

  // Map column indices
  var colMap = {};
  headers.forEach(function(h, i) { colMap[h] = i; });

  var getIndex = function(arr) {
    for (var j = 0; j < arr.length; j++) {
      if (colMap[arr[j]] !== undefined) return colMap[arr[j]];
    }
    return undefined;
  };

  var startTs = startDate.getTime();
  var endTs   = endDate.getTime();
  
  var dateColIndex = getIndex(['Date', 'Day']);
  if (dateColIndex === undefined) return [];

  // Dynamic header mappings (old vs new)
  var prodCol = getIndex(['Product Name', 'Product']);
  var typeCol = getIndex(['Ad Type']);
  var formatCol = getIndex(['Type of Ad (Static/Reel)', 'Type of Ad', 'Ad Type']);
  var influencerCol = getIndex(['Name of the Influencer', 'Influencer Name', 'Influencer']);
  var narrativeCol = getIndex(['Narrative']);
  var languageCol = getIndex(['Language of the Ad', 'Language']);
  var categoryCol = getIndex(['Category']);

  var ctrCol = getIndex(['CTR %', 'CTR']);
  var cpcCol = getIndex(['CPC']);
  var roasCol = getIndex(['ROAS']);
  var freqCol = getIndex(['Frequency']);
  var cpmCol = getIndex(['CPM']);
  var hookCol = getIndex(['Hook Rate %', 'Hook Rate']);
  var holdCol = getIndex(['Hold Rate %', 'Hold Rate']);

  var spendCol = getIndex(['Spend', 'Amount spent (INR)']);
  var impCol = getIndex(['Impressions']);
  var clickCol = getIndex(['Clicks', 'Link clicks', 'Link Clicks']);
  var linkClicksCol = getIndex(['Link Clicks', 'Link clicks', 'Clicks']);
  var lpvCol = getIndex(['Landing Page Views', 'Landing page views', 'LP Views']);
  var convCol = getIndex(['Conversions', 'Purchases']);
  var revCol = getIndex(['Revenue', 'Purchases conversion value']);
  var atcCol = getIndex(['Add to Cart', 'Add To Cart']);
  
  var vViewsCol = getIndex(['Video Views', 'Video views', '3-second video plays']);
  var thruPlayCol = getIndex(['ThruPlay', 'Thruplay', 'Video plays at 25%']);
  var campCol = getIndex(['Campaign', 'Campaign name']);
  var adsetCol = getIndex(['Ad Set', 'Ad Group', 'Ad set name']);
  var adCol = getIndex(['Ad Name', 'Ad name']);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row.length === 0 || !row[0]) continue;
    var rowDateRaw = row[dateColIndex];
    if (!rowDateRaw) continue;

    var rowDate = (rowDateRaw instanceof Date) ? rowDateRaw : new Date(rowDateRaw);
    var rowTs = rowDate.getTime();
    if (isNaN(rowTs)) continue;
    if (rowTs < startTs || rowTs > endTs) continue;

    var rowProduct = prodCol !== undefined ? (row[prodCol] || '').toString().trim() : '';
    var rowAdType = typeCol !== undefined ? (row[typeCol] || '').toString().trim() : '';
    var rowTypeOfAd = formatCol !== undefined ? (row[formatCol] || '').toString().trim() : '';
    var rowInfluencer = influencerCol !== undefined ? (row[influencerCol] || '').toString().trim() : '';
    var rowNarrative = narrativeCol !== undefined ? (row[narrativeCol] || '').toString().trim() : '';
    var rowLanguage = languageCol !== undefined ? (row[languageCol] || '').toString().trim() : '';
    var rowCategory = categoryCol !== undefined ? (row[categoryCol] || '').toString().trim() : '';

    var campName = campCol !== undefined ? (row[campCol] || '').toString().trim() : '';
    var adSetName = adsetCol !== undefined ? (row[adsetCol] || '').toString().trim() : '';
    var adName = adCol !== undefined ? (row[adCol] || '').toString().trim() : '';

    // Auto-Tagging Fallback if any is blank
    if (!rowProduct || !rowTypeOfAd || !rowInfluencer || !rowNarrative || !rowLanguage) {
      var tags = autoTagMetadata(campName, adSetName, adName);
      if (!rowProduct) rowProduct = tags.product || 'Other Product';
      if (!rowTypeOfAd) rowTypeOfAd = tags.typeOfAd || 'Reel';
      if (!rowInfluencer) rowInfluencer = tags.influencerName || 'In-House';
      if (!rowNarrative) rowNarrative = tags.narrative || 'Brand Story';
      if (!rowLanguage) rowLanguage = tags.language || 'English';
      if (!rowCategory) rowCategory = tags.category || 'Accessories';
    }

    // Always calculate rowAdType since it's a derived field not stored directly in the Sheet headers
    if (!rowAdType) {
      var upperTypeOfAd = (rowTypeOfAd || '').toUpperCase();
      var upperInfluencer = (rowInfluencer || '').toUpperCase();
      
      if (upperTypeOfAd.indexOf('STATIC') !== -1 || upperTypeOfAd.indexOf('IMAGE') !== -1) {
        rowAdType = 'Static Creative';
      } else if (upperTypeOfAd.indexOf('CAROUSEL') !== -1) {
        rowAdType = 'Carousel';
      } else if (upperInfluencer && upperInfluencer !== 'IN-HOUSE' && upperInfluencer !== 'INhouse' && upperInfluencer !== '—' && upperInfluencer !== 'NONE') {
        rowAdType = 'Influencer';
      } else if (upperTypeOfAd.indexOf('VIDEO') !== -1 || upperTypeOfAd.indexOf('REEL') !== -1) {
        rowAdType = 'Video Ad';
      } else {
        rowAdType = 'Internal UGC';
      }
    }

    if (product !== 'all' && rowProduct !== product) continue;
    if (adType !== 'all' && rowAdType !== adType) continue;

    var rowSpend = spendCol !== undefined ? cleanParseFloat(row[spendCol]) : 0;
    var rowImp = impCol !== undefined ? cleanParseFloat(row[impCol]) : 0;
    var rowClicks = clickCol !== undefined ? cleanParseFloat(row[clickCol]) : 0;
    var rowLinkClicks = linkClicksCol !== undefined ? cleanParseFloat(row[linkClicksCol]) : 0;
    var rowLpv = lpvCol !== undefined ? cleanParseFloat(row[lpvCol]) : 0;
    var rowConv = convCol !== undefined ? cleanParseFloat(row[convCol]) : 0;
    var rowRev = revCol !== undefined ? cleanParseFloat(row[revCol]) : 0;
    var rowAtc = atcCol !== undefined ? cleanParseFloat(row[atcCol]) : 0;

    var rowCtr = ctrCol !== undefined ? cleanParseFloat(row[ctrCol]) : (rowImp > 0 ? (rowClicks / rowImp) * 100 : 0);
    var rowCpc = cpcCol !== undefined ? cleanParseFloat(row[cpcCol]) : (rowClicks > 0 ? rowSpend / rowClicks : 0);
    var rowRoas = roasCol !== undefined ? cleanParseFloat(row[roasCol]) : (rowSpend > 0 ? rowRev / rowSpend : 0);
    var rowCpm = cpmCol !== undefined ? cleanParseFloat(row[cpmCol]) : (rowImp > 0 ? (rowSpend / rowImp) * 1000 : 0);
    var rowFreq = freqCol !== undefined ? cleanParseFloat(row[freqCol]) : 1.0;
    
    var raw3s = getIndex(['3-second video plays']) !== undefined ? cleanParseFloat(row[getIndex(['3-second video plays'])]) : 0;
    var raw25 = getIndex(['Video plays at 25%']) !== undefined ? cleanParseFloat(row[getIndex(['Video plays at 25%'])]) : 0;
    
    var rowVideoViews = vViewsCol !== undefined ? cleanParseFloat(row[vViewsCol]) : 0;
    var rowThruPlay = thruPlayCol !== undefined ? cleanParseFloat(row[thruPlayCol]) : 0;
    
    var rowHookRate = hookCol !== undefined ? cleanParseFloat(row[hookCol]) : (rowImp > 0 && raw3s > 0 ? (raw3s / rowImp) * 100 : 0);
    var rowHoldRate = holdCol !== undefined ? cleanParseFloat(row[holdCol]) : (raw3s > 0 && raw25 > 0 ? (raw25 / raw3s) * 100 : 0);

    // Old format fallback if neither new video plays nor rate columns are present
    if (hookCol === undefined && getIndex(['3-second video plays']) === undefined) {
      var oldHook = getIndex(['Hook Rate %', 'Hook Rate']);
      rowHookRate = oldHook !== undefined ? cleanParseFloat(row[oldHook]) : 0;
    }
    if (holdCol === undefined && getIndex(['Video plays at 25%']) === undefined) {
      var oldHold = getIndex(['Hold Rate %', 'Hold Rate']);
      rowHoldRate = oldHold !== undefined ? cleanParseFloat(row[oldHold]) : 0;
    }

    rows.push({
      date:             rowDate,
      dateStr:          formatDateJS(rowDate),
      platform:         sheetName === META_SHEET ? 'Meta' : 'Google',
      campaign:         campName,
      adSet:            adSetName,
      adName:           adName,
      adType:           rowAdType,
      product:          rowProduct,
      typeOfAd:         rowTypeOfAd,
      influencerName:   rowInfluencer,
      narrative:        rowNarrative,
      language:         rowLanguage,
      category:         rowCategory,
      spend:            rowSpend,
      impressions:      rowImp,
      clicks:           rowClicks,
      linkClicks:       rowLinkClicks,
      landingPageViews: rowLpv,
      conversions:      rowConv,
      revenue:          rowRev,
      addToCart:        rowAtc,
      ctr:              rowCtr,
      cpc:              rowCpc,
      roas:             rowRoas,
      frequency:        rowFreq,
      cpm:              rowCpm,
      hookRate:         rowHookRate,
      holdRate:         rowHoldRate,
      videoViews:       rowVideoViews,
      thruPlay:         rowThruPlay
    });
  }
  return rows;
}

// ─── AGGREGATE KPIs ───────────────────────────────────────
function aggregateKPIs(rows) {
  if (!rows.length) return emptyKPIs();

  var totalSpend = 0, totalRevenue = 0, totalClicks = 0, totalLinkClicks = 0;
  var totalLPV = 0, totalConversions = 0, totalImpressions = 0, totalATC = 0;
  var totalVideoViews = 0, totalThruPlay = 0;
  var hookRateSum = 0, holdRateSum = 0, hookCount = 0, holdCount = 0;

  rows.forEach(function(r) {
    totalSpend       += r.spend;
    totalRevenue     += r.revenue;
    totalClicks      += r.clicks;
    totalLinkClicks  += r.linkClicks;
    totalLPV         += r.landingPageViews;
    totalConversions += r.conversions;
    totalImpressions += r.impressions;
    totalATC         += r.addToCart;
    totalVideoViews  += r.videoViews;
    totalThruPlay    += r.thruPlay;
    if (r.hookRate > 0) { hookRateSum += r.hookRate; hookCount++; }
    if (r.holdRate > 0) { holdRateSum += r.holdRate; holdCount++; }
  });

  var roas     = totalSpend > 0 ? totalRevenue / totalSpend : 0;
  var ctr      = totalImpressions > 0 ? (totalClicks / totalImpressions) * 100 : 0;
  var cpc      = totalClicks > 0 ? totalSpend / totalClicks : 0;
  var convRate = totalClicks > 0 ? (totalConversions / totalClicks) * 100 : 0;
  var cpm      = totalImpressions > 0 ? (totalSpend / totalImpressions) * 1000 : 0;
  var cpa      = totalConversions > 0 ? totalSpend / totalConversions : 0;
  // Funnel rates
  var lpvRate  = totalLinkClicks > 0 ? (totalLPV / totalLinkClicks) * 100 : 0;
  var atcRate  = totalLPV > 0 ? (totalATC / totalLPV) * 100 : 0;
  var avgHook  = hookCount > 0 ? hookRateSum / hookCount : 0;
  var avgHold  = holdCount > 0 ? holdRateSum / holdCount : 0;
  var thruRate = totalVideoViews > 0 ? (totalThruPlay / totalVideoViews) * 100 : 0;

  return {
    spend:       round2(totalSpend),
    revenue:     round2(totalRevenue),
    profit:      round2(totalRevenue - totalSpend),
    clicks:      Math.round(totalClicks),
    linkClicks:  Math.round(totalLinkClicks),
    lpv:         Math.round(totalLPV),
    conversions: Math.round(totalConversions),
    impressions: Math.round(totalImpressions),
    addToCart:   Math.round(totalATC),
    roas:        round2(roas),
    ctr:         round2(ctr),
    cpc:         round2(cpc),
    convRate:    round2(convRate),
    cpm:         round2(cpm),
    cpa:         round2(cpa),
    lpvRate:     round2(lpvRate),
    atcRate:     round2(atcRate),
    hookRate:    round2(avgHook),
    holdRate:    round2(avgHold),
    thruRate:    round2(thruRate),
    videoViews:  Math.round(totalVideoViews)
  };
}

function emptyKPIs() {
  return { spend:0, revenue:0, profit:0, clicks:0, linkClicks:0, lpv:0,
           conversions:0, impressions:0, addToCart:0,
           roas:0, ctr:0, cpc:0, convRate:0, cpm:0, cpa:0,
           lpvRate:0, atcRate:0, hookRate:0, holdRate:0, thruRate:0, videoViews:0 };
}

// ─── TIME SERIES ──────────────────────────────────────────
function buildTimeSeries(rows, startDate, endDate) {
  var byDate = {};
  rows.forEach(function(r) {
    if (!byDate[r.dateStr]) {
      byDate[r.dateStr] = { spend:0, revenue:0, clicks:0, linkClicks:0, lpv:0,
                            conversions:0, impressions:0, addToCart:0,
                            hookRateSum:0, holdRateSum:0, hookCount:0, holdCount:0 };
    }
    byDate[r.dateStr].spend       += r.spend;
    byDate[r.dateStr].revenue     += r.revenue;
    byDate[r.dateStr].clicks      += r.clicks;
    byDate[r.dateStr].linkClicks  += r.linkClicks;
    byDate[r.dateStr].lpv         += r.landingPageViews;
    byDate[r.dateStr].conversions += r.conversions;
    byDate[r.dateStr].impressions += r.impressions;
    byDate[r.dateStr].addToCart   += r.addToCart;
    if (r.hookRate > 0) { byDate[r.dateStr].hookRateSum += r.hookRate; byDate[r.dateStr].hookCount++; }
    if (r.holdRate > 0) { byDate[r.dateStr].holdRateSum += r.holdRate; byDate[r.dateStr].holdCount++; }
  });

  var labels = [], spend = [], revenue = [], clicks = [], conversions = [], roas = [];
  var lpv = [], addToCart = [], hookRate = [], holdRate = [], atcRate = [];
  var cur = new Date(startDate);
  while (cur <= endDate) {
    var ds = formatDateJS(cur);
    var d  = byDate[ds] || { spend:0, revenue:0, clicks:0, linkClicks:0, lpv:0,
                              conversions:0, impressions:0, addToCart:0,
                              hookRateSum:0, holdRateSum:0, hookCount:0, holdCount:0 };
    labels.push(ds);
    spend.push(round2(d.spend));
    revenue.push(round2(d.revenue));
    clicks.push(Math.round(d.clicks));
    conversions.push(Math.round(d.conversions));
    roas.push(d.spend > 0 ? round2(d.revenue / d.spend) : 0);
    lpv.push(Math.round(d.lpv));
    addToCart.push(Math.round(d.addToCart));
    hookRate.push(d.hookCount > 0 ? round2(d.hookRateSum / d.hookCount) : 0);
    holdRate.push(d.holdCount > 0 ? round2(d.holdRateSum / d.holdCount) : 0);
    atcRate.push(d.lpv > 0 ? round2((d.addToCart / d.lpv) * 100) : 0);
    cur.setDate(cur.getDate() + 1);
  }
  return { labels:labels, spend:spend, revenue:revenue, clicks:clicks,
           conversions:conversions, roas:roas, lpv:lpv, addToCart:addToCart,
           hookRate:hookRate, holdRate:holdRate, atcRate:atcRate };
}

// ─── AGGREGATE BY CAMPAIGN ────────────────────────────────
function aggregateByCampaign(rows, config) {
  var map = {};
  rows.forEach(function(r) {
    var key = r.platform + '||' + r.campaign;
    if (!map[key]) {
      map[key] = { platform:r.platform, campaign:r.campaign, adSetsMap:{},
                   spend:0, revenue:0, clicks:0, linkClicks:0, lpv:0,
                   conversions:0, impressions:0, addToCart:0,
                   hookRateSum:0, holdRateSum:0, hookCount:0, holdCount:0, rows:[] };
    }
    map[key].spend       += r.spend;
    map[key].revenue     += r.revenue;
    map[key].clicks      += r.clicks;
    map[key].linkClicks  += r.linkClicks;
    map[key].lpv         += r.landingPageViews;
    map[key].conversions += r.conversions;
    map[key].impressions += r.impressions;
    map[key].addToCart   += r.addToCart;
    if (r.hookRate > 0) { map[key].hookRateSum += r.hookRate; map[key].hookCount++; }
    if (r.holdRate > 0) { map[key].holdRateSum += r.holdRate; map[key].holdCount++; }
    map[key].adSetsMap[r.adSet] = 1;
    map[key].rows.push(r);
  });

  config = config || getConfig();
  var results = [];
  Object.keys(map).forEach(function(k) {
    var c = map[k];
    var roas     = c.spend > 0 ? c.revenue / c.spend : 0;
    var ctr      = c.impressions > 0 ? (c.clicks / c.impressions) * 100 : 0;
    var convRate = c.clicks > 0 ? (c.conversions / c.clicks) * 100 : 0;
    var cpc      = c.clicks > 0 ? c.spend / c.clicks : 0;
    var atcRate  = c.lpv > 0 ? (c.addToCart / c.lpv) * 100 : 0;
    var lpvRate  = c.linkClicks > 0 ? (c.lpv / c.linkClicks) * 100 : 0;
    var avgHook  = c.hookCount > 0 ? c.hookRateSum / c.hookCount : 0;
    var avgHold  = c.holdCount > 0 ? c.holdRateSum / c.holdCount : 0;
    var signal   = computeSignal(roas, ctr, cpc, convRate, c.spend, config);
    var trend    = computeTrend(c.rows);

    results.push({
      platform:    c.platform,
      campaign:    c.campaign,
      adSets:      Object.keys(c.adSetsMap).length,
      spend:       round2(c.spend),
      revenue:     round2(c.revenue),
      roas:        round2(roas),
      ctr:         round2(ctr),
      cpc:         round2(cpc),
      convRate:    round2(convRate),
      conversions: Math.round(c.conversions),
      impressions: Math.round(c.impressions),
      linkClicks:  Math.round(c.linkClicks),
      lpv:         Math.round(c.lpv),
      addToCart:   Math.round(c.addToCart),
      atcRate:     round2(atcRate),
      lpvRate:     round2(lpvRate),
      hookRate:    round2(avgHook),
      holdRate:    round2(avgHold),
      signal:      signal,
      trend:       trend
    });
  });

  results.sort(function(a, b) { return b.roas - a.roas; });
  return results;
}

// ─── AGGREGATE BY AD SET ──────────────────────────────────
function aggregateByAdSet(rows, config) {
  var map = {};
  rows.forEach(function(r) {
    if (!r.adSet) return;
    var key = r.platform + '||' + r.adSet;
    if (!map[key]) {
      map[key] = { platform:r.platform, adSet:r.adSet, campaign:r.campaign,
                   spend:0, revenue:0, clicks:0, linkClicks:0, lpv:0,
                   conversions:0, impressions:0, addToCart:0,
                   hookRateSum:0, holdRateSum:0, hookCount:0, holdCount:0, rows:[] };
    }
    map[key].spend       += r.spend;
    map[key].revenue     += r.revenue;
    map[key].clicks      += r.clicks;
    map[key].linkClicks  += r.linkClicks;
    map[key].lpv         += r.landingPageViews;
    map[key].conversions += r.conversions;
    map[key].impressions += r.impressions;
    map[key].addToCart   += r.addToCart;
    if (r.hookRate > 0) { map[key].hookRateSum += r.hookRate; map[key].hookCount++; }
    if (r.holdRate > 0) { map[key].holdRateSum += r.holdRate; map[key].holdCount++; }
    map[key].rows.push(r);
  });

  config = config || getConfig();
  var results = [];
  Object.keys(map).forEach(function(k) {
    var as = map[k];
    var roas     = as.spend > 0 ? as.revenue / as.spend : 0;
    var ctr      = as.impressions > 0 ? (as.clicks / as.impressions) * 100 : 0;
    var convRate = as.clicks > 0 ? (as.conversions / as.clicks) * 100 : 0;
    var cpc      = as.clicks > 0 ? as.spend / as.clicks : 0;
    var atcRate  = as.lpv > 0 ? (as.addToCart / as.lpv) * 100 : 0;
    var lpvRate  = as.linkClicks > 0 ? (as.lpv / as.linkClicks) * 100 : 0;
    var avgHook  = as.hookCount > 0 ? as.hookRateSum / as.hookCount : 0;
    var avgHold  = as.holdCount > 0 ? as.holdRateSum / as.holdCount : 0;
    var signal   = computeSignal(roas, ctr, cpc, convRate, as.spend, config);
    var trend    = computeTrend(as.rows);

    results.push({
      platform:    as.platform,
      adSet:       as.adSet,
      campaign:    as.campaign,
      spend:       round2(as.spend),
      revenue:     round2(as.revenue),
      roas:        round2(roas),
      ctr:         round2(ctr),
      cpc:         round2(cpc),
      convRate:    round2(convRate),
      conversions: Math.round(as.conversions),
      impressions: Math.round(as.impressions),
      linkClicks:  Math.round(as.linkClicks),
      lpv:         Math.round(as.lpv),
      addToCart:   Math.round(as.addToCart),
      atcRate:     round2(atcRate),
      lpvRate:     round2(lpvRate),
      hookRate:    round2(avgHook),
      holdRate:    round2(avgHold),
      profit:      round2(as.revenue - as.spend),
      signal:      signal,
      trend:       trend
    });
  });

  results.sort(function(a, b) { return b.spend - a.spend; });
  return results;
}

// ─── AGGREGATE BY AD ──────────────────────────────────────
function aggregateByAd(rows, config) {
  var map = {};
  rows.forEach(function(r) {
    var key = r.platform + '||' + r.adName;
    if (!map[key]) {
      map[key] = { platform:r.platform, adName:r.adName, adType:r.adType,
                   campaign:r.campaign, adSet:r.adSet, product:r.product,
                   spend:0, revenue:0, clicks:0, linkClicks:0, lpv:0,
                   conversions:0, impressions:0, addToCart:0,
                   hookRateSum:0, holdRateSum:0, hookCount:0, holdCount:0, rows:[] };
    }
    map[key].spend       += r.spend;
    map[key].revenue     += r.revenue;
    map[key].clicks      += r.clicks;
    map[key].linkClicks  += r.linkClicks;
    map[key].lpv         += r.landingPageViews;
    map[key].conversions += r.conversions;
    map[key].impressions += r.impressions;
    map[key].addToCart   += r.addToCart;
    if (r.hookRate > 0) { map[key].hookRateSum += r.hookRate; map[key].hookCount++; }
    if (r.holdRate > 0) { map[key].holdRateSum += r.holdRate; map[key].holdCount++; }
    map[key].rows.push(r);
  });

  config = config || getConfig();
  var results = [];
  Object.keys(map).forEach(function(k) {
    var a = map[k];
    var roas     = a.spend > 0 ? a.revenue / a.spend : 0;
    var ctr      = a.impressions > 0 ? (a.clicks / a.impressions) * 100 : 0;
    var convRate = a.clicks > 0 ? (a.conversions / a.clicks) * 100 : 0;
    var cpc      = a.clicks > 0 ? a.spend / a.clicks : 0;
    var atcRate  = a.lpv > 0 ? (a.addToCart / a.lpv) * 100 : 0;
    var lpvRate  = a.linkClicks > 0 ? (a.lpv / a.linkClicks) * 100 : 0;
    var avgHook  = a.hookCount > 0 ? a.hookRateSum / a.hookCount : 0;
    var avgHold  = a.holdCount > 0 ? a.holdRateSum / a.holdCount : 0;
    var signal   = computeSignal(roas, ctr, cpc, convRate, a.spend, config);
    var sparkline = buildSparkline(a.rows);

    results.push({
      platform:    a.platform,
      adName:      a.adName,
      adType:      a.adType,
      campaign:    a.campaign,
      adSet:       a.adSet,
      product:     a.product,
      spend:       round2(a.spend),
      revenue:     round2(a.revenue),
      roas:        round2(roas),
      ctr:         round2(ctr),
      cpc:         round2(cpc),
      convRate:    round2(convRate),
      conversions: Math.round(a.conversions),
      linkClicks:  Math.round(a.linkClicks),
      lpv:         Math.round(a.lpv),
      addToCart:   Math.round(a.addToCart),
      atcRate:     round2(atcRate),
      lpvRate:     round2(lpvRate),
      hookRate:    round2(avgHook),
      holdRate:    round2(avgHold),
      signal:      signal,
      sparkline:   sparkline
    });
  });

  results.sort(function(a, b) { return b.roas - a.roas; });
  return results;
}

// ─── AGGREGATE BY PRODUCT ─────────────────────────────────
function aggregateByProduct(rows) {
  var map = {};
  rows.forEach(function(r) {
    var key = r.product || 'Unknown';
    if (!map[key]) map[key] = { product:key, spend:0, revenue:0, clicks:0, conversions:0, impressions:0 };
    map[key].spend       += r.spend;
    map[key].revenue     += r.revenue;
    map[key].clicks      += r.clicks;
    map[key].conversions += r.conversions;
    map[key].impressions += r.impressions;
  });

  var results = [];
  Object.keys(map).forEach(function(k) {
    var p = map[k];
    var roas    = p.spend > 0 ? p.revenue / p.spend : 0;
    var convRate = p.clicks > 0 ? (p.conversions / p.clicks) * 100 : 0;
    var cpc     = p.clicks > 0 ? p.spend / p.clicks : 0;
    results.push({
      product:     p.product,
      spend:       round2(p.spend),
      revenue:     round2(p.revenue),
      roas:        round2(roas),
      convRate:    round2(convRate),
      cpc:         round2(cpc),
      conversions: Math.round(p.conversions),
      clicks:      Math.round(p.clicks)
    });
  });

  results.sort(function(a, b) { return b.revenue - a.revenue; });
  return results;
}

// ─── AGGREGATE BY AD TYPE ─────────────────────────────────
function aggregateByAdType(rows) {
  var map = {};
  rows.forEach(function(r) {
    var key = r.adType || 'Unknown';
    if (!map[key]) map[key] = { adType:key, spend:0, revenue:0, clicks:0, lpv:0,
                                conversions:0, impressions:0, addToCart:0,
                                hookRateSum:0, holdRateSum:0, hookCount:0, holdCount:0 };
    map[key].spend       += r.spend;
    map[key].revenue     += r.revenue;
    map[key].clicks      += r.clicks;
    map[key].lpv         += r.landingPageViews;
    map[key].conversions += r.conversions;
    map[key].impressions += r.impressions;
    map[key].addToCart   += r.addToCart;
    if (r.hookRate > 0) { map[key].hookRateSum += r.hookRate; map[key].hookCount++; }
    if (r.holdRate > 0) { map[key].holdRateSum += r.holdRate; map[key].holdCount++; }
  });

  var results = [];
  Object.keys(map).forEach(function(k) {
    var t = map[k];
    var roas     = t.spend > 0 ? t.revenue / t.spend : 0;
    var ctr      = t.impressions > 0 ? (t.clicks / t.impressions) * 100 : 0;
    var convRate = t.clicks > 0 ? (t.conversions / t.clicks) * 100 : 0;
    var atcRate  = t.lpv > 0 ? (t.addToCart / t.lpv) * 100 : 0;
    var avgHook  = t.hookCount > 0 ? t.hookRateSum / t.hookCount : 0;
    var avgHold  = t.holdCount > 0 ? t.holdRateSum / t.holdCount : 0;
    results.push({
      adType: t.adType, spend: round2(t.spend), revenue: round2(t.revenue),
      roas: round2(roas), ctr: round2(ctr), convRate: round2(convRate),
      conversions: Math.round(t.conversions), addToCart: Math.round(t.addToCart),
      atcRate: round2(atcRate), hookRate: round2(avgHook), holdRate: round2(avgHold)
    });
  });
  results.sort(function(a, b) { return b.roas - a.roas; });
  return results;
}

// ─── GET CONFIG ───────────────────────────────────────────
function getConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET);
  var defaults = { scaleRoas:4.0, stopRoas:1.5, scaleCtr:2.5, stopCtr:0.5,
                   scaleCpc:30, stopCpc:100, minSpend:200, fatigueFreq:4.0 };
  if (!sheet || sheet.getLastRow() <= 1) return defaults;

  var data = sheet.getDataRange().getValues();
  var config = Object.assign({}, defaults);
  for (var i = 1; i < data.length; i++) {
    var metric = data[i][0];
    var scale  = cleanParseFloat(data[i][2]) || 0;
    var stop   = cleanParseFloat(data[i][3]) || 0;
    if (metric === 'ROAS')     { config.scaleRoas = scale; config.stopRoas = stop; }
    if (metric === 'CTR')      { config.scaleCtr  = scale; config.stopCtr  = stop; }
    if (metric === 'CPC')      { config.scaleCpc  = scale; config.stopCpc  = stop; }
    if (metric === 'Frequency') config.fatigueFreq = stop;
    if (metric === 'Min Spend') config.minSpend   = scale;
  }
  return config;
}

// ─── COMPUTE SIGNAL ───────────────────────────────────────
function computeSignal(roas, ctr, cpc, convRate, spend, config) {
  // STOP conditions
  if (roas > 0 && roas <= config.stopRoas) return 'STOP';
  if (spend > config.minSpend && convRate === 0) return 'STOP';
  if (ctr > 0 && ctr <= config.stopCtr) return 'STOP';

  // SCALE conditions
  if (roas >= config.scaleRoas && ctr >= config.scaleCtr) return 'SCALE';
  if (roas >= config.scaleRoas && convRate >= 2.0) return 'SCALE';

  // MONITOR
  return 'MONITOR';
}

// ─── TREND DETECTION ──────────────────────────────────────
function computeTrend(rows) {
  if (rows.length < 2) return 'flat';
  rows.sort(function(a, b) { return a.date - b.date; });
  var half = Math.floor(rows.length / 2);
  var firstHalf = rows.slice(0, half);
  var secondHalf = rows.slice(half);
  var avg1 = firstHalf.reduce(function(s, r) { return s + r.roas; }, 0) / firstHalf.length;
  var avg2 = secondHalf.reduce(function(s, r) { return s + r.roas; }, 0) / secondHalf.length;
  if (avg2 > avg1 * 1.1) return 'up';
  if (avg2 < avg1 * 0.9) return 'down';
  return 'flat';
}

// ─── BUILD SPARKLINE DATA ─────────────────────────────────
function buildSparkline(rows) {
  var byDate = {};
  rows.forEach(function(r) {
    if (!byDate[r.dateStr]) byDate[r.dateStr] = { roas: 0, count: 0 };
    byDate[r.dateStr].roas  += r.roas;
    byDate[r.dateStr].count += 1;
  });
  var dates = Object.keys(byDate).sort();
  return dates.slice(-7).map(function(d) {
    return round2(byDate[d].roas / byDate[d].count);
  });
}

// ─── GET UNIQUE VALUES FOR DROPDOWNS ─────────────────────
function getUniqueValues(ss, sheetNames, colIndex) {
  var seen = {};
  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;
    var col = sheet.getRange(2, colIndex + 1, sheet.getLastRow() - 1, 1).getValues();
    col.forEach(function(row) {
      var v = (row[0] || '').toString().trim();
      if (v) seen[v] = 1;
    });
  });
  return Object.keys(seen).sort();
}

// ─── GENERATE INTELLIGENCE ────────────────────────────────
function generateIntelligence(campaigns, ads, products, adTypes, kpis, config, allRows) {
  var insights = [];
  config = config || getConfig();

  // Best performing ad
  if (ads.length > 0) {
    var best = ads[0];
    insights.push({
      type: 'success',
      icon: '🏆',
      title: 'Top Performing Ad',
      message: '"' + best.adName + '" (' + best.adType + ') is your #1 ad with ' +
               best.roas + 'x ROAS and ' + best.convRate + '% conv rate on ' + best.platform + '.',
      action: 'Scale budget on this ad immediately.'
    });
  }

  // Ads to stop
  var stopAds = ads.filter(function(a) { return a.signal === 'STOP'; });
  if (stopAds.length > 0) {
    insights.push({
      type: 'danger',
      icon: '🛑',
      title: stopAds.length + ' Ad(s) Burning Budget',
      message: stopAds.slice(0, 3).map(function(a) { return '"' + a.adName + '"'; }).join(', ') +
               (stopAds.length > 3 ? ' and ' + (stopAds.length - 3) + ' more' : '') +
               ' are underperforming with sub-' + config.stopRoas + 'x ROAS.',
      action: 'Pause or refresh creative immediately.'
    });
  }

  // Campaigns to scale
  var scaleCampaigns = campaigns.filter(function(c) { return c.signal === 'SCALE'; });
  if (scaleCampaigns.length > 0) {
    var sc = scaleCampaigns[0];
    insights.push({
      type: 'scale',
      icon: '🚀',
      title: 'Campaign Ready to Scale',
      message: '"' + sc.campaign + '" (' + sc.platform + ') is hitting ' + sc.roas + 'x ROAS with a ' +
               sc.trend + ' trend. Allocate more budget.',
      action: 'Increase daily budget by 20-30%.'
    });
  }

  // ROAS health check
  if (kpis.roas < config.stopRoas && kpis.roas > 0) {
    insights.push({
      type: 'warning', icon: '⚠️', title: 'Overall ROAS Below Target',
      message: 'Blended ROAS is ' + kpis.roas + 'x against a ' + config.stopRoas + 'x minimum. Review all active campaigns.',
      action: 'Audit top spend campaigns immediately.'
    });
  }

  // Hook Rate signal
  if (kpis.hookRate > 0) {
    var hookGood = kpis.hookRate >= 30;
    insights.push({
      type: hookGood ? 'success' : 'warning', icon: '🪝',
      title: hookGood ? 'Strong Hook Rate' : 'Weak Hook Rate — Creatives Not Grabbing Attention',
      message: 'Average Hook Rate is ' + kpis.hookRate + '%. ' +
               (hookGood ? 'Your creatives are grabbing attention in the first 3 seconds.' :
                'Less than ' + kpis.hookRate + '% of viewers watch past 3 seconds — your opening needs work.'),
      action: hookGood ? 'A/B test hooks to push past 40%.' : 'Rewrite opening 3 seconds — use pattern interrupts, bold text, or a strong hook statement.'
    });
  }

  // Hold Rate signal
  if (kpis.holdRate > 0) {
    var holdGood = kpis.holdRate >= 20;
    insights.push({
      type: holdGood ? 'success' : 'warning', icon: '⏱️',
      title: holdGood ? 'Good Audience Retention' : 'Low Hold Rate — Viewers Dropping Off',
      message: 'Average Hold Rate (watched 25%+) is ' + kpis.holdRate + '%. ' +
               (holdGood ? 'Audience is staying engaged through your content.' :
                'Most viewers drop off before 25% — video is not compelling enough.'),
      action: holdGood ? 'Test longer-form content for retargeting.' : 'Shorten videos to 15-20s. Lead with product benefit immediately.'
    });
  }

  // ATC Rate signal
  if (kpis.atcRate > 0) {
    var atcGood = kpis.atcRate >= 5;
    insights.push({
      type: atcGood ? 'success' : 'warning', icon: '🛒',
      title: atcGood ? 'Healthy Add-to-Cart Rate' : 'Low ATC Rate — Landing Page Underperforming',
      message: 'ATC Rate is ' + kpis.atcRate + '% of Landing Page Views. Total ATCs: ' + fmtIntel(kpis.addToCart) + '.',
      action: atcGood ? 'Optimize checkout flow to convert more ATCs to purchases.' : 'Review landing page UX, product images, and price anchoring.'
    });
  }

  // LPV Drop-off
  if (kpis.linkClicks > 0 && kpis.lpv > 0) {
    var lpvRate = round2((kpis.lpv / kpis.linkClicks) * 100);
    var lpvGood = lpvRate >= 70;
    insights.push({
      type: lpvGood ? 'info' : 'warning', icon: '🔗',
      title: lpvGood ? 'Good Link-to-LPV Rate' : 'High Link Drop-off Before Landing Page',
      message: lpvRate + '% of link clicks actually view the landing page (' + fmtIntel(kpis.linkClicks) + ' clicks → ' + fmtIntel(kpis.lpv) + ' LPVs).',
      action: lpvGood ? 'Maintain fast page load speeds.' : 'Improve page speed (target <2s load). Check for broken links or redirect issues.'
    });
  }

  // Best ad type
  if (adTypes.length > 0) {
    var bestType = adTypes[0];
    insights.push({
      type: 'info', icon: '🎬', title: 'Best Creative Format',
      message: bestType.adType + ' ads deliver ' + bestType.roas + 'x ROAS' +
               (bestType.hookRate > 0 ? ', ' + bestType.hookRate + '% Hook Rate' : '') +
               (bestType.atcRate > 0 ? ', ' + bestType.atcRate + '% ATC Rate' : '') + '.',
      action: 'Produce more ' + bestType.adType + ' content.'
    });
  }

  // Best product
  if (products.length > 0) {
    var bp = products[0];
    insights.push({
      type: 'info', icon: '📦', title: 'Top Revenue Product',
      message: '"' + bp.product + '" generated ₹' + formatNum(bp.revenue) + ' in revenue with ' + bp.roas + 'x ROAS.',
      action: 'Prioritize this product in creative briefs.'
    });
  }

  // Budget efficiency
  if (kpis.spend > 0) {
    var efficiency = ((kpis.revenue - kpis.spend) / kpis.spend * 100).toFixed(1);
    insights.push({
      type: efficiency >= 0 ? 'success' : 'warning', icon: '💰', title: 'Marketing ROI',
      message: 'For every ₹1 spent, you generated ₹' + round2(kpis.roas) + ' in revenue. Net efficiency: ' + efficiency + '%.',
      action: efficiency >= 100 ? 'Strong returns — consider scaling total budget.' : 'Reduce wastage by pausing STOP-signal ads.'
    });
  }

  // Product-specific adset AI signals (3d & 7d window comparison)
  if (allRows && allRows.length > 0) {
    var maxTs = 0;
    allRows.forEach(function(r) {
      var t = r.date.getTime();
      if (t > maxTs) maxTs = t;
    });
    var maxDate = maxTs > 0 ? new Date(maxTs) : new Date();
    maxDate.setHours(23, 59, 59, 999);
    var ts3d = new Date(maxDate); ts3d.setDate(maxDate.getDate() - 3); ts3d.setHours(0,0,0,0);
    var ts7d = new Date(maxDate); ts7d.setDate(maxDate.getDate() - 7); ts7d.setHours(0,0,0,0);

    var prodAdSets = {};
    allRows.forEach(function(r) {
      if (!r.product || !r.adSet) return;
      var key = r.product + '||' + r.adSet;
      if (!prodAdSets[key]) {
        prodAdSets[key] = {
          product: r.product,
          adSet: r.adSet,
          platform: r.platform,
          spend_3d: 0, rev_3d: 0,
          spend_7d: 0, rev_7d: 0
        };
      }
      var pa = prodAdSets[key];
      var rTime = r.date.getTime();
      if (rTime >= ts3d.getTime()) {
        pa.spend_3d += r.spend; pa.rev_3d += r.revenue;
      }
      if (rTime >= ts7d.getTime()) {
        pa.spend_7d += r.spend; pa.rev_7d += r.revenue;
      }
    });

    var minTargetRoas = config.stopRoas || 1.5;
    var scaleTargetRoas = config.scaleRoas || 4.0;

    var prodSignals = {};
    Object.keys(prodAdSets).forEach(function(k) {
      var pa = prodAdSets[k];
      var roas_3d = pa.spend_3d > 0 ? round2(pa.rev_3d / pa.spend_3d) : 0;
      var roas_7d = pa.spend_7d > 0 ? round2(pa.rev_7d / pa.spend_7d) : 0;
      
      if (!prodSignals[pa.product]) {
        prodSignals[pa.product] = { scale: [], pause: [], stop: [] };
      }
      var pSig = prodSignals[pa.product];
      
      if (pa.spend_3d > 50) { // Only warn if there is meaningful spend
        if (roas_3d >= scaleTargetRoas) {
          pSig.scale.push(pa.adSet + ' (' + roas_3d.toFixed(2) + 'x ROAS)');
        } else if (roas_3d < minTargetRoas && roas_7d < minTargetRoas) {
          pSig.stop.push(pa.adSet + ' (3d: ' + roas_3d.toFixed(2) + 'x, 7d: ' + roas_7d.toFixed(2) + 'x ROAS)');
        } else if (roas_3d < minTargetRoas) {
          pSig.pause.push(pa.adSet + ' (' + roas_3d.toFixed(2) + 'x ROAS)');
        }
      }
    });

    Object.keys(prodSignals).forEach(function(prod) {
      var s = prodSignals[prod];
      if (s.scale.length > 0) {
        insights.push({
          type: 'success', icon: '🚀',
          title: 'Scale Adsets for ' + prod,
          message: 'The following adsets have high performance and should be scaled: ' + s.scale.join(', '),
          action: 'Increase budget by 20-30%'
        });
      }
      if (s.stop.length > 0) {
        insights.push({
          type: 'danger', icon: '🛑',
          title: 'Stop Adsets for ' + prod,
          message: 'The following adsets have consistently underperformed over both 3-day and 7-day windows and should be stopped immediately: ' + s.stop.join(', '),
          action: 'Turn off these adsets'
        });
      }
      if (s.pause.length > 0) {
        insights.push({
          type: 'warning', icon: '⏸️',
          title: 'Pause / Review Adsets for ' + prod,
          message: 'The following adsets show a short-term dip below target ROAS and should be paused or optimized: ' + s.pause.join(', '),
          action: 'Pause and check hook/hold rates'
        });
      }
    });
  }

  return insights;
}

// ─── REFRESH INTELLIGENCE ─────────────────────────────────
function refreshIntelligence() {
  SpreadsheetApp.getUi().alert('✅ Intelligence refreshed! Open the Dashboard to see updated recommendations.');
}

// ─── HELPERS ──────────────────────────────────────────────
function round2(n) { return Math.round(n * 100) / 100; }

function cleanParseFloat(val) {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') return val;
  var valStr = val.toString().trim();
  if (!valStr) return 0;
  var clean = valStr.replace(/[^\d.eE+-]/g, '');
  var num = parseFloat(clean);
  return isNaN(num) ? 0 : num;
}

function formatNum(n) {
  if (n >= 100000) return (n / 100000).toFixed(1) + 'L';
  if (n >= 1000)   return (n / 1000).toFixed(1) + 'K';
  return Math.round(n).toString();
}

// Compact number for intelligence text
function fmtIntel(n) {
  if (n >= 1000000) return (n / 1000000).toFixed(1) + 'M';
  if (n >= 1000)    return (n / 1000).toFixed(1) + 'K';
  return Math.round(n).toString();
}

function getDefaultStartDate() {
  var d = new Date();
  d.setDate(d.getDate() - 30);
  return d;
}

// ─── COHORT & LTV CALCULATION ─────────────────────────────
function buildCohortData(ss) {
  var sheet = ss.getSheetByName(SALES_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var cohorts = {}; // key: cohortMonth -> { size: 0, customers: {}, revenue: [0,0,0,0], activeCount: [0,0,0,0] }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var custId = row[1];
    var orderDateRaw = row[2];
    var revenue = cleanParseFloat(row[3]) || 0;
    var cohortMonthRaw = row[6];
    if (!cohortMonthRaw) continue;

    var cohortMonth = (cohortMonthRaw instanceof Date) ? formatMonthJS(cohortMonthRaw) : String(cohortMonthRaw).trim();
    if (!cohortMonth || cohortMonth.length < 7) continue;

    if (!cohorts[cohortMonth]) {
      cohorts[cohortMonth] = {
        cohort: cohortMonth,
        size: 0,
        customers: {},
        revenue: [0, 0, 0, 0],
        activeCount: [0, 0, 0, 0],
        activeCustMap: [{}, {}, {}, {}]
      };
    }

    var c = cohorts[cohortMonth];
    if (!c.customers[custId]) {
      c.customers[custId] = true;
      c.size++;
    }

    var orderDate = (orderDateRaw instanceof Date) ? orderDateRaw : new Date(orderDateRaw);
    var orderTs = orderDate.getTime();
    if (isNaN(orderTs)) continue;

    var orderMonthStr = formatMonthJS(orderDate);
    var diff = getDiffMonths(cohortMonth, orderMonthStr);

    if (diff >= 0 && diff <= 3) {
      c.revenue[diff] += revenue;
      if (!c.activeCustMap[diff][custId]) {
        c.activeCustMap[diff][custId] = true;
        c.activeCount[diff]++;
      }
    }
  }

  // Format results
  var results = [];
  var cohortKeys = Object.keys(cohorts).sort();
  cohortKeys.forEach(function(key) {
    var c = cohorts[key];
    // Cumulative LTV
    var ltv = [0, 0, 0, 0];
    var cumulativeRevenue = 0;
    for (var m = 0; m < 4; m++) {
      cumulativeRevenue += c.revenue[m];
      ltv[m] = c.size > 0 ? round2(cumulativeRevenue / c.size) : 0;
    }
    
    // Retention rates
    var retention = [0, 0, 0, 0];
    for (var m = 0; m < 4; m++) {
      retention[m] = c.size > 0 ? round2((c.activeCount[m] / c.size) * 100) : 0;
    }

    results.push({
      cohort: key,
      size: c.size,
      revenue: c.revenue.map(round2),
      ltv: ltv,
      retention: retention
    });
  });

  return results;
}

function getDiffMonths(startStr, endStr) {
  if (!startStr || !endStr) return -1;
  var startParts = startStr.split('-');
  var endParts = endStr.split('-');
  if (startParts.length < 2 || endParts.length < 2) return -1;
  var startYear = parseInt(startParts[0], 10);
  var startMonth = parseInt(startParts[1], 10);
  var endYear = parseInt(endParts[0], 10);
  var endMonth = parseInt(endParts[1], 10);
  if (isNaN(startYear) || isNaN(startMonth) || isNaN(endYear) || isNaN(endMonth)) return -1;
  return (endYear - startYear) * 12 + (endMonth - startMonth);
}

// ─── BUDGET REALLOCATION OPTIMIZER ────────────────────────
function buildBudgetRecommendations(campaignData) {
  var actions = [];
  var totalWastedSpend = 0;
  var scaleCandidates = [];

  campaignData.forEach(function(c) {
    if (c.signal === 'STOP') {
      totalWastedSpend += c.spend;
      actions.push({
        type: 'STOP',
        campaign: c.campaign,
        platform: c.platform,
        currentSpend: c.spend,
        currentRoas: c.roas,
        recommendation: 'Pause Campaign',
        message: 'Pause "' + c.campaign + '" (' + c.platform + ') to cut wastage of ₹' + formatNum(c.spend) + '. Current ROAS is ' + c.roas + 'x against 1.5x min target.',
        actionItem: 'Deallocate ₹' + formatNum(c.spend) + ' immediately.'
      });
    } else if (c.signal === 'SCALE') {
      scaleCandidates.push(c);
    }
  });

  scaleCandidates.forEach(function(c) {
    var suggestedIncrease = Math.round(c.spend * 0.25);
    if (suggestedIncrease < 100) suggestedIncrease = 500;
    var projectedRevenue = Math.round(suggestedIncrease * c.roas);
    
    actions.push({
      type: 'SCALE',
      campaign: c.campaign,
      platform: c.platform,
      currentSpend: c.spend,
      currentRoas: c.roas,
      recommendation: 'Increase Budget by 25%',
      message: 'Increase budget of "' + c.campaign + '" (' + c.platform + ') by ₹' + formatNum(suggestedIncrease) + '. Current ROAS is ' + c.roas + 'x and trending ' + c.trend + '.',
      actionItem: 'Allocate ₹' + formatNum(suggestedIncrease) + ' (Est. Gain: +₹' + formatNum(projectedRevenue) + ' revenue).'
    });
  });

  return {
    actions: actions,
    totalWastedSpend: round2(totalWastedSpend),
    potentialOptimizedRevenue: round2(scaleCandidates.reduce(function(sum, c) {
      return sum + (c.spend * 0.25 * c.roas);
    }, 0))
  };
}

// ─── AD CREATIVE FATIGUE DETECTION ────────────────────────
function buildFatigueData(allRows, adData) {
  var adDailyMap = {};
  allRows.forEach(function(r) {
    var key = r.platform + '||' + r.adName;
    if (!adDailyMap[key]) adDailyMap[key] = [];
    adDailyMap[key].push(r);
  });

  var results = [];
  adData.forEach(function(ad) {
    var key = ad.platform + '||' + ad.adName;
    var rows = adDailyMap[key] || [];
    if (rows.length < 12) {
      results.push({
        adName: ad.adName,
        adType: ad.adType,
        platform: ad.platform,
        frequency: ad.frequency,
        currentCtr: ad.ctr,
        recentCtr: ad.ctr,
        baselineCtr: ad.ctr,
        ctrDrop: 0,
        status: 'HEALTHY',
        recommendation: 'Monitor baseline performance'
      });
      return;
    }

    rows.sort(function(a, b) { return a.date - b.date; });

    var recentRows = rows.slice(-7);
    var baselineRows = rows.slice(-21, -7);

    var recentCtr = calcAvgCtr(recentRows);
    var baselineCtr = calcAvgCtr(baselineRows);

    var ctrDrop = baselineCtr > 0 ? (baselineCtr - recentCtr) / baselineCtr : 0;
    
    var status = 'HEALTHY';
    var recommendation = 'Creative performance is stable. Maintain active budget.';

    if (ctrDrop >= 0.20 && ad.frequency >= 2.5) {
      status = 'HIGH FATIGUE';
      recommendation = 'Critical CTR drop-off. Refresh visual elements or replace creative immediately.';
    } else if (ctrDrop >= 0.10 || (ctrDrop >= 0.05 && ad.frequency >= 3.0)) {
      status = 'MODERATE FATIGUE';
      recommendation = 'CTR is softening. Prepare alternative creative variants to prevent performance drop.';
    }

    results.push({
      adName: ad.adName,
      adType: ad.adType,
      platform: ad.platform,
      frequency: round2(ad.frequency),
      currentCtr: ad.ctr,
      recentCtr: round2(recentCtr),
      baselineCtr: round2(baselineCtr),
      ctrDrop: round2(ctrDrop * 100),
      status: status,
      recommendation: recommendation
    });
  });

  return results;
}

function calcAvgCtr(rows) {
  var totalImpressions = 0, totalClicks = 0;
  rows.forEach(function(r) {
    totalImpressions += r.impressions;
    totalClicks += r.clicks;
  });
  return totalImpressions > 0 ? (totalClicks / totalImpressions) * 100 : 0;
}

// ─── HIGH PERFORMANCE JS DATE FORMATTERS ──────────────────
function formatDateJS(d) {
  var year = d.getFullYear();
  var month = d.getMonth() + 1;
  var day = d.getDate();
  return year + '-' + (month < 10 ? '0' + month : month) + '-' + (day < 10 ? '0' + day : day);
}

function formatMonthJS(d) {
  var year = d.getFullYear();
  var month = d.getMonth() + 1;
  return year + '-' + (month < 10 ? '0' + month : month);
}

// ─── FORCE RE-TAG ALL ROWS (overwrites existing values) ────────────────
function forceReTagAllRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var sheetsToProcess = [META_SHEET, GOOGLE_SHEET];
  var totalTagged = 0;
  
  sheetsToProcess.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;
    
    var allValues = sheet.getDataRange().getValues();
    var headers = allValues[0];
    var colMap = {};
    headers.forEach(function(h, i) { colMap[h.toString().trim()] = i; });
    
    var campaignIdx = colMap['Campaign'] !== undefined ? colMap['Campaign'] : colMap['Campaign name'];
    var adSetIdx    = colMap['Ad Set']    !== undefined ? colMap['Ad Set']    : (colMap['Ad Group'] !== undefined ? colMap['Ad Group'] : colMap['Ad set name']);
    var adNameIdx   = colMap['Ad Name']   !== undefined ? colMap['Ad Name']   : colMap['Ad name'];
    if (campaignIdx === undefined || adNameIdx === undefined) return;
    
    // Auto-add missing tag columns
    var requiredCols = ['Product','Ad Type','Type of Ad','Influencer Name','Narrative','Language'];
    var lastCol = headers.length;
    requiredCols.forEach(function(colName) {
      if (colMap[colName] === undefined) {
        sheet.getRange(1, lastCol + 1).setValue(colName);
        colMap[colName] = lastCol;
        headers.push(colName);
        lastCol++;
      }
    });
    // Re-read after adding headers
    allValues = sheet.getDataRange().getValues();
    headers = allValues[0];
    colMap = {};
    headers.forEach(function(h, i) { colMap[h.toString().trim()] = i; });
    campaignIdx = colMap['Campaign'] !== undefined ? colMap['Campaign'] : colMap['Campaign name'];
    adSetIdx    = colMap['Ad Set']    !== undefined ? colMap['Ad Set']    : (colMap['Ad Group'] !== undefined ? colMap['Ad Group'] : colMap['Ad set name']);
    adNameIdx   = colMap['Ad Name']   !== undefined ? colMap['Ad Name']   : colMap['Ad name'];
    
    var productIdx    = colMap['Product']        !== undefined ? colMap['Product']        : colMap['Product Name'];
    var adTypeIdx     = colMap['Ad Type'];
    var typeOfAdIdx   = colMap['Type of Ad']     !== undefined ? colMap['Type of Ad']     : colMap['Type of Ad (Static/Reel)'];
    var influencerIdx = colMap['Influencer Name'] !== undefined ? colMap['Influencer Name'] : colMap['Name of the Influencer'];
    var narrativeIdx  = colMap['Narrative'];
    var languageIdx   = colMap['Language']       !== undefined ? colMap['Language']       : colMap['Language of the Ad'];
    
    var indices = [productIdx, adTypeIdx, typeOfAdIdx, influencerIdx, narrativeIdx, languageIdx].filter(function(i) { return i !== undefined; });
    if (indices.length === 0) return;
    
    var minCol = Math.min.apply(null, indices);
    var maxCol = Math.max.apply(null, indices);
    var numRows = sheet.getLastRow() - 1;
    var metaRange = sheet.getRange(2, minCol + 1, numRows, maxCol - minCol + 1);
    var metaValues = metaRange.getValues();
    var sheetTagged = 0;
    
    for (var i = 1; i < allValues.length; i++) {
      var row = allValues[i];
      var campaign = String(row[campaignIdx] || '');
      var adSet    = adSetIdx !== undefined ? String(row[adSetIdx] || '') : '';
      var adName   = String(row[adNameIdx] || '');
      if (!adName.trim()) continue;
      
      var tags = autoTagMetadata(campaign, adSet, adName);
      
      // FORCE overwrite all tag columns
      if (productIdx    !== undefined && tags.product)        metaValues[i-1][productIdx    - minCol] = tags.product;
      if (adTypeIdx     !== undefined && tags.adType)         metaValues[i-1][adTypeIdx     - minCol] = tags.adType;
      if (typeOfAdIdx   !== undefined && tags.typeOfAd)       metaValues[i-1][typeOfAdIdx   - minCol] = tags.typeOfAd;
      if (influencerIdx !== undefined && tags.influencerName) metaValues[i-1][influencerIdx - minCol] = tags.influencerName;
      if (narrativeIdx  !== undefined)                        metaValues[i-1][narrativeIdx  - minCol] = tags.narrative; // always write (even blank)
      if (languageIdx   !== undefined && tags.language)       metaValues[i-1][languageIdx   - minCol] = tags.language;
      sheetTagged++;
    }
    
    metaRange.setValues(metaValues);
    totalTagged += sheetTagged;
    logDebug('Force re-tagged ' + sheetTagged + ' rows in ' + sheetName);
  });
  
  ui.alert('✅ Force Re-Tag Complete!\n\nOverwrote tags on ' + totalTagged + ' rows in META_DATA.');
}

// ─── DYNAMIC AUTO-TAGGING ENGINE ──────────────────────────
function autoTagBlankRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var sheetsToProcess = [
    { name: META_SHEET },
    { name: GOOGLE_SHEET }
  ];
  
  var totalTagged = 0;
  
  sheetsToProcess.forEach(function(sheetConf) {
    var sheet = ss.getSheetByName(sheetConf.name);
    if (!sheet || sheet.getLastRow() <= 1) return;
    
    var allRange = sheet.getDataRange();
    var allValues = allRange.getValues();
    var headers = allValues[0];
    
    // Map column indices
    var colMap = {};
    headers.forEach(function(h, i) { colMap[h.toString().trim()] = i; });
    
    var campaignIdx = colMap['Campaign'] !== undefined ? colMap['Campaign'] : colMap['Campaign name'];
    var adSetIdx = colMap['Ad Set'] !== undefined ? colMap['Ad Set'] : (colMap['Ad Group'] !== undefined ? colMap['Ad Group'] : colMap['Ad set name']);
    var adNameIdx = colMap['Ad Name'] !== undefined ? colMap['Ad Name'] : colMap['Ad name'];
    
    if (campaignIdx === undefined || adNameIdx === undefined) return;
    
    // ---- Auto-add missing tag columns ----
    var requiredCols = ['Product', 'Ad Type', 'Type of Ad', 'Influencer Name', 'Narrative', 'Language'];
    var lastCol = headers.length; // 0-indexed last column count
    requiredCols.forEach(function(colName) {
      if (colMap[colName] === undefined) {
        sheet.getRange(1, lastCol + 1).setValue(colName);
        colMap[colName] = lastCol;
        headers.push(colName);
        lastCol++;
      }
    });
    // Re-read data after adding headers
    allRange = sheet.getDataRange();
    allValues = allRange.getValues();
    headers = allValues[0];
    colMap = {};
    headers.forEach(function(h, i) { colMap[h.toString().trim()] = i; });
    campaignIdx = colMap['Campaign'] !== undefined ? colMap['Campaign'] : colMap['Campaign name'];
    adSetIdx = colMap['Ad Set'] !== undefined ? colMap['Ad Set'] : (colMap['Ad Group'] !== undefined ? colMap['Ad Group'] : colMap['Ad set name']);
    adNameIdx = colMap['Ad Name'] !== undefined ? colMap['Ad Name'] : colMap['Ad name'];
    // ---- End auto-add ----
    
    var productIdx = colMap['Product'] !== undefined ? colMap['Product'] : colMap['Product Name'];
    var adTypeIdx = colMap['Ad Type'];
    var typeOfAdIdx = colMap['Type of Ad'] !== undefined ? colMap['Type of Ad'] : colMap['Type of Ad (Static/Reel)'];
    var influencerIdx = colMap['Influencer Name'] !== undefined ? colMap['Influencer Name'] : colMap['Name of the Influencer'];
    var narrativeIdx = colMap['Narrative'];
    var languageIdx = colMap['Language'] !== undefined ? colMap['Language'] : colMap['Language of the Ad'];
    
    // Determine contiguous metadata columns bounds to prevent formula overwrites
    var indices = [productIdx, adTypeIdx, typeOfAdIdx, influencerIdx, narrativeIdx, languageIdx].filter(function(idx) {
      return idx !== undefined;
    });
    if (indices.length === 0) return;
    
    var minCol = Math.min.apply(null, indices);
    var maxCol = Math.max.apply(null, indices);
    var startCol = minCol + 1;
    var width = maxCol - minCol + 1;
    var numRows = sheet.getLastRow() - 1;
    
    var metaRange = sheet.getRange(2, startCol, numRows, width);
    var metaValues = metaRange.getValues();
    
    var sheetTagged = 0;
    
    for (var i = 1; i < allValues.length; i++) {
      var row = allValues[i];
      var campaign = String(row[campaignIdx] || '');
      var adSet = adSetIdx !== undefined ? String(row[adSetIdx] || '') : '';
      var adName = String(row[adNameIdx] || '');
      
      var currentProduct = productIdx !== undefined ? String(row[productIdx] || '').trim() : '';
      var currentAdType = adTypeIdx !== undefined ? String(row[adTypeIdx] || '').trim() : '';
      var currentTypeOfAd = typeOfAdIdx !== undefined ? String(row[typeOfAdIdx] || '').trim() : '';
      var currentInfluencer = influencerIdx !== undefined ? String(row[influencerIdx] || '').trim() : '';
      var currentNarrative = narrativeIdx !== undefined ? String(row[narrativeIdx] || '').trim() : '';
      var currentLanguage = languageIdx !== undefined ? String(row[languageIdx] || '').trim() : '';
      
      // Only auto-tag if at least one of these columns is empty/blank
      if (!currentProduct || !currentAdType || !currentTypeOfAd || !currentInfluencer || !currentNarrative || !currentLanguage) {
        var tags = autoTagMetadata(campaign, adSet, adName);
        var cellsUpdated = false;
        
        if (!currentProduct && tags.product && productIdx !== undefined) {
          metaValues[i - 1][productIdx - minCol] = tags.product;
          cellsUpdated = true;
        }
        if (!currentAdType && tags.adType && adTypeIdx !== undefined) {
          metaValues[i - 1][adTypeIdx - minCol] = tags.adType;
          cellsUpdated = true;
        }
        if (!currentTypeOfAd && tags.typeOfAd && typeOfAdIdx !== undefined) {
          metaValues[i - 1][typeOfAdIdx - minCol] = tags.typeOfAd;
          cellsUpdated = true;
        }
        if (!currentInfluencer && tags.influencerName && influencerIdx !== undefined) {
          metaValues[i - 1][influencerIdx - minCol] = tags.influencerName;
          cellsUpdated = true;
        }
        if (!currentNarrative && tags.narrative && narrativeIdx !== undefined) {
          metaValues[i - 1][narrativeIdx - minCol] = tags.narrative;
          cellsUpdated = true;
        }
        if (!currentLanguage && tags.language && languageIdx !== undefined) {
          metaValues[i - 1][languageIdx - minCol] = tags.language;
          cellsUpdated = true;
        }
        
        if (cellsUpdated) {
          sheetTagged++;
        }
      }
    }
    
    if (sheetTagged > 0) {
      metaRange.setValues(metaValues);
      totalTagged += sheetTagged;
    }
    logDebug("Auto-tagged " + sheetTagged + " rows in " + sheetConf.name);
  });
  
  ui.alert("🤖 Auto-Tagging Complete!\n\nSuccessfully analyzed and updated " + totalTagged + " blank rows in META_DATA and GOOGLE_DATA.");
}

function capitalizeWord(w) {
  if (!w) return '';
  if (w.indexOf('/') !== -1) {
    return w.split('/').map(capitalizeWord).join('/');
  }
  if (w.match(/^\d+%/)) {
    return w.toUpperCase();
  }
  if (w.toLowerCase() === 'aha' || w.toLowerCase() === 'bha' || w.toLowerCase() === 'ugc') {
    return w.toUpperCase();
  }
  return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
}

function isLanguageToken(tok) {
  var tokClean = tok.toLowerCase().trim();
  // Explicit full-match variants first
  var fullMatch = ['hindiandenglish','englishandhindi','hindiinenglishscript','hindiinhindiscript',
                   'hinglish','hindiinromanscript','hindiinenglish'];
  if (fullMatch.indexOf(tokClean) !== -1) return true;
  // Prefix match for base language names
  var langs = ['english','hindi','marathi','tamil','telugu','kannada','malayalam','punjabi'];
  for (var i = 0; i < langs.length; i++) {
    if (tokClean.indexOf(langs[i]) === 0) return true;
  }
  return false;
}

function parseAdNameMetadata(adName) {
  return autoTagMetadata('', '', adName);
}

function autoTagMetadata(campaign, adSet, adName) {
  var adNameLower = (adName || '').toLowerCase().trim();
  var campaignLower = (campaign || '').toLowerCase().trim();
  var adSetLower = (adSet || '').toLowerCase().trim();
  var combined = campaignLower + " | " + adSetLower + " | " + adNameLower;
  
  // 1. PRODUCT SKU MATCHING
  var product = 'Other Product';
  
  if (campaignLower.indexOf('body category') !== -1 || adSetLower.indexOf('body category') !== -1 || campaignLower.indexOf('body_category') !== -1 || adSetLower.indexOf('body_category') !== -1) {
    if (adSetLower.indexOf('category page') !== -1 || adSetLower.indexOf('category_page') !== -1) {
      product = 'Body Category';
    } else if (adNameLower.indexOf('underarm_pigmentation_pack') !== -1 || adNameLower.indexOf('pigmentation_pack') !== -1 || adNameLower.indexOf('upk') !== -1) {
      product = 'Advanced Underarm Pigmentation Kit';
    } else if (adNameLower.indexOf('6% underarm') !== -1 || adNameLower.indexOf('6% roll') !== -1 || adNameLower.indexOf('6% uaro') !== -1) {
      product = '6% AHA BHA Underarm Roll On';
    } else if (adNameLower.indexOf('5%_aha_bha_bodywash') !== -1 || adNameLower.indexOf('5% aha bha bodywash') !== -1 || combined.indexOf('5% bodywash') !== -1 || combined.indexOf('exfoliating body wash') !== -1) {
      product = '5% AHA BHA Exfoliating Body Wash';
    } else if (adSetLower.indexOf('pigmentation') !== -1 || adSetLower.indexOf('upk') !== -1) {
      product = 'Advanced Underarm Pigmentation Kit';
    } else if (adSetLower.indexOf('6% underarm') !== -1 || adSetLower.indexOf('6% roll') !== -1 || adSetLower.indexOf('6% uaro') !== -1) {
      product = '6% AHA BHA Underarm Roll On';
    } else if (adSetLower.indexOf('underarm roll') !== -1 || adSetLower.indexOf('uaro') !== -1 || adSetLower.indexOf('underarm_roll') !== -1) {
      product = '4% AHA BHA Underarm Roll On';
    } else if (adSetLower.indexOf('bodywash') !== -1 || adSetLower.indexOf('body wash') !== -1 || adSetLower.indexOf('body_wash') !== -1) {
      if (adSetLower.indexOf('5%') !== -1 || adSetLower.indexOf('exfoliating') !== -1 || adNameLower.indexOf('5%') !== -1 || adNameLower.indexOf('exfoliating') !== -1) {
        product = '5% AHA BHA Exfoliating Body Wash';
      } else {
        product = '1% Salicylic Acid Body Wash';
      }
    } else if (combined.indexOf('1% sa bodywash') !== -1 || combined.indexOf('sa bodywash') !== -1 || combined.indexOf('salicylic') !== -1 || adNameLower.indexOf('sa_bodywash') !== -1) {
      product = '1% Salicylic Acid Body Wash';
    } else if (adNameLower.indexOf('underarm_roll') !== -1 || adNameLower.indexOf('underarm roll') !== -1 || combined.indexOf('uaro') !== -1) {
      product = '4% AHA BHA Underarm Roll On';
    } else {
      product = 'Body Category';
    }
  } else if (campaignLower.indexOf('hair category') !== -1 || adSetLower.indexOf('hair category') !== -1 || campaignLower.indexOf('hair_category') !== -1 || adSetLower.indexOf('hair_category') !== -1) {
    if (adSetLower.indexOf('category page') !== -1 || adSetLower.indexOf('category_page') !== -1) {
      product = (adSetLower.indexOf('body') !== -1 || campaignLower.indexOf('body') !== -1) ? 'Body Category' : 'Hair Category';
    } else if (adNameLower.indexOf('pack') !== -1 || adNameLower.indexOf('ahgp') !== -1) {
      product = 'AHGP (Roll On)';
    } else if (adSetLower.indexOf('pack') !== -1 || adSetLower.indexOf('ahgp') !== -1) {
      product = 'AHGP (Roll On)';
    } else if (adSetLower.indexOf('gummies') !== -1 || adNameLower.indexOf('gummies') !== -1) {
      product = 'Hair Gummies';
    } else if (adSetLower.indexOf('roll') !== -1 || adSetLower.indexOf('ahgs') !== -1) {
      if (adNameLower.indexOf('pack') !== -1 || adNameLower.indexOf('ahgp') !== -1) {
        product = 'AHGP (Roll On)';
      } else {
        product = 'Hair Growth Roll On';
      }
    } else if (combined.indexOf('hair growth pack') !== -1 || adNameLower.indexOf('hair_growth_pack') !== -1 || combined.indexOf('ahgp') !== -1) {
      product = 'AHGP (Roll On)';
    } else if (combined.indexOf('hair growth roll') !== -1 || combined.indexOf('advanced hair growth') !== -1 || combined.indexOf('advanced_hair_growth') !== -1 || combined.indexOf('ahgs') !== -1) {
      product = 'Hair Growth Roll On';
    } else if (combined.indexOf('hair gummies') !== -1 || combined.indexOf('hair_gummies') !== -1) {
      product = 'Hair Gummies';
    } else {
      product = 'Hair Category';
    }
  } else if (campaignLower === 'hair gummies asc_affluence' || adSetLower === 'hair gummies asc_affluence') {
    product = 'Kits';
  } else if (combined.indexOf('folli advanced') !== -1 || combined.indexOf('folli_advanced') !== -1 || combined.indexOf('folli_microneedle') !== -1 || combined.indexOf('folli microneedle') !== -1) {
    product = 'Hair Growth Roll On Folli Advanced';
  } else if (combined.indexOf('stretch mark') !== -1 || combined.indexOf('stretch_mark') !== -1) {
    product = 'Anti Stretch Mark Roll On';
  } else if (adNameLower.indexOf('30') !== -1 && (combined.indexOf('urea') !== -1 || combined.indexOf('urea_roll') !== -1)) {
    product = '30 Urea Foot Roll On';
  } else if (combined.indexOf('urea kit') !== -1 || combined.indexOf('urea skin') !== -1 || combined.indexOf('heel repair kit') !== -1) {
    if (adNameLower.indexOf('urea_roll') !== -1 || adNameLower.indexOf('urea roll') !== -1 || adNameLower.indexOf('urea_roll_on') !== -1) {
      product = '20% Urea Foot Roll On';
    } else {
      product = 'Urea Skin & Heel Repair Kit';
    }
  } else if (combined.indexOf('20% urea lotion') !== -1 || combined.indexOf('20 urea lotion') !== -1 || combined.indexOf('20_urea_lotion') !== -1 || combined.indexOf('20% ul') !== -1 || combined.indexOf('20 ul') !== -1 || combined.indexOf('20_ul') !== -1) {
    if (adNameLower.indexOf('10_urea') !== -1 || adNameLower.indexOf('10% urea') !== -1) {
      product = 'UL';
    } else {
      product = '20% UL';
    }
  } else if (combined.indexOf('10% urea body lotion') !== -1 || combined.indexOf('10% urea lotion') !== -1 || combined.indexOf('10 urea lotion') !== -1 || combined.indexOf('10_urea_lotion') !== -1 || combined.indexOf('urea lotion') !== -1 || combined.indexOf('urea_lotion') !== -1 || combined.indexOf('10% ul') !== -1 || combined.indexOf('10 ul') !== -1 || combined.indexOf('10_ul') !== -1 || combined.indexOf(' ul ') !== -1 || combined.slice(-3) === ' ul' || combined.indexOf(' ul|') !== -1 || combined.indexOf(' ul_') !== -1) {
    product = 'UL';
  } else if (combined.indexOf('ceramide') !== -1 || combined.indexOf('salicylic_ceramide') !== -1 || combined.indexOf('salicylic + ceramide') !== -1 || adNameLower.indexOf('salicylic_ceramide_bodywash') !== -1) {
    product = '1% Salicylic + Ceramide Bodywash';
  } else if (combined.indexOf('triple milk') !== -1 || combined.indexOf('triple_milk') !== -1) {
    product = 'Triple Milk Body Wash';
  } else if (combined.indexOf('lactic acid') !== -1 || combined.indexOf('lactic_acid') !== -1 || combined.indexOf('lactic lotion') !== -1 || combined.indexOf('lactic_lotion') !== -1) {
    product = '5% Lactic Acid Lotion';
  } else if (combined.indexOf('niacinamide') !== -1 || combined.indexOf('niacinamide_lotion') !== -1) {
    product = '10% Niacinamide Lotion';
  } else if (campaignLower.indexOf('6% underarm pigmentation kit campaign_asc_mcv_tier 2 & tier 3') !== -1) {
    product = '6% Underarm Pigmentation Kit';
  } else if (combined.indexOf('6% underarm pigmentation') !== -1 || combined.indexOf('6% underarm pigmentation kit') !== -1) {
    product = 'Advanced Underarm Pigmentation Kit';
  } else if (combined.indexOf('6% uaro') !== -1 || combined.indexOf('6%_uaro') !== -1 || combined.indexOf('6% underarm roll') !== -1 || combined.indexOf('6% aha bha underarm') !== -1 || combined.indexOf('6% roll on') !== -1) {
    product = '6% AHA BHA Underarm Roll On';
  } else if (combined.indexOf('4% underarm roll') !== -1 || combined.indexOf('4% & 6% underarm roll') !== -1 || combined.indexOf('4% aha bha underarm') !== -1) {
    product = '4% AHA BHA Underarm Roll On';
  } else if (combined.indexOf('advanced underarm pigmentation') !== -1 || combined.indexOf('underarm pigmentation kit') !== -1 || combined.indexOf('upk') !== -1 || adNameLower.indexOf('underarm_pigmentation_pack') !== -1 || adNameLower.indexOf('pigmentation_pack') !== -1 || combined.indexOf('pigmentation pack') !== -1) {
    product = 'Advanced Underarm Pigmentation Kit';
  } else if (combined.indexOf('underarm roll') !== -1 || combined.indexOf('underarm_roll') !== -1 || combined.indexOf('underarm_roll_on') !== -1 || combined.indexOf('uaro') !== -1) {
    product = '4% AHA BHA Underarm Roll On';
  } else if (combined.indexOf('urea roll') !== -1 || combined.indexOf('urea_roll') !== -1 || combined.indexOf('urea') !== -1) {
    product = '20% Urea Foot Roll On';
  } else if (combined.indexOf('calcium gummies') !== -1 || combined.indexOf('calcium_gummies') !== -1 || combined.indexOf('calcium') !== -1) {
    product = 'Calcium Gummies';
  } else if (combined.indexOf('magnesium lotion') !== -1 || combined.indexOf('magnesium_lotion') !== -1) {
    product = 'Magnesium Lotion';
  } else if (combined.indexOf('magnesium gummies') !== -1 || combined.indexOf('magnesium_gummies') !== -1 || combined.indexOf('magnesium') !== -1 || combined.indexOf('magnesium_glycinate') !== -1) {
    product = 'Magnesium Glycinate Gummies';
  } else if (combined.indexOf('glutathione') !== -1 || combined.indexOf('glutathione_gummies') !== -1) {
    product = 'Glutathione Gummies';
  } else if (combined.indexOf('shilajit') !== -1 || combined.indexOf('shilajit_gummies') !== -1) {
    product = 'Shilajit Gummies';
  } else if (combined.indexOf('hair gummies revised mrp') !== -1 || combined.indexOf('revised mrp') !== -1) {
    product = 'Hair Gummies (549)';
  } else if (combined.indexOf('hair gummies postpartum') !== -1 || combined.indexOf('hair_gummies_postpartum') !== -1 || combined.indexOf('postpartum') !== -1 || combined.indexOf('post_partum') !== -1 || combined.match(/\bpp\b/) || combined.indexOf('_pp_') !== -1 || combined.indexOf('_pp |') !== -1) {
    if (adSetLower.indexOf('new internal') !== -1 || adSetLower.indexOf('new_internal') !== -1) {
      product = 'Hair Gummies';
    } else {
      product = 'HG - PP';
    }
  } else if (combined.indexOf('hair growth pack') !== -1 || combined.indexOf('hair_growth_pack') !== -1 || combined.indexOf('ahgp') !== -1) {
    product = 'AHGP (Roll On)';
  } else if (combined.indexOf('hair gummies') !== -1 || combined.indexOf('hair_gummies') !== -1) {
    product = 'Hair Gummies';
  } else if (combined.indexOf('hair growth roll') !== -1 || combined.indexOf('advanced hair growth') !== -1 || combined.indexOf('advanced_hair_growth') !== -1 || combined.indexOf('ahgs') !== -1) {
    product = 'Hair Growth Roll On';
  } else if (combined.indexOf('keto shampoo') !== -1 || combined.indexOf('ketoconazole') !== -1 || combined.indexOf('dandruff shampoo') !== -1 || combined.indexOf('dandruff_shampoo') !== -1 || combined.indexOf('keto_shampoo') !== -1) {
    product = '1% Keto Shampoo';
  } else if (combined.indexOf('glycolic stick') !== -1 || combined.indexOf('glycolic_stick') !== -1 || combined.indexOf('exfoliating body stick') !== -1 || combined.indexOf('exfoliating_body_stick') !== -1 || combined.indexOf('glycolic') !== -1) {
    product = 'Glycolic Stick';
  } else if (combined.indexOf('hydration kit') !== -1 || combined.indexOf('hydration_kit') !== -1) {
    product = 'Hydration Kit';
  } else if (combined.indexOf('5% aha bha exfoliating body wash') !== -1 || combined.indexOf('5% aha bha bodywash') !== -1 || combined.indexOf('5% bodywash') !== -1 || combined.indexOf('exfoliating body wash') !== -1 || combined.indexOf('5%_aha_bha_bodywash') !== -1) {
    product = '5% AHA BHA Exfoliating Body Wash';
  } else if (combined.indexOf('1% sa bodywash adset') !== -1 && combined.indexOf('new launch') === -1) {
    product = '1% Salicylic Acid Body Wash - 250 mL';
  } else if (combined.indexOf('body acne pack') !== -1 || combined.indexOf('body_acne_pack') !== -1) {
    product = 'Body Acne Pack';
  } else if (combined.indexOf('sa bodywash') !== -1 || combined.indexOf('sa_bodywash') !== -1 || combined.indexOf('salicylic acid body wash') !== -1 || combined.indexOf('salicylic_acid_body_wash') !== -1 || combined.indexOf('sa body wash') !== -1) {
    product = '1% Salicylic Acid Body Wash';
  } else if (combined.indexOf('wintercare') !== -1) {
    product = 'Wintercare';
  } else if (combined.indexOf('winter care') !== -1 || combined.indexOf('winter_care') !== -1 || combined.indexOf('winter category') !== -1 || combined.indexOf('winter_category') !== -1) {
    if (combined.indexOf('051125') !== -1) {
      product = 'Wintercare';
    } else {
      product = 'Winter Care';
    }
  } else if (combined.indexOf('strawberry') !== -1 || combined.indexOf('strawberry_skin') !== -1 || combined.indexOf('straw') !== -1) {
    product = 'Straw';
  } else if (combined.indexOf('mehr_hair_growth') !== -1 || combined.indexOf('hgp') !== -1) {
    product = 'HGP';
  } else if (combined.indexOf('all_body_products') !== -1 || combined.indexOf('all_body') !== -1) {
    product = 'Body Category';
  } else if (combined.indexOf('all_hair_products') !== -1 || combined.indexOf('all_hair') !== -1) {
    product = 'Hair Category';
  } else if (combined.indexOf('skin_category') !== -1 || combined.indexOf('body_care') !== -1 || combined.indexOf('acne_treatment') !== -1 || combined.indexOf('sulphur') !== -1 || combined.indexOf('body care') !== -1) {
    product = 'Body Category';
  } else if (combined.indexOf('all_kit') !== -1 || combined.indexOf('kit_page') !== -1 || combined.indexOf('all kit') !== -1 || combined.indexOf('kit page') !== -1 || combined.indexOf('shop all') !== -1 || adNameLower.indexOf('pack_na') !== -1 || adNameLower.indexOf('pack,') !== -1 || campaignLower.indexOf('hgro') !== -1) {
    product = 'Kits';
  }
  
  if (adSetLower.indexOf('asc_all_kit_page') !== -1 || adSetLower.indexOf('all_kit_page') !== -1) {
    product = 'Kits';
  }
  
  // 2. AD TYPE
  var adType = 'Internal UGC';
  var parts = adNameLower.split('_');
  var firstPart = parts.length > 0 ? parts[0] : '';
  if (firstPart === 'influencer' || firstPart === 'infl' || firstPart === 'collab' || firstPart === 'creator' || firstPart === 'inf' || firstPart.indexOf('influ') === 0 || firstPart.indexOf('infl') === 0 || firstPart.indexOf('inf') === 0) {
    adType = 'Influencer';
  } else if (adNameLower.indexOf('influencer') !== -1 || adNameLower.indexOf('_infl_') !== -1 || adNameLower.indexOf('infl_') === 0 || adNameLower.indexOf('_infl') !== -1 || adNameLower.indexOf('collab') !== -1 || adNameLower.indexOf('creator') !== -1 || adNameLower.indexOf('_inf_') !== -1 || adNameLower.indexOf('_inf') !== -1 || adNameLower.indexOf('inf_') === 0) {
    adType = 'Influencer';
  }
  
  // 3. TYPE OF AD
  var typeOfAd = 'Video Ad';
  if (adNameLower.indexOf('carousel') !== -1) {
    typeOfAd = 'Carousel';
  } else if (adNameLower.indexOf('static') !== -1 || adNameLower.indexOf('image') !== -1) {
    typeOfAd = 'Static Creative';
  } else if (adNameLower.indexOf('voice_over') !== -1 || adNameLower.indexOf('voiceover') !== -1 || adNameLower.indexOf('reel') !== -1 || adNameLower.indexOf('video') !== -1 || adNameLower.indexOf('shorts') !== -1) {
    typeOfAd = (adType === 'Influencer') ? 'Reel' : 'Video Ad';
  }
  
  // 4. LANGUAGE
  var language = 'English';
  var languages = ['english', 'hindi', 'hinglish', 'marathi', 'tamil', 'telugu', 'kannada', 'malayalam', 'punjabi'];
  for (var i = 0; i < parts.length; i++) {
    var pClean = parts[i].trim();
    if (languages.indexOf(pClean) !== -1) {
      language = capitalizeWord(pClean);
      break;
    } else if (pClean === 'hindiandenglish' || pClean === 'englishandhindi' || pClean === 'hindiinenglishscript') {
      language = 'Hindi';
      break;
    }
  }
  
  // 5. INFLUENCER NAME
  var influencerName = 'Internal';
  if (campaignLower.indexOf('affluence') !== -1 || adSetLower.indexOf('affluence') !== -1) {
    influencerName = 'Affluence';
  } else if (adType === 'Internal UGC') {
    if (adNameLower.indexOf('bhavi') !== -1) {
      influencerName = 'Bhavi';
    } else if (adNameLower.indexOf('pooja') !== -1) {
      influencerName = 'Pooja';
    } else if (adNameLower.indexOf('esha shetty') !== -1 || adNameLower.indexOf('esha_shetty') !== -1) {
      influencerName = 'Esha Shetty';
    } else {
      influencerName = 'Internal';
    }
  } else if (adType === 'Influencer') {
    var skipTokens = [
      'influencer', 'voice', 'over', 'brand', 'music', 'reel', 'ugc', 'testimonial',
      'static', 'internal', 'model', 'reels', 'image', 'video', 'shorts', 'carousel',
      'infl', 'collab', 'creator', 'routine', 'before', 'after', 'before/after', 'trust',
      'safe', 'ingredients', 'pan', 'india', 'panindia', 'mumbai', 'delhi', 'bangalore',
      'kolkata', 'chennai', 'pune', 'hyderabad', 'signs', 'deficiency', 'stress', 'first',
      'age', 'hook', 'exp', 'bca', 'zero', 'white', 'sugar', 'habit', 'lp', 'zero_white_sugar',
      'zero_white_sugar_ingredient_led', 'prepaid', 'purchasers', 'detailed', 'targeting',
      'lookalike', 'broad', 'targeting', 'interest-based', 'purchase', 'hacks', 'republic_day',
      'republic', 'day', 'off', 'offer', 'price', 'slash', 'discount', 'sale', 'coupon', 'festive'
    ];
    
    var prodTokens = [
      'advanced', 'hair', 'growth', 'roll', 'on', 'pack', 'na', 'gummies', 'gummy',
      'urea', 'foot', 'lotion', 'bodywash', 'body_wash', 'body', 'wash', 'lactic',
      'acid', 'niacinamide', 'shilajit', 'calcium', 'magnesium', 'glycinate', 'glutathione',
      'shampoo', 'keto', 'ketoconazole', 'anti', 'dandruff', 'exfoliating', 'stick',
      'strawberry', 'skin', 'kit', 'hydration', 'anti-stretch', 'stretch', 'mark', 'serum',
      '6%', '4%', '10%', '20%', '30%', '5%', '1%', '7%', 'uaro', 'upk', 'ul', 'sa', 'hgp',
      'underarm', 'pigmentation', 'bodywash', 'body', 'wash', 'acne', 'triple', 'milk',
      'lactic', 'folli', 'microneedle', 'tan', 'sunscreen', 'sunscreens', 'de-tan', 'detan',
      'lotion', 'shampoo', 'glycolic', 'all', 'products', 'category', 'page', 'none',
      'postpartum', 'post_partum', 'pp', 'ahgp', 'ahgs', 'mrp', 'revised', 'winter',
      'wintercare', 'winter_care', 'care', 'straw', 'mehr', 'republic', 'day', 'aha', 'bha',
      // Product variants/flavors/descriptors — NOT person names
      'cherry', 'orange', 'peel', 'lavender', 'rose', 'mint', 'jasmine', 'coconut',
      'almond', 'green', 'blue', 'red', 'yellow', 'pink', 'black', 'white', 'gold',
      'fragrances', 'fragrance', 'scent', 'variant', 'new', 'launch', 'new_launch',
      'pack_na', 'pigmentation_pack', 'ml', '100ml', '200ml', '375ml', '250ml'
    ];
    
    var rawParts = adName.split('_');
    var infParts = [];
    for (var i = 0; i < rawParts.length; i++) {
      var pLow = rawParts[i].toLowerCase().trim();
      if (isLanguageToken(pLow)) {
        break;
      }
      if (skipTokens.indexOf(pLow) !== -1 || prodTokens.indexOf(pLow) !== -1) {
        continue;
      }
      if (pLow.match(/^\d+$/) || pLow.indexOf('exp') === 0 || pLow.match(/\d/)) {
        continue;
      }
      infParts.push(rawParts[i]);
    }
    
    var cleanedInfParts = [];
    for (var i = 0; i < infParts.length; i++) {
      var xLow = infParts[i].toLowerCase().trim();
      if ((xLow === 'dr' || xLow === 'dr.' || xLow === 'dr_') && !(i < infParts.length - 1 && infParts[i+1].toLowerCase().trim() === 'shivanti')) {
        continue;
      }
      cleanedInfParts.push(infParts[i]);
    }
    
    if (cleanedInfParts.length > 0) {
      influencerName = cleanedInfParts.map(capitalizeWord).join(' ');
      var nameMap = {
        'Surabhi Tiwari': 'Surabhi',
        'Surabhi': 'Surabhi',
        'Shilpa Dwivedi': 'Shilpa',
        'Shilpa': 'Shilpa',
        'Timsy': 'Timsy',
        'Chitwan': 'Chitwan Garg',
        'Chitwan Garg': 'Chitwan Garg',
        'Mehvi': 'Mehvi Thapa',
        'Mehvi Thapa': 'Mehvi Thapa',
        'Dr Mehvi Thapa': 'Mehvi Thapa',
        'Dr. Mehvi Thapa': 'Mehvi Thapa',
        'Suhana Grover': 'Suhana Grover',
        'Suhana': 'Suhana Grover',
        'Swati': 'Swati',
        'Swati Chauhan': 'Swati Chauhan',
        'Khushboo Gupta': 'Khushboo',
        'Khushboo': 'Khushboo',
        'Buvana Balaji': 'Buvana',
        'Buvana': 'Buvana',
        'Deepika Rani Sahu': 'Deepika Rani',
        'Deepika Rani': 'Deepika Rani',
        'Deepali Rathore': 'Deepali',
        'Deepali': 'Deepali',
        'Shamli Raizada': 'Shamli',
        'Shamli': 'Shamli',
        'Nikkita': 'Nikkitha',
        'Nikitta': 'Nikkitha',
        'Nikkitha': 'Nikkitha',
        'Sarita Choudhary': 'Sarita',
        'Sarita': 'Sarita',
        'Sumegha Sudan': 'Sumegha',
        'Sumegha': 'Sumegha',
        'Jiya Sharma': 'Jiya',
        'Jiya': 'Jiya',
        'Rini Khanna': 'Rini',
        'Rini': 'Rini',
        'Yogita Toora': 'Yogita',
        'Yogita': 'Yogita',
        'Jyoti Sinha': 'Jyoti',
        'Jyoti': 'Jyoti',
        'Vaishnavi Bhavsar': 'Vaishnavi',
        'Vaishnavi': 'Vaishnavi',
        'Shivanti': 'Dr. Shivanti',
        'Dr. Shivanti': 'Dr. Shivanti',
        'Amrita Khanal': 'Amrita',
        'Amrita': 'Amrita',
        'Pria Sethi': 'Pria',
        'Pria': 'Pria',
        'Mahi Deswal': 'Mahi',
        'Mahi': 'Mahi',
        'Ridhi Khanna': 'Riddhi Khanna',
        'Riddhi Khanna': 'Riddhi Khanna',
        'Pritha': 'Pritha Khot',
        'Pritha Khot': 'Pritha Khot'
      };
      if (nameMap[influencerName] !== undefined) {
        influencerName = nameMap[influencerName];
      }
    } else {
      var scanNames = ['Swati Chauhan', 'Swati', 'Chitwan Garg', 'Chitwan', 'Esha Shetty', 'Parichita Poddar', 'Tanvi Sanghvi', 'Shikha Rastogi', 'Satveer Kaur', 'Satveer', 'Surabhi Tiwari', 'Surabhi', 'Shilpa Dwivedi', 'Shilpa', 'Mehvi Thapa', 'Mehvi'];
      for (var i = 0; i < scanNames.length; i++) {
        if (adNameLower.indexOf(scanNames[i].toLowerCase()) !== -1) {
          influencerName = scanNames[i];
          break;
        }
      }
      if (influencerName === 'Internal') {
        influencerName = 'Influencer';
      }
    }
    
    var overrideNames = ['Bhavi', 'Poorvi', 'Swati Chauhan', 'Swati', 'Esha Shetty', 'Parichita Poddar', 'Chitwan Garg', 'Chitwan', 'Tanvi Sanghvi', 'Sarita', 'Namrata', 'Shivangi Jain', 'Shikha Rastogi', 'Pria', 'Raisha', 'Lulu', 'Mrunal', 'Satveer Kaur', 'Satveer', 'Surabhi Tiwari', 'Surabhi', 'Shilpa Dwivedi', 'Shilpa', 'Mehvi Thapa', 'Mehvi', 'Dr. Rashi', 'Rashi Soni', 'Rashi', 'Sakshi Komal Singh', 'Sakshi', 'Chahat Tewani', 'Chahat', 'Suhana Grover', 'Suhana', 'Pratishtha Singh', 'Pratishtha', 'Yogita Toora', 'Yogita', 'Jyoti Sinha', 'Jyoti', 'Vaishnavi Bhavsar', 'Vaishnavi', 'Shruti Bavisetti', 'Shruthi Bavisetti', 'Shivanti', 'Dr. Shivanti', 'Amrita Khanal', 'Amrita', 'Sreya Roy', 'Sreya', 'Rini Khanna', 'Rini', 'Pria Sethi', 'Ridhi Khanna', 'Riddhi Khanna', 'Mahi Deswal', 'Mahi', 'Pritha Khot', 'Pritha'];
    for (var i = 0; i < overrideNames.length; i++) {
      var nLow = overrideNames[i].toLowerCase();
      var rx = new RegExp('\\b' + nLow + '\\b');
      if (adNameLower.match(rx) || ('_' + adNameLower + '_').indexOf('_' + nLow + '_') !== -1) {
        influencerName = overrideNames[i];
        var nameMap = {
          'Swati Chauhan': 'Swati Chauhan',
          'Swati': 'Swati',
          'Khushboo Gupta': 'Khushboo',
          'Buvana Balaji': 'Buvana',
          'Deepika Rani Sahu': 'Deepika Rani',
          'Deepali Rathore': 'Deepali',
          'Shamli Raizada': 'Shamli',
          'Nikitta': 'Nikkitha',
          'Nikkita': 'Nikkitha',
          'Nikkitha': 'Nikkitha',
          'Sarita Choudhary': 'Sarita',
          'Sumegha Sudan': 'Sumegha',
          'Shikha Rastogi': 'Shikha',
          'Jiya Sharma': 'Jiya',
          'Rini Khanna': 'Rini',
          'Rini': 'Rini',
          'Yogita Toora': 'Yogita',
          'Yogita': 'Yogita',
          'Jyoti Sinha': 'Jyoti',
          'Jyoti': 'Jyoti',
          'Vaishnavi Bhavsar': 'Vaishnavi',
          'Vaishnavi': 'Vaishnavi',
          'Shruti Bavisetti': 'Shruthi Bavisetti',
          'Shruthi Bavisetti': 'Shruthi Bavisetti',
          'Shivanti': 'Dr. Shivanti',
          'Dr. Shivanti': 'Dr. Shivanti',
          'Amrita Khanal': 'Amrita',
          'Amrita': 'Amrita',
          'Sreya Roy': 'Sreya',
          'Sreya': 'Sreya',
          'Pria': 'Pria',
          'Pria Sethi': 'Pria',
          'Tanvi': 'Tanvi',
          'Tanvi Sanghvi': 'Tanvi Sanghvi',
          'Mahi Deswal': 'Mahi',
          'Mahi': 'Mahi',
          'Ridhi Khanna': 'Riddhi Khanna',
          'Riddhi Khanna': 'Riddhi Khanna',
          'Suhana': 'Suhana Grover',
          'Suhana Grover': 'Suhana Grover',
          'Pritha Khot': 'Pritha Khot',
          'Pritha': 'Pritha Khot'
        };
        if (nameMap[influencerName] !== undefined) {
          influencerName = nameMap[influencerName];
        }
        break;
      }
    }
  }
  
  // Derived Category
  var category = 'Skincare';
  var prodLower = product.toLowerCase();
  if (prodLower.indexOf('hair') !== -1 || prodLower.indexOf('gummies') !== -1 || prodLower.indexOf('serum') !== -1 || prodLower.indexOf('shampoo') !== -1 || prodLower.indexOf('hgp') !== -1 || prodLower.indexOf('hg - pp') !== -1 || prodLower.indexOf('ahgp') !== -1) {
    category = 'Haircare';
  }
  
  // Narrative — extract from ad name segment AFTER language token
  // Ad name structure: [type]_[format]_[product]_[influencer]_[language]_[NARRATIVE]_[targeting]_[id]_[date]
  var narrative = '';
  var adNameParts = adNameLower.split('_');
  var langPosIdx = -1;
  for (var ni = 0; ni < adNameParts.length; ni++) {
    if (isLanguageToken(adNameParts[ni].trim())) {
      langPosIdx = ni;
      break;
    }
  }
  var rawNarrStr = '';
  
  var stopWords = [
    'influencer','infl','collab','creator','inf','internal','model','static','image',
    'video','reel','reels','shorts','carousel','voice','over','voiceover','music','ugc',
    'testimonial','review','unboxing','story','brand','916','11','45','30s','15s',
    'pan','india','panindia','bca','exp','zero','white','sugar','lp','prepaid','purchasers',
    'detailed','targeting','lookalike','broad','interest','purchase','others','none','shot','na',
    'new','launch','asc','reedit','advanced','hair','growth','roll','on','pack','gummies',
    'gummy','urea','foot','lotion','bodywash','body','wash','lactic','acid','niacinamide',
    'shilajit','calcium','magnesium','glycinate','glutathione','shampoo','keto','ketoconazole',
    'anti','dandruff','exfoliating','stick','strawberry','skin','kit','hydration','stretch',
    'mark','serum','6%','4%','10%','20%','30%','5%','1%','7%','uaro','upk','ul','sa','hgp',
    'underarm','pigmentation','acne','triple','milk','folli','microneedle','tan','sunscreen',
    'sunscreens','de-tan','detan','glycolic','all','products','category','page','postpartum',
    'post','partum','pp','ahgp','ahgs','mrp','revised','winter','wintercare','care','straw',
    'mehr','aha','bha','cherry','orange','peel','lavender','rose','mint','jasmine','coconut',
    'almond','green','blue','red','yellow','pink','black','white','gold','fragrances',
    'fragrance','scent','variant','ml','100ml','200ml','375ml','250ml',
    'english','hindi','hinglish','marathi','tamil','telugu','kannada','malayalam','punjabi',
    'hindiandenglish','englishandhindi','hindiinenglishscript','hindiinhindiscript','hindiinromanscript','hindiinenglish'
  ];
  
  var knownInfluencerNames = [
    'bhavi','poorvi','swati','chitwan','surabhi','shilpa','mehvi','rashi','sakshi',
    'chahat','pratishtha','chelsy','suhana','yogita','jyoti','vaishnavi','amrita',
    'sreya','rini','mahi','pritha','pria','nikkitha','deepali','shamli','ishita',
    'harman','satveer','timsy','khushboo','buvana','deepika','jiya','sumegha','namrata',
    'tiwari','dwivedi','garg','thapa','grover','chauhan','sahu','rathore','raizada',
    'bavisetti','khanal','khanna','deswal','khot','sethi','singh','soni','komal','rani',
    'akshaya','jaiswal','sidhu','toora','sinha','bhavsar','kumar','sharma','gupta',
    'balaji','poddar','sanghvi','rastogi','kaur','raisha','lulu','mrunal','shivangi','dr',
    'bharti','masoom','pathania','anupallavi','gowda','saumya','bisla','kratika','mohan','riya'
  ];
  stopWords = stopWords.concat(knownInfluencerNames);
  
  var narrativeTokens = [];
  var startIndex = langPosIdx !== -1 ? langPosIdx + 1 : 0;
  
  for (var ni = startIndex; ni < adNameParts.length; ni++) {
    var pt = adNameParts[ni].trim();
    if (!pt) continue;
    if (/^\d+$/.test(pt) && pt.length > 3) break; // ID or long number
    if (/^\d{6}/.test(pt)) break; // date pattern
    if (/^exp/.test(pt) && /\d/.test(pt)) break; // experiment tag
    if (pt === 'pan' || pt === 'india' || pt === 'panindia' || pt === 'bca') break;
    
    // Skip known stop words (products, formats, influencer names, etc.)
    if (stopWords.indexOf(pt) !== -1) continue;
    
    narrativeTokens.push(pt);
  }
  rawNarrStr = narrativeTokens.join('_');
  
  // Map to standard narrative categories using the FULL ad name
  var narrativeSearch = adNameLower;
  
  // === COMPREHENSIVE NARRATIVE MATCHING (most specific first) ===
  // Uses full ad name (narrativeSearch = adNameLower) so even old/broken naming is rescued
  if (narrativeSearch.indexOf('reasons_to_not') !== -1 || narrativeSearch.indexOf('reasons to not') !== -1 ||
      narrativeSearch.indexOf('negative_hook') !== -1 || narrativeSearch.indexOf('negative hook') !== -1) {
    narrative = 'Negative Hook';

  } else if (narrativeSearch.indexOf('doctor_recommended') !== -1 || narrativeSearch.indexOf('doctor recommended') !== -1 ||
             narrativeSearch.indexOf('doctor-recommended') !== -1 || narrativeSearch.indexOf('ca_focused') !== -1) {
    narrative = 'Doctor Recommended';

  } else if (narrativeSearch.indexOf('ingredient_led') !== -1 || narrativeSearch.indexOf('ingredient-led') !== -1 ||
             narrativeSearch.indexOf('ingredient led') !== -1 || narrativeSearch.indexOf('ingredients_led') !== -1) {
    narrative = 'Ingredient-led';

  } else if (narrativeSearch.indexOf('product_centric') !== -1 || narrativeSearch.indexOf('product_first') !== -1 ||
             narrativeSearch.indexOf('product first') !== -1 || narrativeSearch.indexOf('non_messy') !== -1) {
    narrative = 'Product-Centric';

  } else if (narrativeSearch.indexOf('before_after') !== -1 || narrativeSearch.indexOf('before/after') !== -1 ||
             narrativeSearch.indexOf('90_days') !== -1 || narrativeSearch.indexOf('90days') !== -1 ||
             narrativeSearch.indexOf('weekly_consistency') !== -1) {
    narrative = 'Before/After';

  } else if (narrativeSearch.indexOf('problem_solution') !== -1 || narrativeSearch.indexOf('skit_format') !== -1 ||
             narrativeSearch.indexOf('skit format') !== -1) {
    narrative = 'Problem/Solution';

  } else if (narrativeSearch.indexOf('difference_in') !== -1 || narrativeSearch.indexOf('transformation') !== -1) {
    narrative = 'Transformation';

  } else if (narrativeSearch.indexOf('aspirational_hook') !== -1 || narrativeSearch.indexOf('aspirational') !== -1) {
    narrative = 'Aspirational';

  // Educative — before concern-led so 'what_is_' and 'myth' don't fall through
  } else if (narrativeSearch.indexOf('what_is_') !== -1 || narrativeSearch.indexOf('shilajit_for_women') !== -1 ||
             narrativeSearch.indexOf('informational') !== -1 || narrativeSearch.indexOf('myth_vs_truth') !== -1 ||
             narrativeSearch.indexOf('hair_test') !== -1 || narrativeSearch.indexOf('p2c_educative') !== -1 ||
             narrativeSearch.indexOf('educative') !== -1) {
    narrative = 'Educative';

  } else if (narrativeSearch.indexOf('home_remedies') !== -1 || narrativeSearch.indexOf('home_remedy') !== -1 ||
             narrativeSearch.indexOf('home remedies') !== -1 || narrativeSearch.indexOf('home remedy') !== -1) {
    narrative = 'Home Remedies';

  } else if (narrativeSearch.indexOf('cracked_heels') !== -1) {
    narrative = 'Cracked Heels';

  } else if (narrativeSearch.indexOf('comparison') !== -1) {
    narrative = 'Comparison';

  } else if (narrativeSearch.indexOf('rant_style') !== -1) {
    narrative = 'Rant Style';

  } else if (narrativeSearch.indexOf('republic_day') !== -1 || narrativeSearch.indexOf('independence_day') !== -1 ||
             narrativeSearch.indexOf('seasonal_festive') !== -1 || narrativeSearch.indexOf('festive_offer') !== -1 ||
             narrativeSearch.indexOf('festive') !== -1) {
    narrative = 'Seasonal/Festive';

  } else if (narrativeSearch.indexOf('wintercare') !== -1 || narrativeSearch.indexOf('winter_care') !== -1 ||
             narrativeSearch.indexOf('winter') !== -1) {
    narrative = 'Seasonal/Winter';

  } else if (narrativeSearch.indexOf('life_hacks') !== -1 || narrativeSearch.indexOf('hacks') !== -1) {
    narrative = 'Life Hacks';

  } else if (narrativeSearch.indexOf('nutritional_deficien') !== -1 || narrativeSearch.indexOf('deficiency') !== -1 ||
             narrativeSearch.indexOf('signs_of') !== -1) {
    narrative = 'Deficiency/Educative';

  } else if (narrativeSearch.indexOf('routine_led') !== -1 || narrativeSearch.indexOf('routine') !== -1) {
    narrative = 'Routine/Habit';

  } else if (narrativeSearch.indexOf('trust/safe') !== -1 || narrativeSearch.indexOf('trust_safe') !== -1 ||
             narrativeSearch.indexOf('safe_ingredients') !== -1 ||
             (narrativeSearch.indexOf('trust') !== -1 && (narrativeSearch.indexOf('ingredient') !== -1 || narrativeSearch.indexOf('safe') !== -1))) {
    narrative = 'Trust/Safe Ingredients';

  } else if (narrativeSearch.indexOf('all_in_one') !== -1 || narrativeSearch.indexOf('detan') !== -1 ||
             narrativeSearch.indexOf('de-tan') !== -1) {
    narrative = 'All In One/De-tan';

  // Offer-led — catch all offer variants before generic 'sale'
  } else if (narrativeSearch.indexOf('price_slash') !== -1 || narrativeSearch.indexOf('offer_call_out') !== -1 ||
             narrativeSearch.indexOf('offer_led') !== -1 || narrativeSearch.indexOf('offer_static') !== -1 ||
             narrativeSearch.indexOf('seasonal_offer') !== -1 || narrativeSearch.indexOf('discount') !== -1 ||
             narrativeSearch.indexOf('coupon') !== -1 || narrativeSearch.indexOf('sale') !== -1) {
    narrative = 'Offer-led';

  } else if (narrativeSearch.indexOf('voxpop') !== -1 || narrativeSearch.indexOf('reaction_voxpop') !== -1 ||
             narrativeSearch.indexOf('testimonial') !== -1 || narrativeSearch.indexOf('social_proof') !== -1 ||
             narrativeSearch.indexOf('unboxing') !== -1 || narrativeSearch.indexOf('review') !== -1) {
    narrative = 'Social Proof/UGC';

  } else if (narrativeSearch.indexOf('usps') !== -1 || narrativeSearch.indexOf('48_hr_hydration') !== -1 ||
             narrativeSearch.indexOf('benefits_led') !== -1 || narrativeSearch.indexOf('4_fragrances') !== -1 ||
             narrativeSearch.indexOf('benefit') !== -1 || narrativeSearch.indexOf('fragrance') !== -1) {
    narrative = 'Benefits/Feature-led';

  // Concern-led — broad category, covers specific conditions
  } else if (narrativeSearch.indexOf('concern_led') !== -1 || narrativeSearch.indexOf('hyperpigmentation') !== -1 ||
             narrativeSearch.indexOf('body_odour') !== -1 || narrativeSearch.indexOf('body_acne') !== -1 ||
             narrativeSearch.indexOf('bacne') !== -1 || narrativeSearch.indexOf('itchy_skin') !== -1 ||
             narrativeSearch.indexOf('hair_fall') !== -1 || narrativeSearch.indexOf('hairfall') !== -1 ||
             narrativeSearch.indexOf('hair_thinning') !== -1 || narrativeSearch.indexOf('hair_thin') !== -1 ||
             narrativeSearch.indexOf('dry_skin') !== -1 || narrativeSearch.indexOf('snake_skin') !== -1 ||
             narrativeSearch.indexOf('sleep_issues') !== -1 || narrativeSearch.indexOf('sleepless_nights') !== -1 ||
             narrativeSearch.indexOf('anxiety') !== -1 || narrativeSearch.indexOf('stress') !== -1 ||
             narrativeSearch.indexOf('postpartum') !== -1 || narrativeSearch.indexOf('post_partum') !== -1 ||
             narrativeSearch.indexOf('feeling_confident') !== -1 || narrativeSearch.indexOf('feeling_conscious') !== -1 ||
             narrativeSearch.indexOf('struggling_with') !== -1 || narrativeSearch.indexOf('concern') !== -1 ||
             narrativeSearch.indexOf('oestrogen') !== -1) {
    narrative = 'Concern-led';

  } else if (narrativeSearch.indexOf('founder_story') !== -1 || narrativeSearch.indexOf('founder') !== -1 ||
             narrativeSearch.indexOf('journey') !== -1) {
    narrative = 'Founder Story';

  } else if (narrativeSearch.indexOf('how_to_application') !== -1 || narrativeSearch.indexOf('application') !== -1) {
    narrative = 'How-to/Application';

  } else if (narrativeSearch.indexOf('brand_story') !== -1 || narrativeSearch.indexOf('brand story') !== -1) {
    narrative = 'Brand Story';

  } else if (narrativeSearch.indexOf('ugc') !== -1) {
    narrative = 'Social Proof/UGC';

  } else if (rawNarrStr) {
    // Has extracted tokens after language but none matched standard categories — use raw
    var rawParts = rawNarrStr.split('_').filter(function(w) {
      return w.length > 1 && !/^\d+$/.test(w);
    }).slice(0, 3);
    narrative = rawParts.map(capitalizeWord).join(' ');
  }
  
  // Final fallback: if still blank AND it's a static ad → Product-Centric
  if (!narrative) {
    var fmtLow = typeOfAd ? typeOfAd.toLowerCase() : '';
    if (fmtLow === 'static creative' || adNameLower.indexOf('static') !== -1 || adNameLower.indexOf('carousel') !== -1) {
      narrative = 'Product-Centric';
    }
  }
  

  
  return {
    product: product,
    adType: adType,
    typeOfAd: typeOfAd,
    influencerName: influencerName,
    narrative: narrative,
    language: language,
    category: category
  };
}
