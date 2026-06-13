// ============================================================
//  AUTO SYNC ENGINE — AutoSync.gs
//  Pulls data daily from Meta Ads, Shopify into Google Sheets
//  Google Ads data is pulled via Google Ads Scripts (see guide)
// ============================================================

// ─── MAIN DAILY TRIGGER FUNCTION ─────────────────────────
// This runs automatically every morning at 6 AM IST
function dailyAutoSync() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];
  log.push('[' + new Date().toLocaleString('en-IN') + '] Daily Auto-Sync Started');

  var cfg = loadSyncConfig(ss);
  if (!cfg) {
    logToSheet(ss, ['Auto-sync failed: CONFIG sheet missing API keys. Please fill them in.']);
    return;
  }

  try {
    var metaRows = fetchMetaAdsData(cfg);
    if (metaRows.length > 0) {
      appendToSheet(ss, 'META_DATA', metaRows);
      log.push('✅ Meta Ads: ' + metaRows.length + ' rows added');
    } else {
      log.push('⚠️ Meta Ads: No new rows (check token or date range)');
    }
  } catch(e) {
    log.push('❌ Meta Ads Error: ' + e.message);
  }

  try {
    var shopifyRows = fetchShopifyOrders(cfg);
    if (shopifyRows.length > 0) {
      appendToSheet(ss, 'SALES_DATA', shopifyRows);
      log.push('✅ Shopify: ' + shopifyRows.length + ' orders added');
    } else {
      log.push('⚠️ Shopify: No new orders found for yesterday');
    }
  } catch(e) {
    log.push('❌ Shopify Error: ' + e.message);
  }

  try {
    computeAndWriteDailyStats(ss);
    log.push('✅ DAILY_STATS: Recalculated and updated');
  } catch(e) {
    log.push('❌ DAILY_STATS Error: ' + e.message);
  }

  logToSheet(ss, log);
  log.push('[Done] Auto-Sync Complete');
}

// ─── META ADS API FETCH ───────────────────────────────────
function fetchMetaAdsData(cfg) {
  if (!cfg.metaToken || !cfg.metaAccountId) {
    throw new Error('Meta Access Token or Account ID missing in CONFIG sheet');
  }

  // Fetch yesterday's data at ad level (campaign + adset + ad)
  var yesterday = getYesterdayDate();
  var fields = [
    'date_start', 'account_name', 'campaign_name', 'adset_name', 'ad_name',
    'spend', 'impressions', 'clicks', 'actions', 'action_values',
    'video_thruplay_watched_actions', 'video_p25_watched_actions',
    'video_p50_watched_actions', 'video_p75_watched_actions',
    'video_p95_watched_actions', 'video_p100_watched_actions',
    'landing_page_views', 'unique_clicks', 'adset_id', 'ad_id'
  ].join(',');

  var url = 'https://graph.facebook.com/v19.0/act_' + cfg.metaAccountId + '/insights'
    + '?level=ad'
    + '&fields=' + encodeURIComponent(fields)
    + '&time_range={"since":"' + yesterday + '","until":"' + yesterday + '"}'
    + '&limit=500'
    + '&access_token=' + cfg.metaToken;

  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var json = JSON.parse(response.getContentText());

  if (json.error) {
    throw new Error('Meta API Error: ' + json.error.message);
  }

  var rows = [];
  (json.data || []).forEach(function(row) {
    var purchases = getActionValue(row.actions, 'purchase') || 0;
    var purchaseValue = getActionValue(row.action_values, 'purchase') || 0;
    var atc = getActionValue(row.actions, 'add_to_cart') || 0;
    var lpv = row.landing_page_views || 0;
    var v3s = getActionValue(row.video_thruplay_watched_actions, null, true) || 0;

    rows.push([
      row.date_start,                          // Day
      row.account_name || '',                  // Account Name
      row.campaign_name || '',                 // Campaign name
      row.adset_name || '',                    // Ad set name
      row.ad_name || '',                       // Ad name
      'INR',                                   // Currency
      parseFloat(row.spend) || 0,              // Amount spent (INR)
      parseInt(row.impressions) || 0,          // Impressions
      parseInt(row.clicks) || 0,               // Link clicks
      parseInt(lpv) || 0,                      // Landing page views
      parseInt(atc) || 0,                      // Add To Cart
      parseInt(purchases) || 0,               // Purchases
      parseFloat(purchaseValue) || 0,         // Purchases conversion value
      parseInt(v3s) || 0,                      // 3-second video plays
      getActionValue(row.video_p25_watched_actions, null, true) || 0, // 25%
      getActionValue(row.video_p50_watched_actions, null, true) || 0, // 50%
      getActionValue(row.video_p75_watched_actions, null, true) || 0, // 75%
      getActionValue(row.video_p95_watched_actions, null, true) || 0, // 95%
      getActionValue(row.video_p100_watched_actions, null, true) || 0, // 100%
      '',  // Website URL (not in API insights, fill via ad creative separately)
      row.adset_id || '', // Ad set ID
      row.ad_id || '',    // Ad ID
      // Product, Ad Type, Type of Ad, Influencer Name, Narrative, Language
      // → auto-parsed from campaign name via nomenclature in parseProductFromName()
      parseProductFromName(row.campaign_name),   // Product
      parseAdTypeFromName(row.ad_name),          // Ad Type
      parseFormatFromName(row.ad_name),          // Type of Ad
      parseInfluencerFromName(row.ad_name),      // Influencer Name
      parseNarrativeFromName(row.ad_name),       // Narrative
      parseLanguageFromName(row.ad_name)         // Language
    ]);
  });

  return rows;
}

// ─── SHOPIFY ORDERS FETCH ─────────────────────────────────
function fetchShopifyOrders(cfg) {
  if (!cfg.shopifyStore || !cfg.shopifyToken) {
    throw new Error('Shopify Store URL or Access Token missing in CONFIG sheet');
  }

  var yesterday = getYesterdayDate();
  var todayStr = getTodayDate();

  // Fetch orders created yesterday
  var url = 'https://' + cfg.shopifyStore + '/admin/api/2024-01/orders.json'
    + '?status=any'
    + '&created_at_min=' + yesterday + 'T00:00:00+05:30'
    + '&created_at_max=' + todayStr + 'T00:00:00+05:30'
    + '&limit=250'
    + '&fields=id,email,created_at,total_price,line_items,source_name';

  var response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: {
      'X-Shopify-Access-Token': cfg.shopifyToken,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });

  var json = JSON.parse(response.getContentText());
  if (json.errors) {
    throw new Error('Shopify API Error: ' + JSON.stringify(json.errors));
  }

  var rows = [];
  var cohortMonth = yesterday.substring(0, 7); // e.g. "2024-09"

  (json.orders || []).forEach(function(order) {
    var productName = order.line_items && order.line_items[0]
      ? order.line_items[0].title : 'Unknown';
    rows.push([
      order.id,                                // Order ID
      order.email || ('order_' + order.id),   // Customer ID
      order.created_at.substring(0, 10),      // Order Date
      parseFloat(order.total_price) || 0,     // Revenue
      productName,                            // Product
      'Shopify',                              // Platform
      cohortMonth                             // Cohort Month
    ]);
  });

  return rows;
}

// ─── DAILY STATS COMPUTATION ──────────────────────────────
function computeAndWriteDailyStats(ss) {
  var yesterday = getYesterdayDate();
  var metaSheet = ss.getSheetByName('META_DATA');
  var googleSheet = ss.getSheetByName('GOOGLE_DATA');
  var salesSheet = ss.getSheetByName('SALES_DATA');
  var dailySheet = ss.getSheetByName('DAILY_STATS');

  if (!dailySheet) return;

  var totalSpend = 0, totalConversions = 0, totalRevenue = 0;

  // Sum META_DATA for yesterday
  if (metaSheet) {
    var metaData = metaSheet.getDataRange().getValues();
    metaData.forEach(function(row, i) {
      if (i === 0) return; // skip header
      var date = String(row[0]).substring(0, 10);
      if (date === yesterday) {
        totalSpend += parseFloat(row[6]) || 0;       // Amount spent
        totalConversions += parseInt(row[11]) || 0;  // Purchases
        totalRevenue += parseFloat(row[12]) || 0;    // Purchase value
      }
    });
  }

  // Sum GOOGLE_DATA for yesterday
  if (googleSheet) {
    var googleData = googleSheet.getDataRange().getValues();
    googleData.forEach(function(row, i) {
      if (i === 0) return;
      var date = String(row[0]).substring(0, 10);
      if (date === yesterday) {
        totalSpend += parseFloat(row[10]) || 0;     // Spend column
        totalConversions += parseInt(row[14]) || 0; // Conversions
        totalRevenue += parseFloat(row[15]) || 0;   // Revenue
      }
    });
  }

  // AOV from Shopify
  var totalOrders = 0, shopifyRevenue = 0;
  if (salesSheet) {
    var salesData = salesSheet.getDataRange().getValues();
    salesData.forEach(function(row, i) {
      if (i === 0) return;
      var date = String(row[2]).substring(0, 10);
      if (date === yesterday) {
        shopifyRevenue += parseFloat(row[3]) || 0;
        totalOrders++;
      }
    });
  }

  var cac = totalConversions > 0 ? totalSpend / totalConversions : 0;
  var aov = totalOrders > 0 ? shopifyRevenue / totalOrders : 0;
  var roas = totalSpend > 0 ? totalRevenue / totalSpend : 0;

  // Append to DAILY_STATS
  dailySheet.appendRow([yesterday, totalSpend, totalConversions, cac, aov, roas]);
}

// ─── HELPER: APPEND ROWS (deduplicate by date with high performance batch write) ───
function appendToSheet(ss, sheetName, newRows) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || newRows.length === 0) return;

  // Detect the date of the incoming data from the first column of the first row
  var targetDate = String(newRows[0][0]).substring(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(targetDate)) {
    targetDate = getYesterdayDate();
  }

  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  // Keep the header row (index 0) and filter out any existing rows matching targetDate
  var filteredRows = [data[0]];
  for (var i = 1; i < data.length; i++) {
    var rowDate = String(data[i][0]).substring(0, 10);
    if (rowDate !== targetDate) {
      filteredRows.push(data[i]);
    }
  }
  
  // Append new rows
  newRows.forEach(function(row) {
    filteredRows.push(row);
  });
  
  // Clear the existing values in the sheet (keeps formatting rules)
  sheet.clearContents();
  
  // Batch write all rows
  sheet.getRange(1, 1, filteredRows.length, filteredRows[0].length).setValues(filteredRows);
  
  // Keep the sheet compact: delete excess blank rows at the bottom (keeps sheet running lightning fast)
  var maxRows = sheet.getMaxRows();
  var requiredRows = filteredRows.length;
  if (maxRows > requiredRows + 50) {
    sheet.deleteRows(requiredRows + 1, maxRows - requiredRows - 50); // Keep a 50-row buffer
  }
}

// ─── HELPER: LOAD CONFIG FROM CONFIG SHEET ───────────────
function loadSyncConfig(ss) {
  var cfgSheet = ss.getSheetByName('CONFIG');
  if (!cfgSheet) return null;

  var data = cfgSheet.getDataRange().getValues();
  var cfg = {};
  data.forEach(function(row) {
    var key = String(row[0]).trim();
    var val = String(row[1]).trim();
    if (key === 'META_ACCESS_TOKEN') cfg.metaToken = val;
    if (key === 'META_AD_ACCOUNT_ID') cfg.metaAccountId = val;
    if (key === 'SHOPIFY_STORE_URL') cfg.shopifyStore = val;
    if (key === 'SHOPIFY_ACCESS_TOKEN') cfg.shopifyToken = val;
  });
  return cfg;
}

// ─── HELPER: PARSE PRODUCT FROM CAMPAIGN NAME ────────────
// Reads "Bodywash 375ml | TOF | Lookalike | 20240901" → "Bodywash 375ml"
function parseProductFromName(name) {
  if (!name) return '';
  var parts = name.split('|');
  return parts[0] ? parts[0].trim() : '';
}
function parseAdTypeFromName(name) {
  if (!name) return '';
  var parts = name.split('|');
  return parts[0] ? parts[0].trim() : ''; // Format
}
function parseFormatFromName(name) {
  if (!name) return '';
  var n = name.toLowerCase();
  if (n.indexOf('reel') > -1 || n.indexOf('video') > -1) return 'Reel';
  if (n.indexOf('static') > -1 || n.indexOf('image') > -1) return 'Static';
  if (n.indexOf('carousel') > -1) return 'Carousel';
  return 'Static';
}
function parseInfluencerFromName(name) {
  if (!name) return '';
  var parts = name.split('|');
  // Convention: "Reel | Pain_Point | UGC_InfluencerName | v1"
  if (parts.length >= 3) return parts[2].trim().replace('UGC_', '');
  return '';
}
function parseNarrativeFromName(name) {
  if (!name) return '';
  var parts = name.split('|');
  if (parts.length >= 2) return parts[1].trim().replace(/_/g, ' ');
  return '';
}
function parseLanguageFromName(name) {
  if (!name) return '';
  var n = name.toLowerCase();
  if (n.indexOf('hindi') > -1) return 'Hindi';
  if (n.indexOf('english') > -1 || n.indexOf('eng') > -1) return 'English';
  if (n.indexOf('tamil') > -1) return 'Tamil';
  if (n.indexOf('telugu') > -1) return 'Telugu';
  if (n.indexOf('bengali') > -1) return 'Bengali';
  return 'English'; // default
}

// ─── HELPER: GET ACTION VALUE FROM META ACTIONS ARRAY ────
function getActionValue(actionsArr, type, firstOnly) {
  if (!actionsArr || !actionsArr.length) return 0;
  if (firstOnly) return parseFloat(actionsArr[0].value) || 0;
  var match = actionsArr.find(function(a) { return a.action_type === type; });
  return match ? parseFloat(match.value) : 0;
}

// ─── HELPER: DATE UTILITIES ───────────────────────────────
function getYesterdayDate() {
  var d = new Date();
  d.setDate(d.getDate() - 1);
  return Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd');
}
function getTodayDate() {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
}

// ─── HELPER: LOG TO SYNC_LOG SHEET ───────────────────────
function logToSheet(ss, lines) {
  var logSheet = ss.getSheetByName('SYNC_LOG');
  if (!logSheet) logSheet = ss.insertSheet('SYNC_LOG');
  lines.forEach(function(line) {
    logSheet.appendRow([new Date(), line]);
  });
}

// ─── SETUP: CREATE DAILY TRIGGER (run once) ──────────────
function setupDailySyncTrigger() {
  // Remove old triggers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyAutoSync') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create new trigger: every day at 6 AM IST
  ScriptApp.newTrigger('dailyAutoSync')
    .timeBased()
    .everyDays(1)
    .atHour(0) // 12 AM UTC = 5:30 AM IST ≈ 6 AM IST
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Daily Auto-Sync is now set up!\n\n' +
    'Every morning at ~6 AM IST, the system will automatically:\n' +
    '• Pull yesterday\'s Meta Ads data → META_DATA tab\n' +
    '• Pull yesterday\'s Shopify orders → SALES_DATA tab\n' +
    '• Compute blended stats → DAILY_STATS tab\n\n' +
    'For Google Ads, use the Google Ads Google Sheets Add-on (see guide).\n\n' +
    '⚠️ Make sure your API keys are filled in the CONFIG sheet!'
  );
}

// ─── SETUP: WRITE API KEY ROWS TO CONFIG SHEET ───────────
function setupConfigApiKeys() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfgSheet = ss.getSheetByName('CONFIG');
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet('CONFIG');
  }

  // Check if API key rows already exist
  var data = cfgSheet.getDataRange().getValues();
  var existingKeys = data.map(function(r) { return r[0]; });

  var apiKeyRows = [
    ['META_ACCESS_TOKEN', '', 'Your Meta long-lived access token (from developers.facebook.com)'],
    ['META_AD_ACCOUNT_ID', '', 'Your Meta Ad Account ID — numbers only (e.g. 1234567890)'],
    ['SHOPIFY_STORE_URL', '', 'e.g. yourstore.myshopify.com (no https://)'],
    ['SHOPIFY_ACCESS_TOKEN', '', 'From Shopify Admin → Apps → Private apps → Admin API access token']
  ];

  apiKeyRows.forEach(function(row) {
    if (existingKeys.indexOf(row[0]) === -1) {
      cfgSheet.appendRow(row);
    }
  });

  // Format the CONFIG sheet
  cfgSheet.getRange(1, 1, cfgSheet.getLastRow(), 3)
    .setFontFamily('Arial')
    .setFontSize(11);
  cfgSheet.setColumnWidth(1, 220);
  cfgSheet.setColumnWidth(2, 320);
  cfgSheet.setColumnWidth(3, 480);

  SpreadsheetApp.getUi().alert(
    '✅ CONFIG sheet updated with API key rows!\n\n' +
    'Now fill in column B:\n' +
    '• META_ACCESS_TOKEN\n' +
    '• META_AD_ACCOUNT_ID\n' +
    '• SHOPIFY_STORE_URL\n' +
    '• SHOPIFY_ACCESS_TOKEN\n\n' +
    'Then run: 🚀 D2C Intelligence → Setup Daily Auto-Sync Trigger'
  );
}

// ─── MANUAL TRIGGER: Run sync for a specific date ────────
function manualSyncForDate() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Manual Sync',
    'Enter date to sync (format: YYYY-MM-DD, e.g. 2024-09-15):',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  var dateStr = result.getResponseText().trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    ui.alert('Invalid date format. Use YYYY-MM-DD');
    return;
  }

  // Temporarily override yesterday function
  var origGetYesterdayDate = getYesterdayDate;
  // Note: can't override functions in GAS, so we pass date as global
  PropertiesService.getScriptProperties().setProperty('SYNC_OVERRIDE_DATE', dateStr);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = loadSyncConfig(ss);
  if (!cfg) { ui.alert('CONFIG sheet API keys not set.'); return; }

  try {
    var metaRows = fetchMetaAdsDataForDate(cfg, dateStr);
    if (metaRows.length > 0) appendToSheet(ss, 'META_DATA', metaRows);
    var shopifyRows = fetchShopifyOrdersForDate(cfg, dateStr);
    if (shopifyRows.length > 0) appendToSheet(ss, 'SALES_DATA', shopifyRows);
    computeAndWriteDailyStats(ss);
    ui.alert('✅ Sync complete for ' + dateStr + '!\nMeta rows: ' + metaRows.length + '\nShopify orders: ' + shopifyRows.length);
  } catch(e) {
    ui.alert('❌ Sync error: ' + e.message);
  }

  PropertiesService.getScriptProperties().deleteProperty('SYNC_OVERRIDE_DATE');
}

function fetchMetaAdsDataForDate(cfg, dateStr) {
  // Same as fetchMetaAdsData but for a specific date
  if (!cfg.metaToken || !cfg.metaAccountId) throw new Error('Meta credentials missing');
  var fields = 'date_start,account_name,campaign_name,adset_name,ad_name,spend,impressions,clicks,actions,action_values,landing_page_views,adset_id,ad_id';
  var url = 'https://graph.facebook.com/v19.0/act_' + cfg.metaAccountId + '/insights'
    + '?level=ad&fields=' + encodeURIComponent(fields)
    + '&time_range={"since":"' + dateStr + '","until":"' + dateStr + '"}'
    + '&limit=500&access_token=' + cfg.metaToken;
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var json = JSON.parse(response.getContentText());
  if (json.error) throw new Error(json.error.message);
  return (json.data || []).map(function(row) {
    return [
      row.date_start, row.account_name || '', row.campaign_name || '',
      row.adset_name || '', row.ad_name || '', 'INR',
      parseFloat(row.spend) || 0, parseInt(row.impressions) || 0,
      parseInt(row.clicks) || 0, parseInt(row.landing_page_views) || 0,
      getActionValue(row.actions, 'add_to_cart'),
      getActionValue(row.actions, 'purchase'),
      getActionValue(row.action_values, 'purchase'),
      0, 0, 0, 0, 0, 0, '', row.adset_id || '', row.ad_id || '',
      parseProductFromName(row.campaign_name), '', '', '', '', ''
    ];
  });
}

function fetchShopifyOrdersForDate(cfg, dateStr) {
  if (!cfg.shopifyStore || !cfg.shopifyToken) throw new Error('Shopify credentials missing');
  var nextDay = new Date(dateStr);
  nextDay.setDate(nextDay.getDate() + 1);
  var nextDayStr = Utilities.formatDate(nextDay, 'Asia/Kolkata', 'yyyy-MM-dd');
  var url = 'https://' + cfg.shopifyStore + '/admin/api/2024-01/orders.json'
    + '?status=any&created_at_min=' + dateStr + 'T00:00:00+05:30'
    + '&created_at_max=' + nextDayStr + 'T00:00:00+05:30'
    + '&limit=250&fields=id,email,created_at,total_price,line_items';
  var response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: { 'X-Shopify-Access-Token': cfg.shopifyToken },
    muteHttpExceptions: true
  });
  var json = JSON.parse(response.getContentText());
  if (json.errors) throw new Error(JSON.stringify(json.errors));
  return (json.orders || []).map(function(order) {
    return [
      order.id, order.email || '', order.created_at.substring(0, 10),
      parseFloat(order.total_price) || 0,
      order.line_items && order.line_items[0] ? order.line_items[0].title : 'Unknown',
      'Shopify', dateStr.substring(0, 7)
    ];
  });
}
