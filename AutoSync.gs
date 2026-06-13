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

    var info = parseNomenclatureMeta(row.campaign_name || '', row.adset_name || '', row.ad_name || '');

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
      info.product,       // Product
      info.creatorType,   // Ad Type
      info.format,        // Type of Ad
      info.influencer,    // Influencer Name
      info.narrative,     // Narrative
      info.language       // Language
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

// ─── HELPER: PARSE METADATA FROM CAMPAIGN/ADSET/AD NAME ────
function parseNomenclatureMeta(campaign, adset, adname) {
  // Try structured nomenclature in Ad Name first, then Ad Set, then Campaign Name
  var parsed = parseStructuredNomenclature(adname);
  if (!parsed) parsed = parseStructuredNomenclature(adset);
  if (!parsed) parsed = parseStructuredNomenclature(campaign);

  if (parsed) {
    return {
      product: parsed.product,
      creatorType: parsed.adType, // "Internal UGC" or "Influencer"
      format: parsed.format, // "Reel", "Static", "Carousel"
      influencer: parsed.adType === 'Influencer' ? parsed.cleanName : '',
      narrative: parsed.narrative,
      language: parsed.language,
      cleanAdName: parsed.cleanName
    };
  }

  // Fallback: keyword search
  var combined = (campaign + ' | ' + adset + ' | ' + adname).toLowerCase();
  
  // Product mapping
  var product = 'Other Product';
  if (combined.indexOf('serum') !== -1) product = 'Hair Growth Serum';
  else if (combined.indexOf('bodywash') !== -1 || combined.indexOf('body wash') !== -1) product = '5% AHA BHA Exfoliating Body Wash';
  else if (combined.indexOf('gummies') !== -1) product = 'Biotin Hair Gummies';
  else if (combined.indexOf('underarm') !== -1) product = 'Underarm Roll On';

  // Format / Type of Ad mapping
  var format = 'Video Ad';
  if (combined.indexOf('static') !== -1 || combined.indexOf('image') !== -1) format = 'Static Creative';
  else if (combined.indexOf('carousel') !== -1) format = 'Carousel';
  else if (combined.indexOf('reel') !== -1) format = 'Reel';

  // Creator type
  var creatorType = 'Internal UGC';
  if (combined.indexOf('influencer') !== -1 || combined.indexOf('infl') !== -1 || combined.indexOf('collab') !== -1) {
    creatorType = 'Influencer';
  }

  // Narrative
  var narrative = 'Brand Story';
  if (combined.indexOf('discount') !== -1 || combined.indexOf('sale') !== -1 || combined.indexOf('offer') !== -1) {
    narrative = 'Offer-led';
  } else if (combined.indexOf('benefit') !== -1 || combined.indexOf('quality') !== -1) {
    narrative = 'Feature-led';
  } else if (combined.indexOf('review') !== -1 || combined.indexOf('trust') !== -1) {
    narrative = 'Social Proof/UGC';
  }

  // Language
  var language = 'English';
  if (combined.indexOf('hinglish') !== -1) language = 'Hinglish';
  else if (combined.indexOf('hindi') !== -1) language = 'Hindi';

  // Influencer name if influencer
  var influencer = '';
  if (creatorType === 'Influencer') {
    var parts = adname.split('|');
    if (parts.length >= 3) {
      influencer = parts[2].trim().replace('UGC_', '');
    } else {
      influencer = 'Influencer Creator';
    }
  }

  return {
    product: product,
    creatorType: creatorType,
    format: format,
    influencer: influencer,
    narrative: narrative,
    language: language,
    cleanAdName: adname || 'Generic Creative'
  };
}

function parseStructuredNomenclature(str) {
  if (!str) return null;
  var parts = str.split('|').map(function(p) { return p.trim(); });
  if (parts.length >= 4) {
    var productRaw = parts[0];
    var narrativeRaw = parts[1];
    var creatorTypeRaw = parts[2];
    var cleanName = parts[3];
    var adTypeRaw = parts[4] || '';
    var languageRaw = parts[5] || '';

    // Standard mappings
    var product = 'Other Product';
    var pL = productRaw.toLowerCase();
    if (pL.indexOf('serum') !== -1) product = 'Hair Growth Serum';
    else if (pL.indexOf('bodywash') !== -1 || pL.indexOf('body wash') !== -1) product = '5% AHA BHA Exfoliating Body Wash';
    else if (pL.indexOf('gummies') !== -1) product = 'Biotin Hair Gummies';
    else if (pL.indexOf('underarm') !== -1) product = 'Underarm Roll On';
    else product = productRaw;

    var narrative = 'Brand Story';
    var nL = narrativeRaw.toLowerCase();
    if (nL.indexOf('offer') !== -1 || nL.indexOf('discount') !== -1 || nL.indexOf('sale') !== -1) narrative = 'Offer-led';
    else if (nL.indexOf('feature') !== -1 || nL.indexOf('benefit') !== -1) narrative = 'Feature-led';
    else if (nL.indexOf('proof') !== -1 || nL.indexOf('review') !== -1 || nL.indexOf('trust') !== -1) narrative = 'Social Proof/UGC';
    else narrative = narrativeRaw;

    var adType = 'Internal UGC';
    var cL = creatorTypeRaw.toLowerCase();
    if (cL.indexOf('influencer') !== -1 || cL.indexOf('creator') !== -1 || cL.indexOf('collab') !== -1) adType = 'Influencer';

    var format = 'Video Ad';
    var fL = (adTypeRaw || cleanName).toLowerCase();
    if (fL.indexOf('reel') !== -1) format = 'Reel';
    else if (fL.indexOf('static') !== -1 || fL.indexOf('image') !== -1) format = 'Static Creative';
    else if (fL.indexOf('carousel') !== -1) format = 'Carousel';

    var language = 'English';
    var lL = (languageRaw || cleanName).toLowerCase();
    if (lL.indexOf('hinglish') !== -1) language = 'Hinglish';
    else if (lL.indexOf('hindi') !== -1) language = 'Hindi';

    return {
      product: product,
      narrative: narrative,
      adType: adType,
      cleanName: cleanName,
      format: format,
      language: language
    };
  }
  return null;
}

// Keep old signature placeholders for backwards compatibility in other spreadsheets
function parseProductFromName(name) {
  var info = parseNomenclatureMeta('', '', name || '');
  return info.product;
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
    var info = parseNomenclatureMeta(row.campaign_name || '', row.adset_name || '', row.ad_name || '');
    return [
      row.date_start, row.account_name || '', row.campaign_name || '',
      row.adset_name || '', row.ad_name || '', 'INR',
      parseFloat(row.spend) || 0, parseInt(row.impressions) || 0,
      parseInt(row.clicks) || 0, parseInt(row.landing_page_views) || 0,
      getActionValue(row.actions, 'add_to_cart'),
      getActionValue(row.actions, 'purchase'),
      getActionValue(row.action_values, 'purchase'),
      0, 0, 0, 0, 0, 0, '', row.adset_id || '', row.ad_id || '',
      info.product, info.creatorType, info.format, info.influencer, info.narrative, info.language
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
