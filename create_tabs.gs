// ============================================================
//  ONE-TIME SETUP: Create all required sheet tabs with headers
//  Paste this in Apps Script Editor → Run createAllTabs()
// ============================================================

function createAllTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Auto-rename the spreadsheet to a premium, easy-to-see name
  try {
    ss.rename('D2C Scaler Command Center');
  } catch (e) {
    Logger.log('Error renaming spreadsheet: ' + e.message);
  }

  // --- GOOGLE_DATA (Performance + Search Keyword level) ---
  createTabWithHeaders(ss, 'GOOGLE_DATA', [
    // Core identifiers
    'Date', 'Campaign', 'Campaign Type', 'Ad Group', 'Ad Name',
    // Search-specific (leave blank for non-Search campaigns)
    'Keyword', 'Match Type', 'Search Term', 'Quality Score', 'Impression Share %',
    // Universal performance metrics
    'Spend', 'Impressions', 'Clicks', 'Landing Page Views',
    'Conversions', 'Revenue', 'Add to Cart',
    'CTR %', 'CPC', 'CPM', 'ROAS', 'CAC',
    'Conv Rate %',
    // Creative tagging (fill for Display/YouTube, leave blank for Search)
    'Narrative', 'Type of Ad (Search/Display/Video/Shopping)',
    'Language of the Ad', 'Name of the Influencer',
    // Mapping tags (auto-parsed from Campaign name via nomenclature)
    'Product Name', 'Category', 'Funnel Stage'
  ]);

  // --- SALES_DATA ---
  createTabWithHeaders(ss, 'SALES_DATA', [
    'Order ID', 'Customer ID', 'Order Date', 'Revenue',
    'Product', 'Platform', 'Cohort Month'
  ]);

  // --- PRODUCT_MASTER ---
  createTabWithHeaders(ss, 'PRODUCT_MASTER', [
    'Product', 'Price', 'Cost of Goods', 'Margin', 'Target CPA'
  ]);

  // --- DAILY_STATS ---
  createTabWithHeaders(ss, 'DAILY_STATS', [
    'Date', 'Spend', 'Conversions', 'CAC', 'AOV', 'ROAS'
  ]);

  // --- CONFIG ---
  var cfg = createTabWithHeaders(ss, 'CONFIG', [
    'Metric', 'Target', 'Scale Threshold', 'Stop Threshold', 'Unit'
  ]);
  if (cfg && cfg.getLastRow() <= 1) {
    cfg.getRange('A2:E11').setValues([
      ['ROAS',       3.0, 4.0, 1.5, 'x'],
      ['CTR',        1.5, 2.5, 0.5, '%'],
      ['CPC',        50,  30,  100, '₹'],
      ['Conv Rate',  2.0, 3.5, 0.5, '%'],
      ['CPM',        200, 150, 500, '₹'],
      ['Frequency',  2.5, 1.5, 4.0, 'x'],
      ['Min Spend',  200, 500, 0,   '₹'],
      ['Hook Rate',  25,  35,  10,  '%'],
      ['Hold Rate',  20,  30,  8,   '%'],
      ['ATC Rate',   5,   8,   1,   '%']
    ]);
  }

  SpreadsheetApp.getUi().alert(
    '✅ Done! Created tabs:\n• GOOGLE_DATA\n• SALES_DATA\n• PRODUCT_MASTER\n• DAILY_STATS\n• CONFIG\n\nYour META_DATA tab is untouched. Refresh the dashboard to sync.'
  );
}

function createTabWithHeaders(ss, name, headers) {
  var existing = ss.getSheetByName(name);
  if (existing) {
    Logger.log('Tab already exists: ' + name);
    return existing;
  }
  var sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  // Bold + freeze the header row
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  Logger.log('Created tab: ' + name);
  return sheet;
}
