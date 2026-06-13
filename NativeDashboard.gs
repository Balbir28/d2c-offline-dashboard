// ============================================================
//  NATIVE SPREADSHEET DASHBOARD BUILDER — NativeDashboard.gs
//  Generates a high-performance native dashboard sheet tab
// ============================================================

function createOrGetSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.clear();
    // Reset formatting
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .setBackground('#ffffff')
      .setFontColor('#000000')
      .setFontWeight('normal')
      .setFontStyle('normal')
      .setFontSize(10)
      .setHorizontalAlignment('left')
      .setBorder(false, false, false, false, false, false);
  } else {
    sheet = ss.insertSheet(name);
  }
  sheet.setHiddenGridlines(true);
  return sheet;
}

function setupNativeDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = createOrGetSheet(ss, 'DASHBOARD');
  
  // Set sheet columns width
  sheet.setColumnWidth(1, 180); // Metric / Cohort Label
  sheet.setColumnWidth(2, 110); // Value
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 110);
  sheet.setColumnWidth(9, 110);
  sheet.setColumnWidth(10, 110);
  sheet.setColumnWidth(11, 40);  // Divider column
  sheet.setColumnWidth(12, 280); // AI Signals / Fatigue
  sheet.setColumnWidth(13, 220); // Action Item
  sheet.setColumnWidth(14, 150); // Additional Context

  // 1. Banner Header
  sheet.getRange('A1:N2').merge();
  var header = sheet.getRange('A1');
  header.setValue('🚀 D2C MARKETING INTELLIGENCE COMMAND CENTER')
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#0f172a')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // 2. Date Filter Controls
  sheet.getRange('A4').setValue('📅 Date Filters:').setFontWeight('bold').setFontSize(11);
  
  sheet.getRange('B4').setValue('Start Date:').setHorizontalAlignment('right').setFontColor('#64748b').setFontWeight('bold');
  var startDateCell = sheet.getRange('C4');
  startDateCell.setValue(getDefaultStartDateJS()).setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('center');
  
  sheet.getRange('D4').setValue('End Date:').setHorizontalAlignment('right').setFontColor('#64748b').setFontWeight('bold');
  var endDateCell = sheet.getRange('E4');
  endDateCell.setValue(new Date()).setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('center');
  
  sheet.getRange('F4').setValue('🔄 Refresh AI Metrics')
    .setBackground('#2563eb')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Format filter row
  sheet.getRange('A4:F4').setVerticalAlignment('middle');
  sheet.getRange('C4:E4').setBorder(true, true, true, true, false, false, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  
  // Helper functions for dates
  function getDefaultStartDateJS() {
    var d = new Date();
    d.setDate(d.getDate() - 30);
    return d;
  }

  // 3. KPI block layout function
  function formatKPICard(rangeCell, label, formula, isCurrency, isPercentage, isROAS, bg) {
    var parts = rangeCell.split(':');
    var startCell = parts[0];
    var endCell = parts[1];
    
    var startCol = startCell.charCodeAt(0) - 64;
    var startRow = parseInt(startCell.slice(1), 10);
    var endCol = endCell.charCodeAt(0) - 64;
    var endRow = parseInt(endCell.slice(1), 10);
    
    // Label Row: merge top row
    var labelRange = sheet.getRange(startRow, startCol, 1, (endCol - startCol + 1));
    labelRange.merge()
      .setValue(label)
      .setFontSize(9)
      .setFontColor('#64748b')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
      
    // Value Row: merge bottom rows
    var valueRange = sheet.getRange(startRow + 1, startCol, (endRow - startRow), (endCol - startCol + 1));
    valueRange.merge()
      .setFormula(formula)
      .setFontSize(16)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
      
    if (isCurrency) {
      valueRange.setNumberFormat('₹#,##0');
      valueRange.setFontColor('#1e293b');
    } else if (isPercentage) {
      valueRange.setNumberFormat('0.00%');
      valueRange.setFontColor('#2563eb');
    } else if (isROAS) {
      valueRange.setNumberFormat('0.00"x"');
      valueRange.setFontColor('#15803d');
    } else {
      valueRange.setNumberFormat('#,##0');
      valueRange.setFontColor('#1e293b');
    }

    // Border and Background
    var blockRange = sheet.getRange(startRow, startCol, (endRow - startRow + 1), (endCol - startCol + 1));
    blockRange.setBackground(bg)
      .setBorder(true, true, true, true, false, false, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  }

  // Build KPI formulas
  var condStart = '">="&C4';
  var condEnd = '"<="&E4';
  
  var formulaSpend = `=SUMIFS(META_DATA!G:G, META_DATA!A:A, ${condStart}, META_DATA!A:A, ${condEnd}) + SUMIFS(GOOGLE_DATA!G:G, GOOGLE_DATA!A:A, ${condStart}, GOOGLE_DATA!A:A, ${condEnd})`;
  var formulaRevenue = `=SUMIFS(META_DATA!M:M, META_DATA!A:A, ${condStart}, META_DATA!A:A, ${condEnd}) + SUMIFS(GOOGLE_DATA!M:M, GOOGLE_DATA!A:A, ${condStart}, GOOGLE_DATA!A:A, ${condEnd})`;
  var formulaConversions = `=SUMIFS(META_DATA!L:L, META_DATA!A:A, ${condStart}, META_DATA!A:A, ${condEnd}) + SUMIFS(GOOGLE_DATA!L:L, GOOGLE_DATA!A:A, ${condStart}, GOOGLE_DATA!A:A, ${condEnd})`;

  // Add KPI Cards (4 cards matching cols A-J)
  formatKPICard('A6:B8', '💵 TOTAL SPEND', formulaSpend, true, false, false, '#f8fafc');
  formatKPICard('C6:E8', '💰 TOTAL REVENUE', formulaRevenue, true, false, false, '#f8fafc');
  formatKPICard('F6:H8', '🎯 BLENDED ROAS', `=IF(A7>0, C7/A7, 0)`, false, false, true, '#f0fdf4');
  formatKPICard('I6:J8', '✅ CONVERSIONS', formulaConversions, false, false, false, '#f8fafc');

  // 4. Section Titles
  function setSectionTitle(range, text) {
    sheet.getRange(range).merge()
      .setValue(text)
      .setFontSize(13)
      .setFontWeight('bold')
      .setFontColor('#0f172a')
      .setBackground('#f1f5f9')
      .setVerticalAlignment('middle')
      .setBorder(false, false, true, false, false, false, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  }

  setSectionTitle('A10:J10', '👥 Customer Retention MoM & LTV Cohorts');
  setSectionTitle('L10:N10', '⚡ Real-Time AI Signals');
  setSectionTitle('L22:N22', '💰 AI Budget Reallocation Recommendations');
  setSectionTitle('A24:J24', '📊 Ad Creative Fatigue & Wear-Out Monitor');

  // 5. Setup empty cohort tables structure
  var cohortHeaders = [
    'Cohort Month', 'New Customers', 'LTV Month 0', 'LTV Month 1', 'LTV Month 2', 'LTV Month 3', 'Ret. Month 0', 'Ret. Month 1', 'Ret. Month 2', 'Ret. Month 3'
  ];
  sheet.getRange('A11:J11').setValues([cohortHeaders])
    .setFontColor('#ffffff')
    .setBackground('#1e293b')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 6. Setup empty AI Signals list
  sheet.getRange('L11').setValue('Loading AI Recommendations... Click menu 🚀 D2C Intelligence -> Refresh Native Dashboard.')
    .setFontStyle('italic')
    .setFontColor('#94a3b8');

  // 7. Setup empty Budget recommendations structure
  var budgetHeaders = ['Campaign', 'Platform', 'Signal', 'Spend', 'Suggestion', 'Action Item'];
  sheet.getRange('L23:Q23').setValues([budgetHeaders])
    .setFontColor('#ffffff')
    .setBackground('#1e293b')
    .setFontWeight('bold');

  // 8. Setup empty fatigue table structure
  var fatigueHeaders = [
    'Ad Name', 'Ad Type', 'Platform', 'Frequency', 'Baseline CTR', 'Recent CTR (7D)', 'Performance Drop', 'Wear-out Status', 'AI Recommendation'
  ];
  sheet.getRange('A25:I25').setValues([fatigueHeaders])
    .setFontColor('#ffffff')
    .setBackground('#1e293b')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Set grid style on main layouts
  sheet.getRange('A1:N60').setVerticalAlignment('middle');

  // Initialize and run the first Refresh
  refreshNativeIntelligence();
}

function refreshNativeIntelligence() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('DASHBOARD');
  if (!sheet) {
    return;
  }

  // Clear previous values
  // Clear Cohort data rows (rows 12-22)
  sheet.getRange('A12:J22').clearContent().setBackground('#ffffff');
  // Clear AI Signals (rows 11-21, cols L-N)
  sheet.getRange('L11:N21').clearContent().setBackground('#ffffff');
  // Clear Budget recommendations (rows 24-35, cols L-Q)
  sheet.getRange('L24:Q35').clearContent().setBackground('#ffffff');
  // Clear Fatigue data (rows 26-60, cols A-I)
  sheet.getRange('A26:I60').clearContent().setBackground('#ffffff');
  // Clear Meta Performance tables (rows 62-600)
  sheet.getRange('A62:O600').clearContent().setBackground('#ffffff').setFontWeight('normal').setFontStyle('normal').setFontColor('#000000').setHorizontalAlignment('left').setBorder(false, false, false, false, false, false);

  // Fetch filter dates
  var startDate = sheet.getRange('C4').getValue();
  var endDate = sheet.getRange('E4').getValue();
  
  var filters = {
    startDate: startDate,
    endDate: endDate,
    platform: 'all',
    product: 'all',
    adType: 'all'
  };

  // Run backend data aggregation
  var data = getDashboardData(filters);

  // 1. Fill AI Signals
  var ins = data.intelligence;
  if (!ins || !ins.length) {
    sheet.getRange('L11').setValue('No AI insights found for the date range. Check your data sheets.');
  } else {
    for (var i = 0; i < Math.min(10, ins.length); i++) {
      var rowIdx = 11 + i;
      sheet.getRange(rowIdx, 12).setValue(ins[i].icon + ' ' + ins[i].title).setFontWeight('bold');
      sheet.getRange(rowIdx, 13).setValue(ins[i].message).setFontSize(9).setFontColor('#475569');
      sheet.getRange(rowIdx, 14).setValue(ins[i].action).setFontSize(9).setFontColor('#2563eb');
    }
  }

  // 2. Fill Cohort Analysis Heatmap
  var cohorts = data.cohortData;
  if (!cohorts || !cohorts.length) {
    sheet.getRange('A12').setValue('No Cohorts data found. Generate mock transaction data first.');
  } else {
    var cohortRows = [];
    cohorts.forEach(function(c) {
      cohortRows.push([
        c.cohort,
        c.size,
        c.ltv[0], c.ltv[1], c.ltv[2], c.ltv[3],
        c.retention[0] / 100, c.retention[1] / 100, c.retention[2] / 100, c.retention[3] / 100
      ]);
    });
    
    // Write cohort data
    var cohortRange = sheet.getRange(12, 1, cohortRows.length, 10);
    cohortRange.setValues(cohortRows);
    sheet.getRange(12, 1, cohortRows.length, 1).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(12, 2, cohortRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
    sheet.getRange(12, 3, cohortRows.length, 4).setNumberFormat('₹#,##0');
    sheet.getRange(12, 7, cohortRows.length, 4).setNumberFormat('0.0%').setHorizontalAlignment('center');

    // Apply Heatmap conditional formatting on Retention columns (cols 7-10)
    var retentionRange = sheet.getRange(12, 7, cohortRows.length, 4);
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpoint('#fee2e2')
      .setGradientMidpointWithValue('#fef08a', SpreadsheetApp.InterpolationType.PERCENTILE, '50')
      .setGradientMaxpoint('#15803d')
      .setRanges([retentionRange])
      .build();
    var rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
  }

  // 3. Fill Budget Recommendations
  var bo = data.budgetRecommendations;
  if (!bo || !bo.actions || !bo.actions.length) {
    sheet.getRange('L24').setValue('All active campaigns are currently pacing within efficiency thresholds.');
  } else {
    var boRows = [];
    bo.actions.forEach(function(a) {
      boRows.push([
        a.campaign,
        a.platform,
        a.type,
        a.currentSpend,
        a.recommendation,
        a.actionItem
      ]);
    });
    sheet.getRange(24, 12, boRows.length, 6).setValues(boRows);
    
    // Formatting signals
    for (var i = 0; i < boRows.length; i++) {
      var r = 24 + i;
      var signalCell = sheet.getRange(r, 14);
      if (boRows[i][2] === 'STOP') {
        signalCell.setBackground('#fee2e2').setFontColor('#ef4444').setFontWeight('bold');
      } else {
        signalCell.setBackground('#dcfce7').setFontColor('#22c55e').setFontWeight('bold');
      }
    }
  }

  // 4. Fill Ad Fatigue Monitor
  var fatigue = data.fatigueData;
  if (!fatigue || !fatigue.length) {
    sheet.getRange('A26').setValue('Fatigue tracker requires at least 12 days of ad impressions to render baselines.');
  } else {
    var fatRows = [];
    fatigue.forEach(function(ad) {
      fatRows.push([
        ad.adName,
        ad.adType,
        ad.platform,
        ad.frequency,
        ad.baselineCtr / 100,
        ad.recentCtr / 100,
        ad.ctrDrop / 100,
        ad.status,
        ad.recommendation
      ]);
    });
    
    var fatRange = sheet.getRange(26, 1, fatRows.length, 9);
    fatRange.setValues(fatRows);
    sheet.getRange(26, 1, fatRows.length, 1).setFontWeight('bold');
    sheet.getRange(26, 4, fatRows.length, 1).setNumberFormat('0.0"x"').setHorizontalAlignment('center');
    sheet.getRange(26, 5, fatRows.length, 3).setNumberFormat('0.00%').setHorizontalAlignment('center');
    sheet.getRange(26, 8, fatRows.length, 1).setHorizontalAlignment('center');

    // Format Fatigue status
    for (var i = 0; i < fatRows.length; i++) {
      var r = 26 + i;
      var statusCell = sheet.getRange(r, 8);
      var status = fatRows[i][7];
      if (status === 'HIGH FATIGUE') {
        statusCell.setBackground('#fee2e2').setFontColor('#ef4444').setFontWeight('bold');
      } else if (status === 'MODERATE FATIGUE') {
        statusCell.setBackground('#fef9c3').setFontColor('#ca8a04').setFontWeight('bold');
      } else {
        statusCell.setBackground('#dcfce7').setFontColor('#22c55e').setFontWeight('bold');
      }
    }
  }

  // 5. Create Native Charts
  createOrUpdateNativeCharts(sheet);
  
  // 6. Fill Meta Ads Performance (Ad Sets & Ad Name Creative level)
  renderMetaPerformance(sheet, data);
}

function createOrUpdateNativeCharts(sheet) {
  var charts = sheet.getCharts();
  if (charts.length > 0) {
    // Skip recreation to avoid Google Sheets chart rendering lag on subsequent refreshes
    return;
  }
  
  // Spend vs Revenue Line Chart (Daily stats from metadata sheet)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var metaSheet = ss.getSheetByName(META_SHEET);
  if (!metaSheet || metaSheet.getLastRow() <= 1) return;

  var lastRow = Math.min(60, metaSheet.getLastRow()); // limit to recent 60 rows for visualization
  
  // Chart 1: Spend vs Revenue trend
  var spendRange = metaSheet.getRange(1, 7, lastRow, 1);  // Amount spent (INR) (Col G)
  var revRange   = metaSheet.getRange(1, 13, lastRow, 1); // Purchases conversion value (Col M)
  var dateRange  = metaSheet.getRange(1, 1, lastRow, 1);  // Day (Col A)

  var chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dateRange)
    .addRange(spendRange)
    .addRange(revRange)
    .setPosition(48, 1, 0, 0)
    .setOption('title', '📈 Revenue vs Spend Trend (Daily)')
    .setOption('hAxis', {title: 'Date', textStyle: {fontSize: 9}})
    .setOption('vAxis', {title: 'Amount (₹)'})
    .setOption('width', 900)
    .setOption('height', 240)
    .setOption('colors', ['#3b82f6', '#10b981'])
    .setOption('curveType', 'function');
    
  sheet.insertChart(chartBuilder.build());
}

function renderMetaPerformance(sheet, data) {
  var startRow = 62;
  
  // Section Header
  sheet.getRange(startRow, 1, 1, 15).merge()
    .setValue('🎯 Meta Ads Performance Optimizer (Ad Sets & Creative Level)')
    .setFontSize(13)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1e293b')
    .setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
    
  // 1. Meta Ad Sets Table
  var adSetRow = startRow + 2;
  var adSetHeaders = [
    'Meta Ad Set Name', 'Spend (3d)', 'Spend (7d)', 'Spend (10d)', 'ROAS (3d)', 'ROAS (7d)', 'ROAS (10d)', 'Trend', 'Recommendation'
  ];
  sheet.getRange(adSetRow, 1, 1, 9).setValues([adSetHeaders])
    .setFontColor('#ffffff')
    .setBackground('#334155')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
    
  var metaAdSets = (data.optimizerAdSets || []).filter(function(as) { return as.platform === 'Meta'; });
  var adSetRows = [];
  metaAdSets.forEach(function(as) {
    adSetRows.push([
      as.name,
      as.spend[0], as.spend[1], as.spend[2],
      as.roas[0], as.roas[1], as.roas[2],
      as.trend,
      as.recommendation
    ]);
  });
  
  if (adSetRows.length > 0) {
    sheet.getRange(adSetRow + 1, 1, adSetRows.length, 9).setValues(adSetRows);
    sheet.getRange(adSetRow + 1, 1, adSetRows.length, 1).setFontWeight('bold');
    sheet.getRange(adSetRow + 1, 2, adSetRows.length, 3).setNumberFormat('₹#,##0').setHorizontalAlignment('right');
    sheet.getRange(adSetRow + 1, 5, adSetRows.length, 3).setNumberFormat('0.00"x"').setHorizontalAlignment('right');
    sheet.getRange(adSetRow + 1, 8, adSetRows.length, 2).setHorizontalAlignment('center');
    
    // Format Recommendation badges
    for (var i = 0; i < adSetRows.length; i++) {
      var r = adSetRow + 1 + i;
      var recCell = sheet.getRange(r, 9);
      var rec = adSetRows[i][8];
      if (rec.indexOf('Scale') !== -1) {
        recCell.setBackground('#dcfce7').setFontColor('#15803d').setFontWeight('bold');
      } else if (rec.indexOf('Kill') !== -1 || rec.indexOf('Soft') !== -1) {
        recCell.setBackground('#fee2e2').setFontColor('#b91c1c').setFontWeight('bold');
      } else if (rec.indexOf('Hold') !== -1) {
        recCell.setBackground('#fef9c3').setFontColor('#a16207').setFontWeight('bold');
      }
    }
  } else {
    sheet.getRange(adSetRow + 1, 1).setValue('No Meta Ad Set data available. Check META_DATA.').setFontStyle('italic').setFontColor('#94a3b8');
  }
  
  // 2. Meta Ad Creatives Table
  var adRow = adSetRow + Math.max(1, adSetRows.length) + 3;
  var adHeaders = [
    'Meta Ad Name', 'Product', 'Ad Type (Content)', 'Type of Ad (Format)', 'Influencer Name',
    'Spend (3d)', 'Spend (7d)', 'Spend (10d)', 'ROAS (3d)', 'ROAS (7d)', 'ROAS (10d)',
    'Hook Rate', 'Hold Rate', 'Trend', 'Recommendation'
  ];
  sheet.getRange(adRow, 1, 1, 15).setValues([adHeaders])
    .setFontColor('#ffffff')
    .setBackground('#334155')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
    
  var metaAds = (data.optimizerAds || []).filter(function(ad) { return ad.platform === 'Meta'; });
  var adRows = [];
  metaAds.forEach(function(ad) {
    adRows.push([
      ad.name,
      ad.product,
      ad.adType,            // internal/influencer
      ad.typeOfAd,          // static/reel
      ad.influencerName,
      ad.spend[0], ad.spend[1], ad.spend[2],
      ad.roas[0], ad.roas[1], ad.roas[2],
      ad.hook / 100, ad.hold / 100,
      ad.trend,
      ad.recommendation
    ]);
  });
  
  if (adRows.length > 0) {
    sheet.getRange(adRow + 1, 1, adRows.length, 15).setValues(adRows);
    sheet.getRange(adRow + 1, 1, adRows.length, 1).setFontWeight('bold');
    sheet.getRange(adRow + 1, 2, adRows.length, 4).setHorizontalAlignment('left');
    sheet.getRange(adRow + 1, 6, adRows.length, 3).setNumberFormat('₹#,##0').setHorizontalAlignment('right');
    sheet.getRange(adRow + 1, 9, adRows.length, 3).setNumberFormat('0.00"x"').setHorizontalAlignment('right');
    sheet.getRange(adRow + 1, 12, adRows.length, 2).setNumberFormat('0.0%').setHorizontalAlignment('center');
    sheet.getRange(adRow + 1, 14, adRows.length, 2).setHorizontalAlignment('center');
    
    // Format Recommendation badges
    for (var i = 0; i < adRows.length; i++) {
      var r = adRow + 1 + i;
      var recCell = sheet.getRange(r, 15);
      var rec = adRows[i][14];
      if (rec.indexOf('Scale') !== -1) {
        recCell.setBackground('#dcfce7').setFontColor('#15803d').setFontWeight('bold');
      } else if (rec.indexOf('Kill') !== -1 || rec.indexOf('Soft') !== -1) {
        recCell.setBackground('#fee2e2').setFontColor('#b91c1c').setFontWeight('bold');
      } else if (rec.indexOf('Hold') !== -1) {
        recCell.setBackground('#fef9c3').setFontColor('#a16207').setFontWeight('bold');
      } else if (rec.indexOf('Fix') !== -1) {
        recCell.setBackground('#ffe4e6').setFontColor('#db2777').setFontWeight('bold');
      }
    }
  } else {
    sheet.getRange(adRow + 1, 1).setValue('No Meta Ad data available. Check META_DATA.').setFontStyle('italic').setFontColor('#94a3b8');
  }

  // 3. Top 5 vs Bottom 5 Performing Ad Sets Table
  var nextSectionRow = adRow + Math.max(1, adRows.length) + 3;
  var topRow = nextSectionRow + 2;
  sheet.getRange(topRow - 2, 1, 1, 13).merge()
    .setValue('🏆 Top 5 vs 🛑 Bottom 5 Performing Ad Sets (by ROAS, Spend > ₹500)')
    .setFontSize(12)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#0f172a')
    .setVerticalAlignment('middle');

  sheet.getRange(topRow - 1, 1, 1, 6).merge()
    .setValue('Top 5 Performing Ad Sets')
    .setFontWeight('bold')
    .setFontColor('#15803d')
    .setBackground('#f0fdf4')
    .setHorizontalAlignment('center');

  sheet.getRange(topRow - 1, 8, 1, 6).merge()
    .setValue('Bottom 5 Performing Ad Sets')
    .setFontWeight('bold')
    .setFontColor('#b91c1c')
    .setBackground('#fee2e2')
    .setHorizontalAlignment('center');

  var adSetSubHeaders = ['Platform', 'Ad Set Name', 'Spend', 'Revenue', 'ROAS', 'Conversions'];
  sheet.getRange(topRow, 1, 1, 6).setValues([adSetSubHeaders]).setFontColor('#ffffff').setBackground('#334155').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(topRow, 8, 1, 6).setValues([adSetSubHeaders]).setFontColor('#ffffff').setBackground('#334155').setFontWeight('bold').setHorizontalAlignment('center');

  var topAdSets = data.topAdSets || [];
  var bottomAdSets = data.bottomAdSets || [];

  for (var i = 0; i < 5; i++) {
    var r = topRow + 1 + i;
    // Top ad set
    if (i < topAdSets.length) {
      var item = topAdSets[i];
      sheet.getRange(r, 1, 1, 6).setValues([[
        item.platform, item.adSetName, item.spend, item.revenue, item.roas, item.conversions
      ]]);
      sheet.getRange(r, 2).setFontWeight('bold');
      sheet.getRange(r, 3, 1, 2).setNumberFormat('₹#,##0').setHorizontalAlignment('right');
      sheet.getRange(r, 5).setNumberFormat('0.00"x"').setHorizontalAlignment('right');
      sheet.getRange(r, 6).setNumberFormat('#,##0').setHorizontalAlignment('center');
    } else {
      sheet.getRange(r, 1, 1, 6).merge().setValue('-').setHorizontalAlignment('center').setFontColor('#94a3b8');
    }

    // Bottom ad set
    if (i < bottomAdSets.length) {
      var item = bottomAdSets[i];
      sheet.getRange(r, 8, 1, 6).setValues([[
        item.platform, item.adSetName, item.spend, item.revenue, item.roas, item.conversions
      ]]);
      sheet.getRange(r, 9).setFontWeight('bold');
      sheet.getRange(r, 10, 1, 2).setNumberFormat('₹#,##0').setHorizontalAlignment('right');
      sheet.getRange(r, 12).setNumberFormat('0.00"x"').setHorizontalAlignment('right');
      sheet.getRange(r, 13).setNumberFormat('#,##0').setHorizontalAlignment('center');
    } else {
      sheet.getRange(r, 8, 1, 6).merge().setValue('-').setHorizontalAlignment('center').setFontColor('#94a3b8');
    }
  }

  // 4. Creative Messaging & Language Optimizer Table
  var creativeIntelRow = topRow + 8;
  sheet.getRange(creativeIntelRow, 1, 1, 13).merge()
    .setValue('🧠 Creative Messaging & Language Optimizer')
    .setFontSize(12)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#0f172a')
    .setVerticalAlignment('middle');

  sheet.getRange(creativeIntelRow + 1, 1, 1, 6).merge()
    .setValue('Creative Narrative / Angle Performance')
    .setFontWeight('bold')
    .setFontColor('#0f172a')
    .setBackground('#f1f5f9')
    .setHorizontalAlignment('center');

  sheet.getRange(creativeIntelRow + 1, 8, 1, 6).merge()
    .setValue('Language Performance')
    .setFontWeight('bold')
    .setFontColor('#0f172a')
    .setBackground('#f1f5f9')
    .setHorizontalAlignment('center');

  var narrativeHeaders = ['Narrative Angle', 'Spend', 'Revenue', 'ROAS', 'Conversions', 'Status'];
  var languageHeaders = ['Language', 'Spend', 'Revenue', 'ROAS', 'Conversions', 'Status'];
  sheet.getRange(creativeIntelRow + 2, 1, 1, 6).setValues([narrativeHeaders]).setFontColor('#ffffff').setBackground('#334155').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(creativeIntelRow + 2, 8, 1, 6).setValues([languageHeaders]).setFontColor('#ffffff').setBackground('#334155').setFontWeight('bold').setHorizontalAlignment('center');

  var narrativeStats = data.narrativeStats || [];
  var languageStats = data.languageStats || [];
  var maxCreativeLen = Math.max(narrativeStats.length, languageStats.length, 1);

  for (var i = 0; i < maxCreativeLen; i++) {
    var r = creativeIntelRow + 3 + i;
    // Narrative
    if (i < narrativeStats.length) {
      var item = narrativeStats[i];
      sheet.getRange(r, 1, 1, 6).setValues([[
        item.narrative, item.spend, item.revenue, item.roas, item.conversions, item.status
      ]]);
      sheet.getRange(r, 1).setFontWeight('bold');
      sheet.getRange(r, 2, 1, 2).setNumberFormat('₹#,##0').setHorizontalAlignment('right');
      sheet.getRange(r, 4).setNumberFormat('0.00"x"').setHorizontalAlignment('right');
      sheet.getRange(r, 5).setNumberFormat('#,##0').setHorizontalAlignment('center');
      sheet.getRange(r, 6).setHorizontalAlignment('center');
      if (item.status.indexOf('Working ✅') !== -1) {
        sheet.getRange(r, 6).setBackground('#dcfce7').setFontColor('#15803d').setFontWeight('bold');
      } else if (item.status.indexOf('Not Working ❌') !== -1) {
        sheet.getRange(r, 6).setBackground('#fee2e2').setFontColor('#b91c1c').setFontWeight('bold');
      }
    } else {
      sheet.getRange(r, 1, 1, 6).merge().setValue('-').setHorizontalAlignment('center').setFontColor('#94a3b8');
    }

    // Language
    if (i < languageStats.length) {
      var item = languageStats[i];
      sheet.getRange(r, 8, 1, 6).setValues([[
        item.language, item.spend, item.revenue, item.roas, item.conversions, item.status
      ]]);
      sheet.getRange(r, 8).setFontWeight('bold');
      sheet.getRange(r, 9, 1, 2).setNumberFormat('₹#,##0').setHorizontalAlignment('right');
      sheet.getRange(r, 11).setNumberFormat('0.00"x"').setHorizontalAlignment('right');
      sheet.getRange(r, 12).setNumberFormat('#,##0').setHorizontalAlignment('center');
      sheet.getRange(r, 13).setHorizontalAlignment('center');
      if (item.status.indexOf('Working ✅') !== -1) {
        sheet.getRange(r, 13).setBackground('#dcfce7').setFontColor('#15803d').setFontWeight('bold');
      } else if (item.status.indexOf('Not Working ❌') !== -1) {
        sheet.getRange(r, 13).setBackground('#fee2e2').setFontColor('#b91c1c').setFontWeight('bold');
      }
    } else {
      sheet.getRange(r, 8, 1, 6).merge().setValue('-').setHorizontalAlignment('center').setFontColor('#94a3b8');
    }
  }

  // 5. Product-wise Daily Performance Trend Table
  var trendRow = creativeIntelRow + 3 + maxCreativeLen + 3;
  sheet.getRange(trendRow, 1, 1, 6).merge()
    .setValue('📈 Product-Wise Daily Performance Timeline Trend')
    .setFontSize(12)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#0f172a')
    .setVerticalAlignment('middle');

  var trendHeaders = ['Date', 'Product SKU', 'Spend', 'Revenue', 'ROAS', 'Conversions'];
  sheet.getRange(trendRow + 1, 1, 1, 6).setValues([trendHeaders]).setFontColor('#ffffff').setBackground('#334155').setFontWeight('bold').setHorizontalAlignment('center');

  var productTrend = data.productDailyTrend || [];
  if (productTrend.length > 0) {
    var trendValues = productTrend.map(function(item) {
      return [
        item.date,
        item.product,
        item.spend,
        item.revenue,
        item.roas,
        item.conversions
      ];
    });
    sheet.getRange(trendRow + 2, 1, trendValues.length, 6).setValues(trendValues);
    sheet.getRange(trendRow + 2, 1, trendValues.length, 1).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(trendRow + 2, 2, trendValues.length, 1).setHorizontalAlignment('left');
    sheet.getRange(trendRow + 2, 3, trendValues.length, 2).setNumberFormat('₹#,##0').setHorizontalAlignment('right');
    sheet.getRange(trendRow + 2, 5, trendValues.length, 1).setNumberFormat('0.00"x"').setHorizontalAlignment('right');
    sheet.getRange(trendRow + 2, 6, trendValues.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
  } else {
    sheet.getRange(trendRow + 2, 1, 1, 6).merge().setValue('No daily product trend data available. Check your inputs.').setFontStyle('italic').setFontColor('#94a3b8');
  }
}
