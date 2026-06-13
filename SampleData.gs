// ============================================================
//  SAMPLE DATA GENERATOR — SampleData.gs
//  90-day realistic D2C data with transaction sales ledger
// ============================================================

function insertSampleData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets();

  var today = new Date();
  var startDate = new Date(today);
  startDate.setDate(today.getDate() - 90);

  var products  = ['Running Shoes X1', 'Sport Tee Pro', 'Recovery Foam Roller', 'Gym Bag Elite', 'Protein Shaker'];
  var campaigns = [
    'Brand Awareness Q2 - Sport Tee Pro', 
    'Performance Scale - Running Shoes X1', 
    'Retargeting Wave - Recovery Foam Roller', 
    'New Launch - Shoes - Running Shoes X1', 
    'Loyalty Winback - Protein Shaker'
  ];
  var adSets    = ['Lookalike 1-3%', 'Interest - Fitness', 'Retarget 30Day', 'Broad 18-35', 'Custom - Buyers'];
  var adTypes   = ['Influencer', 'Internal UGC', 'Static Creative', 'Video Ad', 'Carousel'];

  var adNamesMap = {
    'Influencer':      ['@fit_ananya_review_influencer_Ananya_hindi', '@gymbro_rahul_unbox_influencer_Rahul_local', '@wellness_meera_vid_influencer_Meera_hinglish', '@strongarjun_collab_influencer_Arjun'],
    'Internal UGC':    ['Customer Review Compilation v1_creator_Sarah', 'Before After Transformation_creator_Alex_hindi', '30-Day Challenge Result_creator_Emma_hinglish', 'User Workout Reel_creator_Mike'],
    'Static Creative': ['Product Hero Shot Red_creator_Elena', 'Lifestyle Flat Lay Blue_creator_Ryan_hindi', 'Feature Callout Card_creator_Sophia', 'Discount Banner 20%OFF_creator_Chloe_hinglish'],
    'Video Ad':        ['Brand Story 60s_creator_Marcus', 'Product Demo 15s_creator_Brandon_hindi', 'Testimonial Montage_creator_Ryan', 'Founder Story Cut_creator_Emma'],
    'Carousel':        ['Top 5 Products Carousel_creator_Alex_hindi', 'Collection Showcase_creator_Sarah', 'Before After Swipe_creator_Mike_hinglish', 'Feature Comparison_creator_David']
  };

  // Hook/Hold rates by ad type — video-first formats get high hook/hold
  var hookByType = {
    'Influencer': 38, 'Internal UGC': 32, 'Static Creative': 0,
    'Video Ad': 42, 'Carousel': 18
  };
  var holdByType = {
    'Influencer': 28, 'Internal UGC': 22, 'Static Creative': 0,
    'Video Ad': 35, 'Carousel': 12
  };

  var productPrices = {
    'Running Shoes X1': 3999,
    'Sport Tee Pro': 1299,
    'Recovery Foam Roller': 999,
    'Gym Bag Elite': 1999,
    'Protein Shaker': 499
  };

  var metaSheet   = ss.getSheetByName(META_SHEET);
  var googleSheet = ss.getSheetByName(GOOGLE_SHEET);
  var salesSheet  = ss.getSheetByName(SALES_SHEET);

  // Clear old data rows (keep headers)
  if (metaSheet.getLastRow() > 1)   metaSheet.getRange(2, 1, metaSheet.getLastRow() - 1, Math.max(28, metaSheet.getLastColumn())).clearContent();
  if (googleSheet.getLastRow() > 1) googleSheet.getRange(2, 1, googleSheet.getLastRow() - 1, Math.max(25, googleSheet.getLastColumn())).clearContent();
  if (salesSheet.getLastRow() > 1)   salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, Math.max(7, salesSheet.getLastColumn())).clearContent();

  var metaRows   = [];
  var googleRows = [];
  var salesRows  = [];

  var customerRegistry = [];
  var nextCustId = 10001;
  var nextOrderId = 50001;

  for (var day = 0; day < 90; day++) {
    var d = new Date(startDate);
    d.setDate(startDate.getDate() + day);
    var dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var monthStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM');

    var dow = d.getDay();
    var weekendMult = (dow === 0 || dow === 6) ? 1.3 : 1.0;
    // Simulate creative fatigue building — CTR and Hook degrade over time
    var fatigueFactor = day > 60 ? (1 - ((day - 60) * 0.015)) : 1.0;

    campaigns.forEach(function(camp, ci) {
      adTypes.forEach(function(adType, ti) {
        var adNames = adNamesMap[adType];
        var selectedAd = adNames[ci % adNames.length];
        var product = products[ci % products.length];
        var adSet   = adSets[ti % adSets.length];
        var price   = productPrices[product];

        // Base performance per campaign
        var baseRoas  = [2.1, 4.5, 3.2, 1.3, 2.8][ci];
        var baseSpend = [1200, 3500, 800, 2200, 600][ci] / 30;
        var baseCTR   = [1.2, 2.8, 1.8, 0.8, 1.5][ci];
        var baseCPC   = [45, 28, 38, 70, 42][ci];

        // Ad type modifiers
        var roasMod = { 'Influencer':1.4, 'Internal UGC':1.2, 'Static Creative':0.9, 'Video Ad':1.15, 'Carousel':1.05 }[adType];
        var ctrMod  = { 'Influencer':1.5, 'Internal UGC':1.3, 'Static Creative':0.8, 'Video Ad':1.4, 'Carousel':1.1 }[adType];

        // Core metrics (Scaled 30x for realism & cohort sizes)
        var spend       = baseSpend * 30 * weekendMult * noise(0.2);
        var ctr         = baseCTR * ctrMod * fatigueFactor * noise(0.12);
        var cpc         = baseCPC / ctrMod * noise(0.1);
        var impressions = Math.round(spend / (cpc / 1000 * ctr * 10) * noise(0.15));
        var clicks      = Math.round(impressions * (ctr / 100));

        var roas        = baseRoas * roasMod * weekendMult * fatigueFactor * noise(0.15);
        var revenue     = spend * roas;

        // Calculate conversions in sync with AOV
        var conversions = Math.round(revenue / price);
        if (conversions === 0 && revenue > price * 0.45) conversions = 1;
        if (conversions > clicks) conversions = Math.round(clicks * 0.3);
        
        // Re-adjust revenue & ROAS for precision
        revenue = conversions * price;
        roas = spend > 0 ? revenue / spend : 0;
        
        var cpm         = impressions > 0 ? (spend / impressions * 1000) : 0;
        var frequency   = 1.5 + (day / 90) * 3.5 * noise(0.2);

        if (spend < 10) return;

        // ── FUNNEL METRICS ──────────────────────────────────
        var linkClicks = Math.round(clicks * (0.85 + Math.random() * 0.10));
        var lpvRate = 0.70 + Math.random() * 0.20;
        var landingPageViews = Math.round(linkClicks * lpvRate);

        // Add to Cart: 3-12% of LPV
        var atcBaseRate = { 'Influencer':0.10, 'Internal UGC':0.08, 'Static Creative':0.05, 'Video Ad':0.09, 'Carousel':0.07 }[adType];
        var atcRateFinal = atcBaseRate * ([0.8, 1.3, 1.1, 0.6, 0.9][ci]) * noise(0.25);
        var addToCart = Math.round(landingPageViews * atcRateFinal);

        // Hook & Hold
        var baseHook = hookByType[adType] || 0;
        var hookRate = baseHook > 0 ? Math.max(5, baseHook * fatigueFactor * noise(0.15)) : 0;
        var baseHold = holdByType[adType] || 0;
        var holdRate = baseHold > 0 ? Math.max(3, baseHold * fatigueFactor * noise(0.15)) : 0;

        var videoViews = (hookRate > 0) ? Math.round(impressions * 0.6 * noise(0.15)) : 0;
        var thruPlay   = (videoViews > 0) ? Math.round(videoViews * (holdRate / 100) * 0.8) : 0;

        // Generate sales ledger orders for META
        for (var c = 0; c < conversions; c++) {
          var isRepeat = Math.random() < 0.22 && customerRegistry.length > 0;
          var cust;
          if (isRepeat) {
            // Find an existing customer acquired at least 3 days ago
            var eligible = customerRegistry.filter(function(r) {
              return (d - r.acquisitionDate) >= (3 * 24 * 60 * 60 * 1000);
            });
            cust = eligible.length > 0 ? eligible[Math.floor(Math.random() * eligible.length)] : customerRegistry[Math.floor(Math.random() * customerRegistry.length)];
          } else {
            var custId = 'CUST-' + nextCustId++;
            cust = { id: custId, acquisitionDate: new Date(d), cohortMonth: monthStr, platform: 'Meta' };
            customerRegistry.push(cust);
          }
          var orderId = 'ORD-' + nextOrderId++;
          salesRows.push([
            orderId, cust.id, dateStr, round2(price * noise(0.05)), product, 'Meta', cust.cohortMonth
          ]);
        }

        // META row (28 cols)
        var tags = autoTagMetadata(camp, adSet, selectedAd);

        metaRows.push([
          dateStr,                                // Day
          'D2C India Main',                       // Account Name
          camp,                                   // Campaign name
          adSet,                                  // Ad set name
          selectedAd,                             // Ad name
          'INR',                                  // Currency
          round2(spend),                          // Amount spent (INR)
          Math.round(impressions),                // Impressions
          linkClicks,                             // Link clicks
          landingPageViews,                       // Landing page views
          addToCart,                              // Add To Cart
          conversions,                            // Purchases
          round2(revenue),                        // Purchases conversion value
          videoViews,                             // 3-second video plays
          thruPlay,                               // Video plays at 25%
          Math.round(thruPlay * 0.7),             // Video plays at 50%
          Math.round(thruPlay * 0.5),             // Video plays at 75%
          Math.round(thruPlay * 0.3),             // Video plays at 95%
          Math.round(thruPlay * 0.1),             // Video plays at 100%
          'https://example.com/products/' + encodeURIComponent(product.toLowerCase().replace(/ /g, '-')), // Website URL
          'adset_' + (100000000 + ti),            // Ad set ID
          'ad_' + (200000000 + ci * 10 + ti),     // Ad ID
          product,                                // Product
          tags.adType,                            // Ad Type
          tags.typeOfAd,                          // Type of Ad
          tags.influencerName,                    // Influencer Name
          tags.narrative,                         // Narrative
          tags.language                           // Language
        ]);

        // GOOGLE Ads — only Performance Scale & New Launch campaigns
        if (ci === 1 || ci === 3) {
          var gRoas    = roas * 0.9;
          var gSpend   = spend * 0.65;
          var gRevenue = gSpend * gRoas;
          var gImpr    = Math.round(impressions * 0.5);
          var gClicks  = Math.round(gImpr * (ctr * 0.8 / 100));
          var gLinkClk = Math.round(gClicks * 0.92);
          var gLPV     = Math.round(gLinkClk * (0.75 + Math.random() * 0.15));
          var gATC     = Math.round(gLPV * atcRateFinal * 0.85);
          
          var gConv = Math.round(gRevenue / price);
          if (gConv === 0 && gRevenue > price * 0.45) gConv = 1;
          if (gConv > gClicks) gConv = Math.round(gClicks * 0.3);
          
          gRevenue = gConv * price;
          gRoas = gSpend > 0 ? gRevenue / gSpend : 0;
          
          var qScore   = Math.min(10, Math.max(4, Math.round(6 + (gRoas - 2) * 1.5)));
          var gConvRate = gClicks > 0 ? round2((gConv / gClicks) * 100) : 0;
          var gHookRate = (adType === 'Video Ad') ? round2(hookRate * 0.7) : 0;
          var gHoldRate = (adType === 'Video Ad') ? round2(holdRate * 0.7) : 0;

          // Generate sales ledger orders for GOOGLE
          for (var c = 0; c < gConv; c++) {
            var isRepeat = Math.random() < 0.22 && customerRegistry.length > 0;
            var cust;
            if (isRepeat) {
              var eligible = customerRegistry.filter(function(r) {
                return (d - r.acquisitionDate) >= (3 * 24 * 60 * 60 * 1000);
              });
              cust = eligible.length > 0 ? eligible[Math.floor(Math.random() * eligible.length)] : customerRegistry[Math.floor(Math.random() * customerRegistry.length)];
            } else {
              var custId = 'CUST-' + nextCustId++;
              cust = { id: custId, acquisitionDate: new Date(d), cohortMonth: monthStr, platform: 'Google' };
              customerRegistry.push(cust);
            }
            var orderId = 'ORD-' + nextOrderId++;
            salesRows.push([
              orderId, cust.id, dateStr, round2(price * noise(0.05)), product, 'Google', cust.cohortMonth
            ]);
          }

          // GOOGLE row (25 cols)
          googleRows.push([
            dateStr, camp, adSet, selectedAd,
            round2(gSpend), gImpr, gClicks,
            gLinkClk, gLPV,
            gConv, round2(gRevenue), gATC,
            round2(ctr * 0.8), round2(cpc * 1.1), round2(gRoas),
            qScore, gConvRate, gHookRate, gHoldRate,
            tags.narrative, tags.typeOfAd, tags.language, tags.influencerName, product, tags.category
          ]);
        }
      });
    });
  }

  // Batch write META_DATA
  if (metaRows.length > 0) {
    metaSheet.getRange(2, 1, metaRows.length, 28).setValues(metaRows);
  }
  // Batch write GOOGLE_DATA
  if (googleRows.length > 0) {
    googleSheet.getRange(2, 1, googleRows.length, 25).setValues(googleRows);
  }
  // Batch write SALES_DATA
  if (salesRows.length > 0) {
    salesSheet.getRange(2, 1, salesRows.length, 7).setValues(salesRows);
  }

  // Product master
  var pmSheet = ss.getSheetByName(PRODUCT_SHEET);
  if (pmSheet.getLastRow() <= 1) {
    pmSheet.getRange(2, 1, 5, 5).setValues([
      ['Running Shoes X1',    3999, 1400, 0.65, 800],
      ['Sport Tee Pro',        1299,  450, 0.65, 300],
      ['Recovery Foam Roller',  999,  350, 0.65, 250],
      ['Gym Bag Elite',        1999,  700, 0.65, 450],
      ['Protein Shaker',        499,  175, 0.65, 120]
    ]);
  }

  SpreadsheetApp.getUi().alert(
    '✅ Scaled 90-day sample data inserted!\n\n' +
    '📘 Meta rows: ' + metaRows.length + '\n' +
    '🔶 Google rows: ' + googleRows.length + '\n' +
    '👥 Sales Ledger rows: ' + salesRows.length + '\n\n' +
    'Includes:\n' +
    '• 90 days of transactions (for cohorts & LTV)\n' +
    '• Linked customer acquisition & retention paths\n' +
    '• Creative fatigue simulation & advanced funnel parameters\n\n' +
    'Now open the Dashboard to explore!'
  );
}

function noise(range) {
  return 1 + (Math.random() - 0.5) * 2 * range;
}
