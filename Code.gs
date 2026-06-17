

/**
 * Dynamic Apps Script Backend Controller for ITVX Title Framework Platform
 * Custom Range Binding Specification Build
 * File Name Mandate: Code.gs
 */

const SPREADSHEET_ID = "1M6-RXZNZ7fefafdB2bjy8vuHBuwvQhcd8jP5ONR2NgY";

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
      .setTitle('ITVX Performance Mapping Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Returns active configuration variables and Looker deep-link targets
 */
function getAppMetadata() {
  return {
    lookerBaseUrl: "https://looker.itv.com/dashboards/13128",
    defaultCountry: "GB",
    dateThreshold: "after 2025/07/25"
  };
}

/**
 * Main ingestion controller fetching custom ranges across your split sheets
 */
function fetchDashboardData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // =========================================================================
    // LAYER 1: CURRENT BATCH & KPIS (Pivot Current A1:M13)
    // =========================================================================
    const pivotSheet = ss.getSheetByName("Pivot Current");
    const pivotRows = pivotSheet.getRange("A1:M13").getValues();
    
    let currentBatchData = [];
    for (let i = 1; i < pivotRows.length; i++) {
      let row = pivotRows[i];
      if (!row[0] || row[0].toString().trim() === "") continue; 
      
      currentBatchData.push({
        title: row[0] ? row[0].toString() : "",                 // Col A: Title
        sunrise: row[1] ? row[1].toString() : "",               // Col B: Sunrise
        sunset: row[2] ? row[2].toString() : "",                // Col C: Sunset
        daysOnPlatform: row[3] ? parseInt(row[3], 10) : 0,      // Col D: Days on Platform
        epCount: row[4] ? parseInt(row[4], 10) : 0,             // Col E: Ep. count
        epLength: row[5] ? parseInt(row[5], 10) : 0,            // Col F: Ep. Length
        availHrs: row[6] ? parseFloat(row[6]) : 0,              // Col G: Avail. Hrs
        category: row[7] ? row[7].toString() : "",              // Col H: Category
        reachItd: row[8] ? parseFloat(row[8]) : 0,              // Col I: % Reach ITD
        totalHrsItd: row[9] ? parseFloat(row[9]) : 0,           // Col J: Total Hrs ITD
        totalImpressions: row[10] ? parseInt(row[10], 10) : 0,  // Col K: Total impressions
        ctr: row[11] ? parseFloat(row[11]) : 0,                 // Col L: CTR
        ectr: row[12] ? parseFloat(row[12]) : 0                 // Col M: eCTR
      });
    }

    // =========================================================================
    // LAYER 2: CATEGORY PERFORMANCE (Multi-Tab Split Ingestion)
    // =========================================================================
    let categoryPerformanceData = {
      "Breakout": [],
      "Premium": [],
      "Library": []
    };

    // Mapping arrays directly to specified layout matrices
    const tabRangeConfig = [
      { category: "Breakout", type: "Hours", tab: "Breakout Hrs", range: "F39:K44" },
      { category: "Breakout", type: "Reach", tab: "Breakout % Reach", range: "F39:K44" },
      { category: "Premium",  type: "Hours", tab: "Premium Hrs",  range: "F39:K44" },
      { category: "Premium",  type: "Reach", tab: "Premium % Reach", range: "F39:K44" },
      { category: "Library", type: "Hours", tab: "Library Hrs",  range: "F39:K44" },
      { category: "Library", type: "Reach", tab: "Library % Reach", range: "F39:K44" }
    ];

    tabRangeConfig.forEach(cfg => {
      const sheet = ss.getSheetByName(cfg.tab);
      if (sheet) {
        const matrixValues = sheet.getRange(cfg.range).getValues();
        
        matrixValues.forEach(row => {
          if (row[0] && row[0].toString().trim() !== "") {
            categoryPerformanceData[cfg.category].push({
              title: row[0].toString(),
              metricType: cfg.type,
              val7d: row[1] ? parseFloat(row[1]) : 0,
              val28d: row[2] ? parseFloat(row[2]) : 0,
              valItd: row[3] ? parseFloat(row[3]) : 0,
              target28d: row[4] ? parseFloat(row[4]) : 0
            });
          }
        });
      }
    });

    // =========================================================================
    // LAYER 3: STRATEGIC FOOTERS DATA
    // =========================================================================
    let footersData = {
      "Tab1": "Source: Dynamic Matrix Engine Feed via 'Pivot Current' Matrix Panel.",
      "CategoryViews": "Source Ranges: Multi-pulse structural limits blocks (Rows 39-44 Execution Framework)."
    };
    
    return {
      currentBatch: currentBatchData,
      categoryPerformance: categoryPerformanceData,
      footers: footersData
    };
    
  } catch (err) {
    Logger.log("Critical Range Ingestion Failure: " + err.toString());
    throw new Error("Pipeline Ingestion Process Aborted: " + err.message);
  }
}

/**
 * NOTE FOR NEXT MILESTONE IMPLEMENTATION:
 * When adding second-layer drill downs, remember to build a dedicated 'Daily_Trends_Log' 
 * tab inside your G-Sheet workbook layout. A trend extraction method will hook right here.
 */
