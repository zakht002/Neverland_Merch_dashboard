const SPREADSHEET_ID = '1Y2hMSDU9BTlk_nHwvxWL5KmGQMA5QnD3x0et-gebdko';
const EXTERNAL_SPREADSHEET_ID = '1KkGtDYCrXDBX4WTC5i9iwsKCBq6NeKUauwiGQ-C4Nb4';

function doGet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ssExt = SpreadsheetApp.openById(EXTERNAL_SPREADSHEET_ID);
    SpreadsheetApp.flush(); 
  } catch(e) {
    Logger.log("Automatic recalculation flush skipped or sheet unreachable: " + e.message);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ITVX Content Merchandising Report'); 
}

// ===== DOWNLOAD RAW DATA FOR ANY TABLE =====
function downloadTableData(sheetName, range) { 
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName); 
  return sheet.getRange(range).getDisplayValues();
}

// DISPLAY CURRENT TITLES WITH DYNAMIC METRIC ORDERING
function getSheetData(orderByMetric) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Pivot Current');
  const rawData = sheet.getDataRange().getDisplayValues();
  if(rawData.length <= 1) return rawData;
  
  const headers = rawData[0];
  let sortColIdx = -1;
  
  if (orderByMetric) {
    sortColIdx = headers.indexOf(orderByMetric);
  }
  
  if (sortColIdx === -1) {
    sortColIdx = headers.indexOf('TOTAL_HRS_ITD');
  }
  if (sortColIdx === -1) {
    sortColIdx = headers.indexOf('Total Hrs ITD');
  }
  
  const dataRows = rawData.slice(1);
  const cleanedRows = dataRows.filter(row => row[0] !== "Grand Total" && row[0] !== "");
  
  if (sortColIdx > -1) {
    cleanedRows.sort((a, b) => {
      let valA = parseFloat(a[sortColIdx].replace(/[%,\s]/g, '')) || 0;
      let valB = parseFloat(b[sortColIdx].replace(/[%,\s]/g, '')) || 0;
      return valB - valA; 
    });
  }
  
  return [headers].concat(cleanedRows);
}

// EXTRACT UPDATE DATE LOGIC
function getUpdateMetadataDate() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Current');
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0];
    const updateIdx = headers.indexOf('UPDATE_DATE');
    if (updateIdx > -1 && data.length > 1) {
      return data[1][updateIdx];
    }
  } catch(e) {
    Logger.log("Error finding update date: " + e.message);
  }
  return "06th Jul-26"; 
}

// DISPLAY BREAKOUT
function getBreakoutData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    hours: ss.getSheetByName('Breakout hours').getRange('F39:K43').getDisplayValues(),
    reach: ss.getSheetByName('Breakout % Reach').getRange('F40:K44').getDisplayValues()
  };
}

// REAL-TIME CATEGORY PERFORMANCE AGGREGATOR ENGINE
function generateLiveCategorySummaryModel(categoryName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const currentSheet = ss.getSheetByName('Current');
  const currentData = currentSheet.getDataRange().getDisplayValues();
  const currentHeaders = currentData[0];
  const cTitleIdx = currentHeaders.indexOf('CONTENT_TITLE_WITH_SEASON_NUMBER');
  const cCatIdx = currentHeaders.indexOf('title_reach_category');
  
  let titleCategoryMap = {};
  for (let r = 1; r < currentData.length; r++) {
    let tName = currentData[r][cTitleIdx];
    let cName = currentData[r][cCatIdx];
    if (tName) {
      titleCategoryMap[tName] = cName;
    }
  }
  
  const reachSheet = ss.getSheetByName('Daily % Reach (Raw)');
  const reachData = reachSheet.getDataRange().getDisplayValues();
  const reachHeaders = reachData[0];
  
  const rTitleIdx = reachHeaders.indexOf('CONTENT_TITLE_WITH_SEASON_NUMBER');
  const rDayIdx = reachHeaders.indexOf('DAY_INTERVAL'); 
  const rReachIdx = reachHeaders.indexOf('DAILY_ITD_REACH');
  
  const hoursSheet = ss.getSheetByName('Daily Hrs (Raw)');
  const hoursData = hoursSheet.getDataRange().getDisplayValues();
  const hoursHeaders = hoursData[0];
  
  const hTitleIdx = hoursHeaders.indexOf('CONTENT_TITLE_WITH_SEASON_NUMBER');
  const hDayIdx = hoursHeaders.indexOf('DAY_INTERVAL'); 
  const hHrsIdx = hoursHeaders.indexOf('TOTAL_HRS_ITD');

  let rGoal7 = 0.002; let rGoal28 = 0.005;
  let hGoal7 = 15000; let hGoal28 = 46000;
  
  if (categoryName === "Breakout Content") {
    rGoal7 = 0.004; rGoal28 = 0.01;
    hGoal7 = 40000; hGoal28 = 115000; 
  }

  let titleTracker = {};

  for (let i = 1; i < reachData.length; i++) {
    let title = reachData[i][rTitleIdx];
    let mappedCat = titleCategoryMap[title];
    if (!mappedCat || mappedCat !== categoryName) continue;
    
    let day = parseInt(reachData[i][rDayIdx]) || 0;
    
let reachStr = reachData[i][rReachIdx].toString().trim();
let reachVal = 0;
if (reachStr !== "" && reachStr !== "-") {
  reachVal = reachStr.includes('%') ? parseFloat(reachStr.replace('%','')) / 100 : parseFloat(reachStr);
  if (isNaN(reachVal)) reachVal = 0;
}

    if (!titleTracker[title]) {
      titleTracker[title] = { r7Fact: 0, r28Fact: 0, rMax: 0, h7Fact: 0, h28Fact: 0, hMax: 0 };
    }
    if (day === 7) titleTracker[title].r7Fact = reachVal;
    if (day === 28) titleTracker[title].r28Fact = reachVal;
    if (reachVal > titleTracker[title].rMax) titleTracker[title].rMax = reachVal;
  }

  for (let i = 1; i < hoursData.length; i++) {
    let title = hoursData[i][hTitleIdx];
    let mappedCat = titleCategoryMap[title];
    if (!mappedCat || mappedCat !== categoryName) continue;
    
    let day = parseInt(hoursData[i][hDayIdx]) || 0;
    let hrsVal = parseFloat(hoursData[i][hHrsIdx].replace(/,/g, '')) || 0;

    if (!titleTracker[title]) {
      titleTracker[title] = { r7Fact: 0, r28Fact: 0, rMax: 0, h7Fact: 0, h28Fact: 0, hMax: 0 };
    }
    if (day === 7) titleTracker[title].h7Fact = hrsVal;
    if (day === 28) titleTracker[title].h28Fact = hrsVal;
    if (hrsVal > titleTracker[title].hMax) titleTracker[title].hMax = hrsVal;
  }

  let reachRows = [];
  let hoursRows = [];

  Object.keys(titleTracker).forEach(title => {
    let t = titleTracker[title];
    reachRows.push([title, t.r7Fact, t.r28Fact, rGoal7, rGoal28, t.rMax]);
    hoursRows.push([title, t.h7Fact, t.h28Fact, hGoal7, hGoal28, t.hMax]);
  });

  return {
    reach: { headers: ["Title", "7D Fact", "28D Fact", "7D Goal", "28D Goal", "Current % reach"], rows: reachRows },
    hours: { headers: ["Title", "7D Fact", "28D Fact", "7D Goal", "28D Goal", "Current Hours"], rows: hoursRows }
  };
}

function getMonthlyEngagementData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetReach = ss.getSheetByName('Monthly Reach');
  const dataReach = sheetReach.getDataRange().getDisplayValues();
  const reachLabels = [];
  const reachValues = [];
  
  for(let i = 1; i < dataReach.length; i++) {
    if(!dataReach[i][1]) continue;
    reachLabels.push(dataReach[i][1]); 
    let val = parseFloat(dataReach[i][4].replace('%','')) || 0;
    if(val > 1) val = val / 100; 
    reachValues.push(val);
  }

  const sheetHours = ss.getSheetByName('Monthly Hours');
  const dataHours = sheetHours.getDataRange().getDisplayValues();
  const hoursLabels = [];
  const hoursValues = [];
  
  for(let i = 1; i < dataHours.length; i++) {
    if(!dataHours[i][1]) continue;
    hoursLabels.push(dataHours[i][1]); 
    hoursValues.push(parseFloat(dataHours[i][2].replace(/,/g,'')) || 0);
  }

  return {
    reach: { labels: reachLabels, values: reachValues },
    hours: { labels: hoursLabels, values: hoursValues }
  };
}

// ===== UPDATED ACCORDING TO SPECIFICATIONS: OVERALL PERFORMANCE ITD FROM PAST SHEET =====
function getPastPerformanceData(sortKey, sortOrder) {
  const ss = SpreadsheetApp.openById(EXTERNAL_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('All Past Perfomance');
  const rawData = sheet.getDataRange().getDisplayValues();
  if (rawData.length === 0) return [];
  
  const headers = rawData[0];
  const selectCols = [
    'CONTENT_TITLE_WITH_SEASON_NUMBER',
    'LAUNCH_DATE_SEASON',
    'EXPIRES_IN_COUNTRY_DATE_EST_SEASON',
    'Days on Platform',
    '% Reach ITD',
    'Total Hrs ITD'
  ];
  
  let targetIndexes = selectCols.map(c => headers.indexOf(c));
  let finalHeaders = ['CONTENT_TITLE_WITH_SEASON_NUMBER', 'SUNRISE', 'SUNSET', 'Days on Platform', '% Reach ITD', 'Total Hrs ITD'];
  
  let sortIndex = headers.indexOf(sortKey);
  if (sortIndex === -1) {
    sortIndex = headers.indexOf('Total Hrs ITD');
  }
  
  const dataRows = rawData.slice(1);
  if (sortIndex !== -1) {
    dataRows.sort((a, b) => {
      let valA = a[sortIndex].replace(/[^0-9.-]/g, '');
      let valB = b[sortIndex].replace(/[^0-9.-]/g, '');
      
      if (sortKey && (sortKey.toLowerCase().includes('date') || sortKey.toLowerCase().includes('season') || sortKey.toLowerCase().includes('sunrise') || sortKey.toLowerCase().includes('sunset'))) {
        return new Date(a[sortIndex]) - new Date(b[sortIndex]);
      }
      
      let numA = parseFloat(valA) || 0;
      let numB = parseFloat(valB) || 0;
      return numB - numA; 
    });
  }
  
  let transformedRows = dataRows.slice(0, 20).map(row => {
    return targetIndexes.map(idx => idx > -1 ? row[idx] : '');
  });
  
  return [finalHeaders].concat(transformedRows);
}

// MATCH % REACH ITD EXACTLY WITH THE LIVE PIVOT CURRENT TAB AS REQUESTED
function getHeaderMetadataProfile(selectedTitle) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const pivotSheet = ss.getSheetByName('Pivot Current');
  const pivotData = pivotSheet.getDataRange().getDisplayValues();
  const pivotHeaders = pivotData[0];
  
  const pTitleIdx = 0; 
  const pReachIdx = pivotHeaders.indexOf('% Reach ITD');
  const pHrsIdx = pivotHeaders.indexOf('Total Hrs ITD');
  const pDaysIdx = pivotHeaders.indexOf('Days on Platform');
  
  let verifiedReach = "0.0%";
  let verifiedHours = "0";
  let verifiedDays = "0";

  for (let r = 1; r < pivotData.length; r++) {
    if (pivotData[r][pTitleIdx] === selectedTitle || selectedTitle.includes(pivotData[r][pTitleIdx])) {
      if (pReachIdx > -1) verifiedReach = pivotData[r][pReachIdx];
      if (pHrsIdx > -1) verifiedHours = pivotData[r][pHrsIdx];
      if (pDaysIdx > -1) verifiedDays = pivotData[r][pDaysIdx];
      break;
    }
  }

  const currentSheet = ss.getSheetByName('Current');
  const currentData = currentSheet.getDataRange().getDisplayValues();
  const curHeaders = currentData[0];
  const catIdx = curHeaders.indexOf('title_reach_category');
  let titleCategory = "Library Content";

  for (let r = 1; r < currentData.length; r++) {
    if (currentData[r][curHeaders.indexOf('CONTENT_TITLE_WITH_SEASON_NUMBER')] === selectedTitle) {
      if (catIdx > -1) titleCategory = currentData[r][catIdx];
      break;
    }
  }
  
  let artSheet = ss.getSheetByName('Titles Art');
  let artUrl = "";
  if(artSheet) {
    let artData = artSheet.getDataRange().getDisplayValues();
    for(let k=1; k<artData.length; k++) {
      if(artData[k][0] === selectedTitle || selectedTitle.includes(artData[k][0])) {
         artUrl = artData[k][3]; 
         break;
      }
    }
  }
  
  return {
    name: selectedTitle,
    category: titleCategory,
    daysOnPlatform: verifiedDays,
    currentReach: verifiedReach,
    hoursStreamed: verifiedHours,
    artImageThumbnail: artUrl
  };
}

function getAdvancedMerchPivotData(selectedTitle) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Batch Container Merch');
  if (!sheet) return { impressionsByContainer: {}, playsByContainer: {}, curationImpressions: {}, curationPlays: {}, timelineImpressions: {}, timelinePlays: {} };
  
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return { impressionsByContainer: {}, playsByContainer: {}, curationImpressions: {}, curationPlays: {}, timelineImpressions: {}, timelinePlays: {} };
  
  const headers = data[0];
  const titleIdx = headers.indexOf('CONTENT_TITLE_WITH_SEASON_NUMBER');
  const daysOnPlatformIdx = headers.indexOf('DAYS_ON_PLATFORM');
  const containerIdx = headers.indexOf('CONTAINER_TITLE'); 
  const mixIdx = headers.indexOf('IMPRESSION_BEHAVIOUR_EDITORIAL'); 
  const impIdx = headers.indexOf('IMPRESSIONS'); 
  const playIdx = headers.indexOf('PLAYS'); 
  
  let impressionsPivot = {}; let playsPivot = {};
  let mixImpressions = {}; let mixPlays = {};
  let timelineImpressions = {}; let timelinePlays = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowTitle = row[titleIdx];
    const concatenatedTitleMatch = row[daysOnPlatformIdx] + " " + rowTitle;
    
    if (rowTitle !== selectedTitle && concatenatedTitleMatch !== selectedTitle) continue;
    
    let container = row[containerIdx] || "Other Containers";
    let mixType = row[mixIdx] || "3. Algorithmic";
    let dateKey = row[daysOnPlatformIdx] ? "Day " + row[daysOnPlatformIdx] : "Day 1";
    
    let imps = parseFloat(row[impIdx].replace(/,/g,'')) || 0;
    let plays = parseFloat(row[playIdx].replace(/,/g,'')) || 0;
    
    impressionsPivot[container] = (impressionsPivot[container] || 0) + imps;
    playsPivot[container] = (playsPivot[container] || 0) + plays;
    mixImpressions[mixType] = (mixImpressions[mixType] || 0) + imps;
    mixPlays[mixType] = (mixPlays[mixType] || 0) + plays;
    
    timelineImpressions[dateKey] = (timelineImpressions[dateKey] || 0) + imps;
    timelinePlays[dateKey] = (timelinePlays[dateKey] || 0) + plays;
  }
  
  return {
    impressionsByContainer: impressionsPivot, playsByContainer: playsPivot,
    curationImpressions: mixImpressions, curationPlays: mixPlays,
    timelineImpressions: timelineImpressions, timelinePlays: timelinePlays
  };
}

function getTitleReachTrajectoryData(selectedTitle) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. Map Title Categories from Current Tab to find target category definitions
  const currentSheet = ss.getSheetByName('Current');
  const currentData = currentSheet.getDataRange().getDisplayValues();
  const curHeaders = currentData[0];
  const titleIdx = curHeaders.indexOf('CONTENT_TITLE_WITH_SEASON_NUMBER');
  const catIdx = curHeaders.indexOf('title_reach_category');
  
  let titleCategory = "Library Content"; 
  if (titleIdx > -1 && catIdx > -1) {
    for (let r = 1; r < currentData.length; r++) {
      if (currentData[r][titleIdx] === selectedTitle) {
        titleCategory = currentData[r][catIdx] ? currentData[r][catIdx].toString().trim() : "Library Content";
        break;
      }
    }
  }
  
  // 2. Assign Tier Specific Targets dynamically (Using flexible case-insensitive match)
  let target7D = 0.002;  
  let target28D = 0.005; 
  
  let lowCaseCat = titleCategory.toLowerCase();
  if (lowCaseCat.includes("breakout")) {
    target7D = 0.004;  
    target28D = 0.01;   
  } else if (lowCaseCat.includes("premium")) {
    target7D = 0.002;
    target28D = 0.005;
  }
  
  // 3. Process raw information directly from Daily % Reach (Raw)
  const reachSheet = ss.getSheetByName('Daily % Reach (Raw)');
  const reachData = reachSheet.getDataRange().getDisplayValues();
  const reachHeaders = reachData[0];
  
  const rbIdx = reachHeaders.indexOf('CONTENT_TITLE_WITH_SEASON_NUMBER');
  const reIdx = reachHeaders.indexOf('DAY_INTERVAL');
  const rfIdx = reachHeaders.indexOf('DAILY_ITD_REACH');
  
  let parsedReachPoints = [];
  for (let i = 1; i < reachData.length; i++) {
    if (reachData[i][rbIdx] === selectedTitle) {
      let interval = parseInt(reachData[i][reIdx]) || 0;
      let reachStr = reachData[i][rfIdx].toString().trim().replace(/,/g, '');
      let reachFact = null;
      
      if (reachStr !== "" && reachStr !== "-") {
        reachFact = reachStr.includes('%') ? parseFloat(reachStr.replace('%','')) / 100 : parseFloat(reachStr);
        if (isNaN(reachFact)) reachFact = null;
      }
      parsedReachPoints.push({ day: interval, fact: reachFact });
    }
  }
  
  parsedReachPoints.sort((a, b) => a.day - b.day);
  
  // 4. Construct trajectory run path model up to Day 28 bound limit arrays
  let finalReachTrajectoryArr = [];
  let lastKnownReachTrajectory = 0;
  let currentActualReachRate = 0;
  
  for (let day = 1; day <= 28; day++) {
    let match = parsedReachPoints.find(p => p.day === day);
    let factVal = (match && match.fact !== null) ? match.fact : null;
    let computedTrajectory = 0;
    
    if (factVal !== null) {
      computedTrajectory = factVal;
      currentActualReachRate = factVal;
    } else {
      let remainingDays = 28 - (day - 1);
      if (remainingDays <= 0) remainingDays = 1;
      computedTrajectory = lastKnownReachTrajectory + ((target28D - lastKnownReachTrajectory) / remainingDays);
    }
    
    lastKnownReachTrajectory = computedTrajectory;
    finalReachTrajectoryArr.push({
      day: day,
      trajectoryValue: computedTrajectory,
      isFact: (factVal !== null)
    });
  }
  
  let daysElapsed = finalReachTrajectoryArr.filter(p => p.isFact).length;
  let remainingTime = 28 - daysElapsed;
  let targetShortfall = target28D - currentActualReachRate;
  let neededDailyUplift = remainingTime > 0 ? (targetShortfall / remainingTime) : 0;
  
  let currentVelocity = 0;
  if (daysElapsed > 1) {
     currentVelocity = currentActualReachRate / daysElapsed;
  }
  
  let statusComment = "On Target";
  if (currentVelocity < (target28D / 28) && targetShortfall > 0) {
    statusComment = "Below Target";
  }
  
  return {
    title: selectedTitle,
    category: titleCategory,
    target7D: target7D,
    target28D: target28D,
    daysElapsed: daysElapsed,
    currentReachRate: currentActualReachRate,
    neededDailyRate: neededDailyUplift,
    status: statusComment,
    chartReachData: finalReachTrajectoryArr
  };
}

function getTitleTrajectoryData(selectedTitle) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. Map Title Categories from Current Tab to find target category definitions
  const currentSheet = ss.getSheetByName('Current');
  const currentData = currentSheet.getDataRange().getDisplayValues();
  const curHeaders = currentData[0];
  const titleIdx = curHeaders.indexOf('CONTENT_TITLE_WITH_SEASON_NUMBER');
  const catIdx = curHeaders.indexOf('title_reach_category');
  
  let titleCategory = "Library Content";
  if (titleIdx > -1 && catIdx > -1) {
    for (let r = 1; r < currentData.length; r++) {
      if (currentData[r][titleIdx] === selectedTitle) {
        titleCategory = currentData[r][catIdx] ? currentData[r][catIdx].toString().trim() : "Library Content";
        break;
      }
    }
  }
  
  // 2. Assign dynamic tier-specific hour benchmarks
  let targetHrs7D = 15000;
  let targetHrs28D = 46000;
  let lowCaseCat = titleCategory.toLowerCase();
  if (lowCaseCat.includes("breakout")) {
    targetHrs7D = 40000;
    targetHrs28D = 115000;
  }
  
  // 3. Process raw stream hours metrics directly from the sheet
  const hoursSheet = ss.getSheetByName('Daily Hrs (Raw)');
  const hoursData = hoursSheet.getDataRange().getDisplayValues();
  const hoursHeaders = hoursData[0];
  
  const hTitleIdx = hoursHeaders.indexOf('CONTENT_TITLE_WITH_SEASON_NUMBER');
  const hDayIdx = hoursHeaders.indexOf('DAY_INTERVAL');
  const hHrsIdx = hoursHeaders.indexOf('TOTAL_HRS_ITD');
  
  let parsedHoursPoints = [];
  for (let i = 1; i < hoursData.length; i++) {
    if (hoursData[i][hTitleIdx] === selectedTitle) {
      let interval = parseInt(hoursData[i][hDayIdx]) || 0;
      let hrsStr = hoursData[i][hHrsIdx].toString().trim().replace(/,/g, '');
      let hrsFact = (hrsStr !== "" && hrsStr !== "-") ? parseFloat(hrsStr) : null;
      if (isNaN(hrsFact)) hrsFact = null;
      parsedHoursPoints.push({ day: interval, fact: hrsFact });
    }
  }
  
  parsedHoursPoints.sort((a, b) => a.day - b.day);
  
  // 4. Construct historical actual path + forecasted projection bound values
  let finalHoursTrajectoryArr = [];
  let lastKnownHoursTrajectory = 0;
  
  for (let day = 1; day <= 28; day++) {
    let match = parsedHoursPoints.find(p => p.day === day);
    let factVal = (match && match.fact !== null) ? match.fact : null;
    let computedTrajectory = 0;
    
    if (factVal !== null) {
      computedTrajectory = factVal;
    } else {
      let remainingDays = 28 - (day - 1);
      if (remainingDays <= 0) remainingDays = 1;
      computedTrajectory = lastKnownHoursTrajectory + ((targetHrs28D - lastKnownHoursTrajectory) / remainingDays);
    }
    
    lastKnownHoursTrajectory = computedTrajectory;
    finalHoursTrajectoryArr.push({
      day: day,
      trajectoryValue: computedTrajectory,
      isFact: (factVal !== null)
    });
  }
  
  return {
    title: selectedTitle,
    targetHrs7D: targetHrs7D,
    targetHrs28D: targetHrs28D,
    chartHrsData: finalHoursTrajectoryArr
  };
}
