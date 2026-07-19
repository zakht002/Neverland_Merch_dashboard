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
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID)    .getSheetByName(sheetName); 

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
    let reachVal = parseFloat(reachData[i][rReachIdx].replace(/[%]/g, '')) / 100 || 0;

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

function getPastPerformanceData(sortKey, sortOrder) {
  const ss = SpreadsheetApp.openById(EXTERNAL_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('All Past Perfomance');
  const rawData = sheet.getDataRange().getDisplayValues();
  if (rawData.length === 0) return [];
  
  const headers = rawData[0];
  let sortIndex = -1;
  
  if (sortKey) {
    for (let j = 0; j < headers.length; j++) {
      if (headers[j].toLowerCase().trim() === sortKey.toLowerCase().trim() || headers[j] === sortKey) {
        sortIndex = j;
        break;
      }
    }
  }
  
  if (sortIndex === -1) {
    for (let j = 0; j < headers.length; j++) {
      if (headers[j].includes('%') || headers[j].toLowerCase().includes('reach')) {
        sortIndex = j;
        break;
      }
    }
  }
  
  const dataRows = rawData.slice(1);
  if (sortIndex !== -1) {
    dataRows.sort((a, b) => {
      let valA = a[sortIndex].replace(/[%,\s]/g, '');
      let valB = b[sortIndex].replace(/[%,\s]/g, '');
      
      if (headers[sortIndex].toLowerCase().includes('date') || headers[sortIndex].toLowerCase().includes('sun')) {
        return sortOrder === 'ASC' ? new Date(a[sortIndex]) - new Date(b[sortIndex]) : new Date(b[sortIndex]) - new Date(a[sortIndex]);
      }
      
      let numA = parseFloat(valA) || 0;
      let numB = parseFloat(valB) || 0;
      return sortOrder === 'ASC' ? numA - numB : numB - numA;
    });
  }
  
  return [headers].concat(dataRows.slice(0, 20));
}

function getHeaderMetadataProfile(selectedTitle) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const currentSheet = ss.getSheetByName('Current');
  const currentData = currentSheet.getDataRange().getDisplayValues();
  const curHeaders = currentData[0];
  
  const titleIdx = curHeaders.indexOf('CONTENT_TITLE_WITH_SEASON_NUMBER');
  const catIdx = curHeaders.indexOf('title_reach_category');
  const daysIdx = curHeaders.indexOf('DAYS_ON_PLATFORM');
  const reachIdx = curHeaders.indexOf('PERC_REACH_ITD');
  const hrsIdx = curHeaders.indexOf('TOTAL_HRS_ITD');
  
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
  
  let profile = {
    name: selectedTitle,
    category: "Library Content",
    daysOnPlatform: "0",
    currentReach: "0.0%",
    hoursStreamed: "0",
    artImageThumbnail: artUrl
  };
  
  if (titleIdx > -1) {
    for (let r = 1; r < currentData.length; r++) {
      if (currentData[r][titleIdx] === selectedTitle) {
        if(catIdx > -1) profile.category = currentData[r][catIdx];
        if(daysIdx > -1) profile.daysOnPlatform = currentData[r][daysIdx];
        if(reachIdx > -1) {
          let num = parseFloat(currentData[r][reachIdx].replace('%','')) || 0;
          if(num < 1 && num > 0) num = num * 100;
          profile.currentReach = num.toFixed(1) + "%";
        }
        if(hrsIdx > -1) profile.hoursStreamed = currentData[r][hrsIdx];
        break;
      }
    }
  }
  return profile;
}

function getAdvancedMerchPivotData(selectedTitle) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Batch Container Merch');
  const data = sheet.getDataRange().getDisplayValues();
  
  const titleIdx = 4; 
  const containerIdx = 21; 
  const mixIdx = 22; 
  const impIdx = 24; 
  const playIdx = 25; 
  const dateIdx = 29; 
  
  let impressionsPivot = {}; let playsPivot = {};
  let mixImpressions = {}; let mixPlays = {};
  let timelineImpressions = {}; let timelinePlays = {};
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][titleIdx] !== selectedTitle) continue;
    let container = data[i][containerIdx] || "Other Containers";
    let mixType = data[i][mixIdx] || "3. Algorithmic";
    let dateKey = data[i][dateIdx] || "Day 1";
    let imps = parseFloat(data[i][impIdx].replace(/,/g,'')) || 0;
    let plays = parseFloat(data[i][playIdx].replace(/,/g,'')) || 0;
    
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

// ===== INJECTED TRAJECTORY LOGIC FUNCTION ENGINE =====
function getTitleTrajectoryData(selectedTitle) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const currentSheet = ss.getSheetByName('Current');
  const currentData = currentSheet.getDataRange().getDisplayValues();
  const curHeaders = currentData[0];
  const titleIdx = curHeaders.indexOf('CONTENT_TITLE_WITH_SEASON_NUMBER');
  const catIdx = curHeaders.indexOf('title_reach_category');
  
  let titleCategory = "Library Content"; 
  if (titleIdx > -1 && catIdx > -1) {
    for(let r=1; r<currentData.length; r++) {
      if(currentData[r][titleIdx] === selectedTitle) {
        titleCategory = currentData[r][catIdx];
        break;
      }
    }
  }
  
  let target7D = 0.002;  
  let target28D = 0.005; 
  if (titleCategory === "Breakout Content") {
    target7D = 0.004;  
    target28D = 0.01;   
  }
  
  const rawSheet = ss.getSheetByName('Daily % Reach (Raw)');
  const rawData = rawSheet.getDataRange().getDisplayValues();
  const rawHeaders = rawData[0];
  
  const bIdx = rawHeaders.indexOf('CONTENT_TITLE_WITH_SEASON_NUMBER');
  const eIdx = rawHeaders.indexOf('DAY_INTERVAL');
  const fIdx = rawHeaders.indexOf('DAILY_ITD_REACH');
  
  let parsedPoints = [];
  for(let i=1; i<rawData.length; i++) {
    if(rawData[i][bIdx] === selectedTitle) {
      let interval = parseInt(rawData[i][eIdx]) || 0;
      let reachStr = rawData[i][fIdx].toString().trim();
      let reachFact = null;
      
      if(reachStr !== "" && reachStr !== "-") {
        reachFact = reachStr.includes('%') ? parseFloat(reachStr.replace('%',''))/100 : parseFloat(reachStr);
        if(isNaN(reachFact)) reachFact = null;
      }
      
      parsedPoints.push({ day: interval, fact: reachFact });
    }
  }
  
  parsedPoints.sort((a,b) => a.day - b.day);
  
  let finalTrajectoryArr = [];
  let lastKnownTrajectory = 0;
  let currentActualRate = 0;
  
  for(let day = 1; day <= 28; day++) {
    let match = parsedPoints.find(p => p.day === day);
    let factVal = (match && match.fact !== null) ? match.fact : null;
    let computedTrajectory = 0;
    
    if (factVal !== null) {
      computedTrajectory = factVal;
      currentActualRate = factVal; 
    } else {
      let remainingDays = 28 - (day - 1);
      if (remainingDays <= 0) remainingDays = 1;
      computedTrajectory = lastKnownTrajectory + ((target28D - lastKnownTrajectory) / remainingDays);
    }
    
    lastKnownTrajectory = computedTrajectory;
    finalTrajectoryArr.push({
      day: day,
      trajectoryValue: computedTrajectory,
      isFact: (factVal !== null)
    });
  }
  
  let daysElapsed = finalTrajectoryArr.filter(p => p.isFact).length;
  let remainingTime = 28 - daysElapsed;
  let targetShortfall = target28D - currentActualRate;
  let neededDailyUplift = remainingTime > 0 ? (targetShortfall / remainingTime) : 0;
  
  let currentVelocity = 0;
  if(daysElapsed > 1) {
     currentVelocity = currentActualRate / daysElapsed;
  }
  
  let statusComment = "On Target";
  if(currentVelocity < (target28D / 28) && targetShortfall > 0) {
    statusComment = "Below Target";
  }

  // ========================================================
  // ========== DYNAMIC HOURS TRAJECTORY SUBSYSTEM ==========
  // ========================================================
  let targetHrs7D = 15000; 
  let targetHrs28D = 46000;
  if (titleCategory === "Breakout Content") {
    targetHrs7D = 40000;
    targetHrs28D = 115000;
  }

  const hrsSheet = ss.getSheetByName('Daily Hrs (Raw)');
  const hrsData = hrsSheet.getDataRange().getDisplayValues();
  const hrsHeaders = hrsData[0];

  const hbIdx = hrsHeaders.indexOf('CONTENT_TITLE_WITH_SEASON_NUMBER');
  const heIdx = hrsHeaders.indexOf('DAY_INTERVAL');
  const hfIdx = hrsHeaders.indexOf('TOTAL_HRS_ITD');

  let parsedHrsPoints = [];
  for (let i = 1; i < hrsData.length; i++) {
    if (hrsData[i][hbIdx] === selectedTitle) {
      let interval = parseInt(hrsData[i][heIdx]) || 0;
      let hrsStr = hrsData[i][hfIdx].toString().trim().replace(/,/g, '');
      let hrsFact = null;
      if (hrsStr !== "" && hrsStr !== "-") {
        hrsFact = parseFloat(hrsStr);
        if (isNaN(hrsFact)) hrsFact = null;
      }
      parsedHrsPoints.push({ day: interval, fact: hrsFact });
    }
  }

  parsedHrsPoints.sort((a, b) => a.day - b.day);

  let finalHrsTrajectoryArr = [];
  let lastKnownHrsTrajectory = 0;
  let currentActualHrsRate = 0;

  for (let day = 1; day <= 28; day++) {
    let match = parsedHrsPoints.find(p => p.day === day);
    let factVal = (match && match.fact !== null) ? match.fact : null;
    let computedTrajectory = 0;

    if (factVal !== null) {
      computedTrajectory = factVal;
      currentActualHrsRate = factVal;
    } else {
      let remainingDays = 28 - (day - 1);
      if (remainingDays <= 0) remainingDays = 1;
      computedTrajectory = lastKnownHrsTrajectory + ((targetHrs28D - lastKnownHrsTrajectory) / remainingDays);
    }

    lastKnownHrsTrajectory = computedTrajectory;
    finalHrsTrajectoryArr.push({
      day: day,
      trajectoryValue: computedTrajectory,
      isFact: (factVal !== null)
    });
  }
  
  return {
    title: selectedTitle,
    category: titleCategory,
    target7D: target7D,
    target28D: target28D,
    daysElapsed: daysElapsed,
    currentRate: currentActualRate,
    neededDailyRate: neededDailyUplift,
    status: statusComment,
    chartData: finalTrajectoryArr,
    targetHrs7D: targetHrs7D,
    targetHrs28D: targetHrs28D,
    currentHrsRate: currentActualHrsRate,
    chartHrsData: finalHrsTrajectoryArr
  };
}

function getAllUniqueBatchTitlesList() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Pivot Current');
  const data = sheet.getDataRange().getDisplayValues();
  let titles = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0] !== "Grand Total" && data[i][0] !== "Row Labels") titles.push(data[i][0]);
  }
  return titles;
}
