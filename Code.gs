const SPREADSHEET_ID = '1Y2hMSDU9BTlk_nHwvxWL5KmGQMA5QnD3x0et-gebdko';
const EXTERNAL_SPREADSHEET_ID = '1KkGtDYCrXDBX4WTC5i9iwsKCBq6NeKUauwiGQ-C4Nb4';

function doGet() {
  // AUTOMATIC REFRESH LAYER: Forces Google Sheets to flush caches and evaluate latest formula changes upon opening
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ssExt = SpreadsheetApp.openById(EXTERNAL_SPREADSHEET_ID);
    SpreadsheetApp.flush(); // Forces a live background calculation check
  } catch(e) {
    Logger.log("Automatic recalculation flush skipped or sheet unreachable: " + e.message);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ITVX Content Merchandising Report'); 
}

// ===== DOWNLOAD RAW DATA FOR ANY TABLE =====
function downloadTableData(sheetName, range) { 
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID) 
    .getSheetByName(sheetName); 

  return sheet.getRange(range).getDisplayValues();
}

// DISPLAY CURRENT TITLES
function getSheetData() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Pivot Current');
  return sheet.getDataRange().getDisplayValues();
}

// DISPLAY BREAKOUT
function getBreakoutData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  return {
    hours: ss.getSheetByName('Breakout hours').getRange('F39:K43').getDisplayValues(),
    reach: ss.getSheetByName('Breakout % Reach').getRange('F40:K44').getDisplayValues()
  };
}

// CHART DATA
function getChartData() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID)
    .getSheetByName('Breakout % Reach');

  const data = sheet.getRange('F40:K44').getDisplayValues(); 
  const headers = data[0];

  const titleCol = headers.indexOf('Title');
  const fact7Col = headers.indexOf('7D Fact');
  const fact28Col = headers.indexOf('28D Fact');
  const goal7Col = headers.indexOf('7D Goal');
  const goal28Col = headers.indexOf('28D Goal');
  const currentFactCol = headers.indexOf('Current % reach');

  const titles = [], fact7 = [], fact28 = [], currentFact = [], goal7 = [], goal28 = [];

  function parsePercent(v) {
    if (!v) return 0;
    const str = v.toString().trim();
    if (str.includes('%')) return parseFloat(str.replace('%', '')) / 100;
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  }

  for (let i = 1; i < data.length; i++) {
    if (!data[i][titleCol]) continue;
    titles.push(data[i][titleCol]);
    fact7.push(parsePercent(data[i][fact7Col]));
    fact28.push(parsePercent(data[i][fact28Col]));
    goal7.push(parsePercent(data[i][goal7Col]));
    goal28.push(parsePercent(data[i][goal28Col]));
    currentFact.push(parsePercent(data[i][currentFactCol]));
  }

  return { titles, fact7, fact28, goal7, goal28, currentFact };
}

// Helper function to process ANY sheet range into the format the charts expect
function processBreakoutSheet(sheetName, range) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  const data = sheet.getRange(range).getDisplayValues();
  const headers = data[0];

  const titleCol = headers.indexOf('Title');
  const fact7Col = headers.indexOf('7D Fact');
  const fact28Col = headers.indexOf('28D Fact');
  const goal7Col = headers.indexOf('7D Goal');
  const goal28Col = headers.indexOf('28D Goal');
  const currentFactCol = headers.indexOf('Current % reach') > -1 ? headers.indexOf('Current % reach') : headers.indexOf('Current Hours');

  const titles = [], fact7 = [], fact28 = [], currentFact = [], goal7 = [], goal28 = [];

  function parseVal(v) {
    if (!v) return 0;
    const str = v.toString().trim();
    if (str.includes('%')) return parseFloat(str.replace('%', '')) / 100;
    const num = parseFloat(str.replace(/,/g, ''));
    return isNaN(num) ? 0 : num;
  }

  for (let i = 1; i < data.length; i++) {
    if (!data[i][titleCol]) continue;
    titles.push(data[i][titleCol]);
    fact7.push(parseVal(data[i][fact7Col]));
    fact28.push(parseVal(data[i][fact28Col]));
    goal7.push(parseVal(data[i][goal7Col]));
    goal28.push(parseVal(data[i][goal28Col]));
    currentFact.push(parseVal(data[i][currentFactCol]));
  }
  return { titles, fact7, fact28, goal7, goal28, currentFact };
}

function getBreakoutChartData() {
  return {
    reach: processBreakoutSheet('Breakout % Reach', 'F40:K44'),
    hours: processBreakoutSheet('Breakout hours', 'F39:K43')
  };
}

// TITLE MERCH DATA
function getBatchDetails(title) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Batch Container Merch');
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  
  const rows = data.slice(1).filter(row => row[4] === title);
  
  const headerInfo = rows.length > 0 ? {
    fieldE: rows[0][4],
    fieldH: rows[0][7],
    fieldJ: rows[0][9],
    fieldAA: rows[0][26]
  } : {};

  const tableData = rows.map(r => [
    r[4],
    r[21],
    r[22],
    r[23],
    r[24],
    r[25]
  ]).sort((a, b) =>
    parseFloat(String(b[4]).replace(/,/g, '')) -   
    parseFloat(String(a[4]).replace(/,/g, ''))     
  );

  return {
    headers: ['Title', 'CONTAINER_TITLE', 'IMPRESSION_BEHAVIOUR_EDITORIAL', 'IMPRESSION_BEHAVIOUR_PERSONALISATION', 'IMPRESSIONS', 'PLAYS'],
    rows: tableData,
    summary: headerInfo
  };
}

function getTitleVisualData(title) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Batch Container Merch');
  const data = sheet.getDataRange().getDisplayValues();
  
  const rows = data.slice(1).filter(row => row[4] === title);
  
  return rows.map(r => ({
    container: r[21],
    impressions: parseFloat(r[24]) || 0,
    plays: parseFloat(r[25]) || 0
  }));
}

function getMonthlyEngagementData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Monthly Engagement');
  
  const dataReach = sheet.getRange('A1:C12').getDisplayValues();
  const labelsReach = [];
  const valuesReach = [];
  let sumReach = 0, countReach = 0;
  
  for (let i = 1; i < dataReach.length; i++) {
    if (!dataReach[i][0]) continue;
    labelsReach.push(dataReach[i][0]);
    const val = parseFloat(dataReach[i][2].replace(/,/g, '')) || 0;
    valuesReach.push(val);
    sumReach += val;
    countReach++;
  }
  const avgReach = countReach > 0 ? sumReach / countReach : 0;
  
  const dataHrs = sheet.getRange('A20:B33').getDisplayValues();
  const labelsHrs = [];
  const valuesHrs = [];
  let sumHrs = 0, countHrs = 0;
  
  for (let i = 1; i < dataHrs.length; i++) {
    if (!dataHrs[i][0]) continue;
    labelsHrs.push(dataHrs[i][0]);
    const val = parseFloat(dataHrs[i][1].replace(/,/g, '')) || 0;
    valuesHrs.push(val);
    sumHrs += val;
    countHrs++;
  }
  const avgHrs = countHrs > 0 ? sumHrs / countHrs : 0;
  
  return {
    reach: { labels: labelsReach, values: valuesReach, avg: avgReach },
    hours: { labels: labelsHrs, values: valuesHrs, avg: avgHrs }
  };
}

function parseCustomVal(v) {
  if (!v) return 0;
  const str = v.toString().trim();
  if (str.includes('%')) return parseFloat(str.replace('%', '')) / 100;
  const num = parseFloat(str.replace(/,/g, ''));
  return isNaN(num) ? 0 : num;
}

function processPerformanceTiers(sheetName, rangeX, rangeY) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  const dataX = sheet.getRange(rangeX).getDisplayValues();
  const dataY = sheet.getRange(rangeY).getDisplayValues();
  
  const labels = [];
  const values = [];
  
  for (let i = 1; i < dataX.length; i++) {
    if (!dataX[i][0]) continue;
    labels.push(dataX[i][0]);
  }
  for (let i = 1; i < dataY.length; i++) {
    if (i >= dataY.length) break;
    values.push(parseCustomVal(dataY[i][0]));
  }
  return { labels, values };
}

function getCategoryOverviewDashboardData() {
  return {
    breakoutReach: processPerformanceTiers('All Breakout Hrs', 'A1:A12', 'C1:C12'),
    breakoutHours: processPerformanceTiers('All Breakout % Reach', 'A20:A33', 'B20:B33'),
    premiumReach: processPerformanceTiers('All Premium % Reach', 'A1:A12', 'C1:C12'),
    premiumHours: processPerformanceTiers('All Premium Hrs', 'A20:A33', 'B20:B33'),
    libraryReach: processPerformanceTiers('All Library % Reach', 'A1:A12', 'C1:C12'),
    libraryHours: processPerformanceTiers('All Library Hrs', 'A20:A33', 'B20:B33')
  };
}

function getCleanPerformanceData() {
  const ss = SpreadsheetApp.openById(EXTERNAL_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Overall Perfomance ITD');
  const rawData = sheet.getDataRange().getDisplayValues();
  
  if (rawData.length === 0) return [];
  
  const headers = rawData[0];
  let reachIndex = -1;
  
  for (let j = 0; j < headers.length; j++) {
    const h = headers[j].toLowerCase().trim();
    if (h.includes('reach') && h.includes('%')) {
      reachIndex = j;
      break;
    }
  }
  if (reachIndex === -1) {
    for (let j = 0; j < headers.length; j++) {
      if (headers[j].toLowerCase().trim().includes('reach')) {
        reachIndex = j;
        break;
      }
    }
  }
  
  const dataRows = rawData.slice(1);
  
  if (reachIndex !== -1) {
    dataRows.sort((a, b) => {
      const valA = parseFloat(a[reachIndex].replace('%', '').trim()) || 0;
      const valB = parseFloat(b[reachIndex].replace('%', '').trim()) || 0;
      return valB - valA;
    });
  }
  
  const top20 = dataRows.slice(0, 20);
  return [headers].concat(top20);
}

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
      
      parsedPoints.push({
        day: interval,
        fact: reachFact
      });
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
  
  return {
    title: selectedTitle,
    category: titleCategory,
    target7D: target7D,
    target28D: target28D,
    daysElapsed: daysElapsed,
    currentRate: currentActualRate,
    neededDailyRate: neededDailyUplift,
    status: statusComment,
    chartData: finalTrajectoryArr
  };
}
