const SPREADSHEET_ID = '1Y2hMSDU9BTlk_nHwvxWL5KmGQMA5QnD3x0et-gebdko';
const EXTERNAL_SPREADSHEET_ID = '1KkGtDYCrXDBX4WTC5i9iwsKCBq6NeKUauwiGQ-C4Nb4';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sheet Viewer');
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

// ========================================================
// ========== NEW FUNCTIONS FOR TARGET EXTENSIONS =========
// ========================================================

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
