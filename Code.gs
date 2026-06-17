
const SPREADSHEET_ID = '1Y2hMSDU9BTlk_nHwvxWL5KmGQMA5QnD3x0et-gebdko';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sheet Viewer');
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

  // Assuming headers are in row 40, data starts in 41
  const data = sheet.getRange('F40:K44').getDisplayValues(); 
  const headers = data[0];

  const titleCol = headers.indexOf('Title'); // F
  const fact7Col = headers.indexOf('7D Fact'); // G
  const fact28Col = headers.indexOf('28D Fact'); // H
  const goal7Col = headers.indexOf('7D Goal'); // I
  const goal28Col = headers.indexOf('28D Goal'); // J
  const currentFactCol = headers.indexOf('Current % reach'); // K

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
  // Handle different current fact column names
  const currentFactCol = headers.indexOf('Current % reach') > -1 ? headers.indexOf('Current % reach') : headers.indexOf('Current Hours');

  const titles = [], fact7 = [], fact28 = [], currentFact = [], goal7 = [], goal28 = [];

  function parseVal(v) {
    if (!v) return 0;
    const str = v.toString().trim();
    if (str.includes('%')) return parseFloat(str.replace('%', '')) / 100;
    // Replace commas if they exist in numbers (e.g., 1,000)
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

// THIS IS THE FUNCTION THE FRONT-END NOW CALLS
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
  
  // Filter by Column E (index 4)
  const rows = data.slice(1).filter(row => row[4] === title);
  
  // Get header info from first matching row
  const headerInfo = rows.length > 0 ? {
    fieldE: rows[0][4],
    fieldH: rows[0][7],
    fieldJ: rows[0][9],
    fieldAA: rows[0][26]
  } : {};

  // Map data and sort by Column Y (index 24) Descending
  const tableData = rows.map(r => [r[4], r[21], r[22], r[23], r[24], r[25]])
                        .sort((a, b) => parseFloat(b[4]) - parseFloat(a[4]));

  return {
    headers: ['Title', 'CONTAINER_TITLE', 'IMPRESSION_BEHAVIOUR_EDITORIAL', 'IMPRESSION_BEHAVIOUR_PERSONALISATION', 'IMPRESSIONS', 'PLAYS'],
    rows: tableData,
    summary: headerInfo
  };
}
