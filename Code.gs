
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

  // ✅ Only pull required range
  const data = sheet.getRange('F40:K44').getDisplayValues();

  const headers = data[0];

  const titleCol = headers.indexOf('Title');
  const fact7Col = headers.indexOf('7D Fact');
  const fact28Col = headers.indexOf('28D Fact');
  const goal7Col = headers.indexOf('7D Goal');
  const goal28Col = headers.indexOf('28D Goal');
  const reachCol = headers.indexOf('Current % reach');

  const titles = [];
  const fact7 = [];
  const fact28 = [];
  const reach = [];
  const goal28 = [];

  function parsePercent(v) {
    if (!v) return 0;

    const str = v.toString().trim();

    if (str.includes('%')) {
      return parseFloat(str.replace('%', '')) / 100;
    }

    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  }

  // ✅ LOOP FIXED (< instead of &lt;)
  for (let i = 1; i < data.length; i++) {

    if (!data[i][titleCol]) continue; // skip empty rows

    titles.push(data[i][titleCol]);

    fact7.push(parsePercent(data[i][fact7Col]));
    fact28.push(parsePercent(data[i][fact28Col]));
    reach.push(parsePercent(data[i][reachCol]));
    goal28.push(parsePercent(data[i][goal28Col]));
  }

  return { titles, fact7, fact28, reach, goal28 };
}
