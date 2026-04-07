/**
 * Функция обработки GET-запроса (развертывание как веб-приложение)
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('results_web')
    .setTitle('Quiz Leaderboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Функция, которую вызывает Frontend для получения данных.
 * Берет данные напрямую из "Сводной таблицы".
 */
function getLiveResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("Сводная таблица");
  if (!summarySheet) return [];

  const lastRow = summarySheet.getLastRow();
  const lastCol = summarySheet.getLastColumn();
  if (lastRow < 2) return [];

  const data = summarySheet.getRange(1, 1, lastRow, lastCol).getValues();
  
  const headers = data[0];
  const roundHeaders = headers.slice(2, lastCol - 1);

  const teamRows = data.slice(1);
  
  const results = teamRows.map(row => {
    const roundsData = [];
    roundHeaders.forEach((header, index) => {
      roundsData.push({
        name: header.replace("Раунд ", "Р"),
        fullName: header,
        score: row[index + 2] || 0
      });
    });

    return {
      name: row[1],
      score: row[lastCol - 1],
      rounds: roundsData
    };
  });

  results.sort((a, b) => b.score - a.score);

  return results.map((item, index) => {
    let rankLabel = (index + 1).toString();
    if (index === 0) rankLabel = "🥇";
    else if (index === 1) rankLabel = "🥈";
    else if (index === 2) rankLabel = "🥉";
    
    return { ...item, rank: rankLabel };
  });
}

/**
 * Получение детальных ответов для конкретной команды в конкретном раунде
 */
function getDetailedRoundData(teamName, roundFullName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(roundFullName);
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  const colA = sheet.getRange("A:A").getValues();
  let questionsEndRow = lastRow;
  for (let i = 0; i < colA.length; i++) {
    if (colA[i][0].toString().includes("ИТОГО:")) {
      questionsEndRow = i;
      break;
    }
  }

  const teamHeaders = sheet.getRange(1, 5, 1, 20).getValues()[0];
  let teamColIndex = -1;
  for (let j = 0; j < teamHeaders.length; j++) {
    if (teamHeaders[j] === teamName) {
      teamColIndex = 5 + j;
      break;
    }
  }

  if (teamColIndex === -1) return null;

  const numRows = questionsEndRow - 1;
  if (numRows <= 0) return [];

  const questions = sheet.getRange(2, 1, numRows, 2).getValues();
  const checks = sheet.getRange(2, teamColIndex, numRows, 1).getValues();

  return questions.map((q, idx) => ({
    num: q[0],
    points: q[1],
    correct: checks[idx][0] === true
  }));
}

// =====================================================
// УПРАВЛЕНИЕ АВТООБНОВЛЕНИЕМ — вызывается из меню таблицы
// =====================================================

function enableWebRefresh() {
  PropertiesService.getScriptProperties().setProperty('WEB_AUTO_REFRESH', 'true');
  SpreadsheetApp.getUi().alert('✅ Автообновление LIVE включено!\n\nВеб-страница начнёт обновляться в течение 5 секунд.');
}

function disableWebRefresh() {
  PropertiesService.getScriptProperties().setProperty('WEB_AUTO_REFRESH', 'false');
  SpreadsheetApp.getUi().alert('❌ Автообновление LIVE отключено.\n\nВеб-страница остановит обновление в течение 5 секунд.');
}

// =====================================================
// API ДЛЯ ВЕБ-МОРДЫ (вызываются через google.script.run)
// =====================================================

/**
 * Проверяет пароль администратора.
 * Пароль хранится на листе WEB в ячейке B1.
 * Возвращает true если пароль верный, false если нет.
 */
function checkAdminPassword(inputPassword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const webSheet = ss.getSheetByName("WEB");
  if (!webSheet) return false;
  const correctPassword = webSheet.getRange("B1").getValue().toString();
  return inputPassword.toString() === correctPassword;
}

/**
 * Возвращает текущие настройки автообновления.
 * Вызывается со страницы каждые 5 секунд (polling).
 */
function getRefreshSettings() {
  const props = PropertiesService.getScriptProperties();
  const enabled = props.getProperty('WEB_AUTO_REFRESH') !== 'false';
  const interval = parseInt(props.getProperty('WEB_REFRESH_INTERVAL') || '15', 10);
  return { enabled: enabled, interval: interval };
}

/**
 * Сохраняет настройки автообновления из админ-панели веб-морды.
 * enabled: true/false
 * interval: интервал в секундах
 */
function saveRefreshSettings(enabled, interval) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('WEB_AUTO_REFRESH', enabled ? 'true' : 'false');
  const safeInterval = Math.max(5, Math.min(300, parseInt(interval, 10) || 15));
  props.setProperty('WEB_REFRESH_INTERVAL', safeInterval.toString());
  return { enabled: enabled, interval: safeInterval };
}
