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
 * Теперь берет данные напрямую из "Сводной таблицы", игнорируя визуальный лист "Результаты".
 */
function getLiveResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("Сводная таблица");
  if (!summarySheet) return [];

  const lastRow = summarySheet.getLastRow();
  const lastCol = summarySheet.getLastColumn();
  if (lastRow < 2) return [];

  // Получаем все данные из сводной таблицы
  const data = summarySheet.getRange(1, 1, lastRow, lastCol).getValues();
  
  // Заголовки (1 строка): [ "№", "Команда", "Раунд 1", ..., "ИТОГО" ]
  const headers = data[0];
  const roundHeaders = headers.slice(2, lastCol - 1); // Только названия раундов

  // Данные команд (начиная со 2 строки)
  const teamRows = data.slice(1);
  
  // Формируем массив объектов для фронтенда
  const results = teamRows.map(row => {
    const roundsData = [];
    // Собираем баллы за каждый раунд
    roundHeaders.forEach((header, index) => {
      roundsData.push({
        name: header.replace("Раунд ", "Р"),
        fullName: header, // Для точного поиска листа
        score: row[index + 2] || 0
      });
    });

    return {
      name: row[1],             // Имя команды
      score: row[lastCol - 1],  // Итоговый балл
      rounds: roundsData        // Детализация по раундам
    };
  });

  // Сортируем по убыванию баллов
  results.sort((a, b) => b.score - a.score);

  // Добавляем места (ранги) после сортировки
  return results.map((item, index) => {
    let rankLabel = (index + 1).toString();
    if (index === 0) rankLabel = "🥇";
    else if (index === 1) rankLabel = "🥈";
    else if (index === 2) rankLabel = "🥉";
    
    return {
      ...item,
      rank: rankLabel
    };
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
  // Ищем строку "ИТОГО:" чтобы понять, где заканчиваются вопросы
  const colA = sheet.getRange("A:A").getValues();
  let questionsEndRow = lastRow;
  for (let i = 0; i < colA.length; i++) {
    if (colA[i][0].toString().includes("ИТОГО:")) {
      questionsEndRow = i; // Индекс 0-based соответствует номеру строки (i+1), но нам нужны данные ДО этой строки
      break;
    }
  }

  // Названия команд в 1-й строке (с 5-го столбца)
  const teamHeaders = sheet.getRange(1, 5, 1, 20).getValues()[0];
  let teamColIndex = -1;
  for (let j = 0; j < teamHeaders.length; j++) {
    if (teamHeaders[j] === teamName) {
      teamColIndex = 5 + j;
      break;
    }
  }

  if (teamColIndex === -1) return null;

  // Получаем номера вопросов (A), баллы (B) и чекбоксы (teamColIndex)
  const numRows = questionsEndRow - 1; // Минус шапка
  if (numRows <= 0) return [];

  const questions = sheet.getRange(2, 1, numRows, 2).getValues();
  const checks = sheet.getRange(2, teamColIndex, numRows, 1).getValues();

  return questions.map((q, idx) => ({
    num: q[0],
    points: q[1],
    correct: checks[idx][0] === true
  }));
}