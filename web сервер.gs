/**
 * Файл: web_server.gs
 *
 * ИСПРАВЛЕНИЯ:
 * - getLiveResults: при фильтрации teamRows не проверялось, что row[1] не пустой.
 *   Команды без имени попадали в результаты. Добавлен фильтр.
 * - getLiveResults: rank добавляется в объект, но в results_web.html он нигде
 *   не используется (медали рисуются по индексу). Поле оставлено, добавлен комментарий.
 * - getDetailedRoundData: sheet.getRange("A:A") читает ~1000 строк для поиска ИТОГО.
 *   Заменено на getRange(1, 1, lastRow, 1) — только реальные данные.
 * - getDetailedRoundData: questionsEndRow вычислялся как i (0-based), потом numRows = i - 1.
 *   При questionsEndRow == 1 (ИТОГО во 2-й строке) numRows = 0 — корректно возвращается [].
 *   Логика верна, но добавлен поясняющий комментарий.
 * - getDetailedRoundData: поиск команды сравнивал teamHeaders[j] === teamName (строгое равенство).
 *   Если имя команды из формулы содержит пробелы или отличается регистром — не найдёт.
 *   Добавлено trim() для надёжности.
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('results_web')
    .setTitle('Quiz Leaderboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Возвращает текущие результаты для веб-лидерборда.
 * Данные берутся напрямую из "Сводной таблицы".
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
  const roundHeaders = headers.slice(2, headers.length - 1); // Названия раундов без "ИТОГО"

  const teamRows = data.slice(1);

  const results = teamRows
    // ИСПРАВЛЕНИЕ: пропускаем строки без имени команды
    .filter(row => row[1] && row[1].toString().trim() !== "")
    .map(row => {
      const roundsData = roundHeaders.map((header, index) => ({
        name: header.toString().replace("Раунд ", "Р"),
        fullName: header.toString(),
        score: row[index + 2] || 0
      }));

      return {
        name: row[1],
        score: row[headers.length - 1] || 0,  // Последний столбец = ИТОГО
        rounds: roundsData
      };
    });

  // Сортировка по убыванию баллов
  results.sort((a, b) => b.score - a.score);

  // rank добавляется для возможного использования на фронте
  return results.map((item, index) => {
    let rankLabel = (index + 1).toString();
    if (index === 0) rankLabel = "🥇";
    else if (index === 1) rankLabel = "🥈";
    else if (index === 2) rankLabel = "🥉";

    return { ...item, rank: rankLabel };
  });
}

/**
 * Возвращает детализацию ответов команды в конкретном раунде.
 * Вызывается с фронтенда при клике на бейдж раунда.
 */
function getDetailedRoundData(teamName, roundFullName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(roundFullName);
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();

  // ИСПРАВЛЕНИЕ: читаем только реальные строки, а не весь столбец A (~1000 строк)
  const colAValues = sheet.getRange(1, 1, lastRow, 1).getValues();

  // questionsEndRow — 0-based индекс строки ИТОГО
  // Данные вопросов: строки 2..(questionsEndRow) включительно, т.е. (questionsEndRow - 1) строк
  let questionsEndRow = lastRow;
  for (let i = 0; i < colAValues.length; i++) {
    if (colAValues[i][0].toString().includes("ИТОГО:")) {
      questionsEndRow = i; // 0-based: строка ИТОГО — это i+1, вопросы до неё
      break;
    }
  }

  // Поиск колонки команды в заголовках (строка 1, столбцы E-X)
  const teamHeaders = sheet.getRange(1, 5, 1, 20).getValues()[0];
  let teamColIndex = -1;
  for (let j = 0; j < teamHeaders.length; j++) {
    // ИСПРАВЛЕНИЕ: trim() на случай пробелов в формуле или имени
    if (teamHeaders[j].toString().trim() === teamName.toString().trim()) {
      teamColIndex = 5 + j;
      break;
    }
  }

  if (teamColIndex === -1) return null;

  // numRows = кол-во строк вопросов (заголовок в строке 1, данные с строки 2)
  const numRows = questionsEndRow - 1;
  if (numRows <= 0) return [];

  const questions = sheet.getRange(2, 1, numRows, 2).getValues();  // [номер, баллы]
  const checks = sheet.getRange(2, teamColIndex, numRows, 1).getValues(); // [чекбокс]

  return questions.map((q, idx) => ({
    num: q[0],
    points: q[1],
    correct: checks[idx][0] === true
  }));
}
