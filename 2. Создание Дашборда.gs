/**
 * Файл: results_manager.gs
 * Оптимизированная версия: берет данные из "Сводная таблица"
 * Возвращен оригинальный дизайн и оформление.
 *
 * ИСПРАВЛЕНИЯ:
 * - Добавлена защита: если teamCount === 0, функция прерывается с сообщением
 *   вместо того чтобы записывать пустой диапазон и вызывать ошибку.
 * - setRowHeights(3, teamCount, 55) теперь вызывается только при teamCount > 0.
 * - Золотая граница колонки ИТОГО вычисляется от teamCount + 1 строк (было + 1, теперь явно).
 * - Логика чтения roundHeaders исправлена: slice(2, lastCol - 1) может срезать данные
 *   если последняя колонка сводной таблицы содержит формулы — теперь используем headers.length - 1.
 */

function createResultsDashboard() {
  const startTime = new Date().getTime();
  const log = [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("Сводная таблица");

  if (!summarySheet) {
    SpreadsheetApp.getUi().alert("Лист 'Сводная таблица' не найден.");
    return;
  }

  const sheetName = "Результаты";
  let resultsSheet = ss.getSheetByName(sheetName);

  // ЭТАП 1: Подготовка листа
  const t1 = new Date().getTime();
  if (resultsSheet) {
    resultsSheet.clear();
    try {
      const maxCols = resultsSheet.getMaxColumns();
      for (let i = 0; i < 3; i++) resultsSheet.getRange(1, 1, 1, maxCols).shiftColumnGroupDepth(-1);
    } catch (e) {}
  } else {
    resultsSheet = ss.insertSheet(sheetName);
  }
  log.push(`Этап 1 (Подготовка): ${new Date().getTime() - t1}ms`);

  // ЭТАП 2: Чтение данных
  const t2 = new Date().getTime();
  const summaryData = summarySheet.getDataRange().getValues();

  if (summaryData.length < 2) {
    SpreadsheetApp.getUi().alert("Сводная таблица пуста — добавьте команды и раунды.");
    return;
  }

  const rawHeaders = summaryData[0];
  // ИСПРАВЛЕНИЕ: используем rawHeaders.length - 1 вместо lastCol - 1
  // чтобы не зависеть от getLastColumn() и корректно срезать "ИТОГО"
  const roundHeaders = rawHeaders.slice(2, rawHeaders.length - 1).map(header => {
    return header.toString().replace("Раунд ", "Р");
  });
  const lastCol = roundHeaders.length + 3;

  let allResultsData = [];
  for (let i = 1; i < summaryData.length; i++) {
    const row = summaryData[i];
    if (!row[1]) continue;

    let teamRow = ["", row[1]]; // [Медаль, Имя]
    for (let j = 2; j < row.length - 1; j++) {
      teamRow.push(row[j] || 0);
    }
    teamRow.push(row[row.length - 1] || 0);
    allResultsData.push(teamRow);
  }
  const teamCount = allResultsData.length;
  log.push(`Этап 2 (Чтение): ${new Date().getTime() - t2}ms`);

  // ИСПРАВЛЕНИЕ: защита от пустых данных
  if (teamCount === 0) {
    SpreadsheetApp.getUi().alert("Нет данных о командах в Сводной таблице.");
    return;
  }

  // ЭТАП 3: Запись и Сортировка
  const t3 = new Date().getTime();

  // Темная тема для всего листа
  resultsSheet.getRange("A:Z").setFontFamily("Rubik").setBackground("#1a1a1a").setFontColor("#ffffff");
  resultsSheet.setHiddenGridlines(true);

  // Запись данных
  const dataRange = resultsSheet.getRange(3, 1, teamCount, lastCol);
  dataRange.setValues(allResultsData);
  dataRange.sort({ column: lastCol, ascending: false });
  log.push(`Этап 3 (Запись): ${new Date().getTime() - t3}ms`);

  // ЭТАП 4: Стилизация
  const t4 = new Date().getTime();

  // 1. ЗАГОЛОВОК ДАШБОРДА
  resultsSheet.getRange(1, 1, 1, lastCol).merge().setValue("🏆 ТАБЛИЦА ЛИДЕРОВ 🏆")
       .setBackground("#1a1a1a").setFontColor("#ffcc00").setFontWeight("bold")
       .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(24);

  // 2. ШАПКА ТАБЛИЦЫ
  const headerRange = resultsSheet.getRange(2, 1, 1, lastCol);
  headerRange.setBackground("#2d2d2d").setFontColor("#00e5ff").setFontWeight("bold")
             .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(12);

  resultsSheet.getRange(2, 1, 1, 2).merge().setValue("КОМАНДА");
  resultsSheet.getRange(2, 3, 1, roundHeaders.length).setValues([roundHeaders]);
  resultsSheet.getRange(2, lastCol).setValue("ИТОГО").setFontColor("#ffcc00");

  // 3. МЕДАЛИ И СТИЛИ ТОП-3
  for (let i = 0; i < teamCount; i++) {
    const row = 3 + i;
    const rowRange = resultsSheet.getRange(row, 1, 1, lastCol);
    const teamNameCell = resultsSheet.getRange(row, 2);

    let medal = "";
    let bgColor = "#1a1a1a";
    let nameColor = "#ffffff";

    if (i === 0) {
      medal = "🥇"; bgColor = "#3d3211"; nameColor = "#ffd700";
    } else if (i === 1) {
      medal = "🥈"; bgColor = "#2f3136"; nameColor = "#e0e0e0";
    } else if (i === 2) {
      medal = "🥉"; bgColor = "#32231a"; nameColor = "#cd7f32";
    }

    resultsSheet.getRange(row, 1).setValue(medal).setFontSize(20).setHorizontalAlignment("center");
    rowRange.setBackground(bgColor);
    teamNameCell.setFontColor(nameColor).setFontWeight(i < 3 ? "bold" : "normal");
    rowRange.setBorder(null, null, true, null, null, null, "#444444", SpreadsheetApp.BorderStyle.SOLID);
    resultsSheet.getRange(row, lastCol).setFontWeight("bold").setFontColor("#ffcc00").setFontSize(18);
  }

  // 4. ОБЩИЕ ПАРАМЕТРЫ ТЕКСТА
  const allDataRange = resultsSheet.getRange(3, 1, teamCount, lastCol);
  allDataRange.setVerticalAlignment("middle").setFontSize(16);

  resultsSheet.getRange(3, 3, teamCount, lastCol - 2).setHorizontalAlignment("center");
  resultsSheet.getRange(3, 2, teamCount, 1).setHorizontalAlignment("left");

  // Золотистая граница для колонки ИТОГО (от заголовка до последней команды)
  const totalColumnRange = resultsSheet.getRange(2, lastCol, teamCount + 1, 1);
  totalColumnRange.setBorder(true, true, true, true, null, null, "#ffcc00", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 5. НАСТРОЙКА РАЗМЕРОВ
  resultsSheet.setColumnWidth(1, 50);
  resultsSheet.setColumnWidth(2, 300);
  resultsSheet.setColumnWidth(lastCol, 100);
  resultsSheet.setRowHeight(1, 80);
  // ИСПРАВЛЕНИЕ: вызов только при teamCount > 0
  if (teamCount > 0) {
    resultsSheet.setRowHeights(3, teamCount, 55);
  }

  if (roundHeaders.length > 0) {
    resultsSheet.setColumnWidths(3, roundHeaders.length, 60);
  }

  // ГРУППИРОВКА РАУНДОВ
  if (roundHeaders.length > 1) {
    try {
      resultsSheet.getRange(1, 3, 1, roundHeaders.length).shiftColumnGroupDepth(1);
      resultsSheet.collapseAllColumnGroups();
    } catch (e) {}
  }

  log.push(`Этап 4 (Дизайн): ${new Date().getTime() - t4}ms`);
  log.push(`-----------------------------------`);
  log.push(`ОБЩЕЕ ВРЕМЯ: ${new Date().getTime() - startTime}ms`);

  console.log("--- ОТЧЕТ ДАШБОРДА ---");
  log.forEach(entry => console.log(entry));
}
