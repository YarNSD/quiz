/**
 * Файл: results_manager.gs
 * Визуализация Дашборда. 
 * Изменено: таблица начинается с A1, все раунды сгруппированы.
 */

function createResultsDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("Сводная таблица");
  
  if (!summarySheet) return;

  const sheetName = "Результаты";
  let resultsSheet = ss.getSheetByName(sheetName);
  if (!resultsSheet) {
    resultsSheet = ss.insertSheet(sheetName);
  }

  // 1. ПОЛНАЯ ОЧИСТКА И СБРОС ГРУПП
  resultsSheet.setFrozenRows(0);
  resultsSheet.setFrozenColumns(0);
  
  const maxCols = resultsSheet.getMaxColumns();
  if (maxCols > 0) {
    try { 
      // Разгруппировываем всё, что было (до 8 уровней)
      for (let i = 0; i < 8; i++) {
        resultsSheet.getRange(1, 1, 1, maxCols).shiftColumnGroupDepth(-1); 
      }
    } catch (e) {}
  }
  
  resultsSheet.clear();
  // Заливка всего листа темным фоном
  resultsSheet.getRange(1, 1, resultsSheet.getMaxRows(), resultsSheet.getMaxColumns()).setBackground("#1a1a1a");

  // 2. ПОДГОТОВКА ДАННЫХ
  const lastRowSummary = summarySheet.getLastRow();
  const lastColSummary = summarySheet.getLastColumn();
  if (lastRowSummary < 2) return;

  const summaryData = summarySheet.getRange(1, 1, lastRowSummary, lastColSummary).getValues();
  const roundHeaders = summaryData[0].slice(2, lastColSummary - 1);
  const roundCount = roundHeaders.length;
  const tableWidth = roundCount + 3; // Медаль + Команда + Раунды + Итого

  const rawData = summaryData.slice(1).map(row => row.slice(1)); 
  const activeTeamsData = rawData.filter(row => row[0].toString().trim() !== "" && row[0].toString().trim() !== "0");
  const sortedData = activeTeamsData.sort((a, b) => b[b.length - 1] - a[a.length - 1]);
  const teamCount = sortedData.length;

  // 3. ФОРМИРОВАНИЕ МАССИВОВ ДЛЯ ПЕЧАТИ
  const displayValues = [];
  const backgroundColors = [];
  
  // Строка заголовка (Объединяем в первой строке)
  resultsSheet.getRange(1, 1, 1, tableWidth).merge()
    .setValue("🏆 ОБЩИЙ ЗАЧЕТ КВИЗА 🏆")
    .setFontSize(28)
    .setFontWeight("bold")
    .setFontColor("#ffcc00")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Шапка таблицы (2-я строка)
  const headerRowValues = ["", "КОМАНДА", ...roundHeaders.map(h => h.replace("Раунд ", "Р")), "ИТОГО"];
  displayValues.push(headerRowValues);
  backgroundColors.push(new Array(tableWidth).fill("#1a1a1a"));

  // Данные команд (с 3-й строки)
  sortedData.forEach((row, i) => {
    const medal = i === 0 ? "🥇" : i === 1 ? "🥈" : i === 2 ? "🥉" : "";
    const teamName = row[0];
    const rounds = row.slice(1, 1 + roundCount);
    const total = row[row.length - 1];
    
    displayValues.push([medal, teamName, ...rounds, total]);
    backgroundColors.push(new Array(tableWidth).fill(i < 3 ? "#2d2d2d" : "#1a1a1a"));
  });

  // 4. ЗАПИСЬ И СТИЛИЗАЦИЯ (начиная со второй строки, так как первая занята заголовком)
  const range = resultsSheet.getRange(2, 1, displayValues.length, tableWidth);
  range.setValues(displayValues);
  range.setBackgrounds(backgroundColors);
  range.setFontFamily("Rubik").setFontColor("#ffffff").setVerticalAlignment("middle");

  // Стили шапки (строка 2)
  resultsSheet.getRange(2, 1, 1, tableWidth).setFontWeight("bold").setHorizontalAlignment("center");
  resultsSheet.getRange(2, 2).setFontColor("#00e5ff").setHorizontalAlignment("left").setFontSize(14);
  resultsSheet.getRange(2, tableWidth).setFontColor("#ffcc00").setFontSize(14);

  // Стили тела
  if (teamCount > 0) {
    const dataRows = resultsSheet.getRange(3, 1, teamCount, tableWidth);
    dataRows.setBorder(null, null, true, null, null, null, "#444444", SpreadsheetApp.BorderStyle.SOLID);
    
    resultsSheet.getRange(3, 1, teamCount, 1).setFontSize(26).setHorizontalAlignment("center");
    resultsSheet.getRange(3, 2, teamCount, 1).setFontSize(16).setFontWeight("bold");
    resultsSheet.getRange(3, 3, teamCount, roundCount).setFontSize(14).setHorizontalAlignment("center");
    resultsSheet.getRange(3, tableWidth, teamCount, 1).setFontSize(22).setFontColor("#ffcc00").setFontWeight("bold").setHorizontalAlignment("center");

    // Рамка вокруг столбца ИТОГО
    resultsSheet.getRange(2, tableWidth, teamCount + 1, 1).setBorder(true, true, true, true, null, null, "#ffcc00", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  // 5. ГЕОМЕТРИЯ И ГРУППИРОВКА
  resultsSheet.setHiddenGridlines(true);
  
  resultsSheet.setColumnWidth(1, 60);  // Медаль
  resultsSheet.setColumnWidth(2, 350); // Команда
  resultsSheet.setColumnWidth(tableWidth, 120); // Итого
  
  resultsSheet.setRowHeight(1, 90); // Высота заголовка
  if (teamCount > 0) resultsSheet.setRowHeights(3, teamCount, 60); // Высота строк команд
  
  // Группировка раундов (столбцы между "Командой" и "Итого")
  if (roundCount > 0) {
    resultsSheet.setColumnWidths(3, roundCount, 75);
    const groupRange = resultsSheet.getRange(1, 3, 1, roundCount);
    
    try {
      groupRange.shiftColumnGroupDepth(1);
      resultsSheet.collapseAllColumnGroups(); // Скрываем раунды по умолчанию
    } catch (e) {
      console.log("Группировка уже существует или ошибка: " + e.message);
    }
  }
}