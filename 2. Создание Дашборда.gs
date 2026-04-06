/**
 * Файл: auto_sync.gs
 * Ответственный за отслеживание изменений в составе листов и обновление Техлиста.
 * Автоматическое обновление Сводной таблицы и Дашборда отключено по запросу.
 */

function setupAutoUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  
  // Очистка старых триггеров
  triggers.forEach(t => { 
    const func = t.getHandlerFunction();
    if (func === 'autoUpdateDashboard' || func === 'onEdit') ScriptApp.deleteTrigger(t); 
  });
  
  // Создаем триггер на изменение структуры таблиц (добавление/удаление листов)
  ScriptApp.newTrigger('autoUpdateDashboard')
    .forSpreadsheet(ss)
    .onChange()
    .create();
    
  SpreadsheetApp.getUi().alert("✅ Автообновление Техлиста включено. Сводная таблица теперь обновляется только вручную.");
}

/**
 * Стандартная функция Google Apps Script, срабатывает при редактировании ячеек.
 */
function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const sheetName = range.getSheet().getName();
  
  // Автоматическое обновление Дашборда и Сводной при правке ячеек ОТКЛЮЧЕНО.
}

/**
 * Обработчик изменений структуры (добавление/удаление/переименование листов)
 */
function autoUpdateDashboard(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let techSheet = ss.getSheetByName("Техлист");
  
  // Если техлиста нет, создаем его
  if (!techSheet) {
    updateTechSheetList();
    return;
  }
  
  const currentSheets = ss.getSheets().map(s => s.getName());
  const lastRow = techSheet.getLastRow();
  let oldSheets = [];
  
  // ЧИТАЕМ ИЗ A2:A (столбец 1)
  if (lastRow >= 2) {
    oldSheets = techSheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => r[0]).filter(String);
  }
  
  // Сравниваем списки листов
  const isChanged = currentSheets.length !== oldSheets.length || 
                    currentSheets.some((name, i) => name !== oldSheets[i]);
  
  if (isChanged) {
    // ОБНОВЛЕНИЕ ТЕХЛИСТА ОСТАВЛЕНО
    updateTechSheetList();
    
    // ОБНОВЛЕНИЕ ДАННЫХ ОТКЛЮЧЕНО (refreshAllData)
    console.log("Структура изменилась: Техлист обновлен. Сводная таблица не затронута.");
  }
}

/**
 * Вспомогательная функция для ПОЛНОГО обновления данных (теперь только для ручного запуска)
 */
function refreshAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamListSheet = ss.getSheetByName("Список команд");
  const roundSheets = ss.getSheets().filter(s => s.getName().startsWith("Раунд "));
  
  if (teamListSheet && roundSheets.length > 0) {
    const teams = teamListSheet.getRange("B2:B").getValues().filter(r => r[0] !== "");
    updateSummaryMatrix(teams, roundSheets);
    if (typeof createResultsDashboard === 'function') {
      createResultsDashboard();
    }
  }
}

/**
 * Обновление списка листов в Техлисте (Диапазон A2:A)
 * ЭТА ФУНКЦИЯ РАБОТАЕТ АВТОМАТИЧЕСКИ
 */
function updateTechSheetList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let techSheet = ss.getSheetByName("Техлист");
  
  if (!techSheet) {
    techSheet = ss.insertSheet("Техлист");
    techSheet.hideSheet();
  }
  
  const sheetNames = ss.getSheets().map(s => [s.getName()]);
  
  // Полная очистка столбца A перед записью
  techSheet.getRange("A:A").clearContent();
  techSheet.getRange(1, 1).setValue("Список листов (служебный)").setFontWeight("bold");
  
  if (sheetNames.length > 0) {
    techSheet.getRange(2, 1, sheetNames.length, 1).setValues(sheetNames);
  }
}

/**
 * Умное обновление "Сводной таблицы" (Вызывается только через refreshAllData)
 */
function updateSummaryMatrix(teams, roundSheets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let summarySheet = ss.getSheetByName("Сводная таблица");
  
  if (!summarySheet) {
    summarySheet = ss.insertSheet("Сводная таблица");
    summarySheet.getRange("A:Z").setFontFamily("Rubik").setVerticalAlignment("middle");
    summarySheet.setHiddenGridlines(true);
  }

  if (teams.length === 0 || roundSheets.length === 0) {
    summarySheet.clear();
    return;
  }

  const sortedRoundSheets = roundSheets.slice().sort((a, b) => {
    const numA = parseInt(a.getName().replace(/\D/g, '')) || 0;
    const numB = parseInt(b.getName().replace(/\D/g, '')) || 0;
    return numA - numB;
  });

  let headerRow = ["№", "Команда"];
  sortedRoundSheets.forEach(s => headerRow.push(s.getName()));
  headerRow.push("ИТОГО");

  let newData = [headerRow];
  teams.forEach((team, tIdx) => {
    let row = [tIdx + 1, team[0]];
    const teamColLetter = columnToLetter(5 + tIdx);
    
    sortedRoundSheets.forEach(s => {
      row.push(`=IFERROR(INDIRECT("'${s.getName()}'!${teamColLetter}" & MATCH("ИТОГО:"; '${s.getName()}'!$A:$A; 0)); 0)`);
    });
    
    const startCol = columnToLetter(3);
    const endCol = columnToLetter(2 + sortedRoundSheets.length);
    row.push(`=SUM(${startCol}${tIdx + 2}:${endCol}${tIdx + 2})`);
    newData.push(row);
  });

  const rows = newData.length;
  const cols = headerRow.length;
  
  summarySheet.clear(); 
  summarySheet.getRange(1, 1, rows, cols).setValues(newData);
  applySummaryStyles(summarySheet, rows, cols, sortedRoundSheets.length);
}

/**
 * ФУНКЦИЯ ДЛЯ СОЗДАНИЯ ДАШБОРДА (Результаты)
 * Изменено: отображаются только раунды с ненулевыми очками.
 */
function createResultsDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("Сводная таблица");
  if (!summarySheet) return;

  const sheetName = "Результаты";
  let resultsSheet = ss.getSheetByName(sheetName);
  if (resultsSheet) {
    resultsSheet.clear();
    try {
      const maxCols = resultsSheet.getMaxColumns();
      for (let i = 0; i < 3; i++) resultsSheet.getRange(1, 1, 1, maxCols).shiftColumnGroupDepth(-1);
    } catch (e) {}
  } else {
    resultsSheet = ss.insertSheet(sheetName);
  }

  const summaryData = summarySheet.getDataRange().getValues();
  const rawHeaders = summaryData[0];
  const allRows = summaryData.slice(1).filter(r => r[1]); // Только строки с именами команд

  // Определяем, в каких раундах есть хотя бы одно значение > 0
  // Колонки раундов в Сводной начинаются с индекса 2 до (length - 2)
  const activeRoundIndices = [];
  for (let col = 2; col < rawHeaders.length - 1; col++) {
    const hasPoints = allRows.some(row => (parseFloat(row[col]) || 0) > 0);
    if (hasPoints) {
      activeRoundIndices.push(col);
    }
  }

  // Заголовки для активных раундов (Р1, Р2...)
  const roundHeaders = activeRoundIndices.map(idx => {
    return rawHeaders[idx].toString().replace("Раунд ", "Р");
  });
  const lastCol = roundHeaders.length + 3; // Медаль + Имя + Раунды + Итого

  let allResultsData = [];
  allRows.forEach(row => {
    let teamRow = ["", row[1]]; // [Медаль, Имя]
    activeRoundIndices.forEach(idx => {
      teamRow.push(row[idx] || 0);
    });
    teamRow.push(row[row.length - 1] || 0); // ИТОГО
    allResultsData.push(teamRow);
  });

  const teamCount = allResultsData.length;
  if (teamCount === 0) return;

  // Оформление
  resultsSheet.getRange("A:Z").setFontFamily("Rubik").setBackground("#1a1a1a").setFontColor("#ffffff");
  resultsSheet.setHiddenGridlines(true);

  const dataRange = resultsSheet.getRange(3, 1, teamCount, lastCol);
  dataRange.setValues(allResultsData);
  dataRange.sort({column: lastCol, ascending: false});

  // Заголовок
  resultsSheet.getRange(1, 1, 1, lastCol).merge().setValue("🏆 ТАБЛИЦА ЛИДЕРОВ 🏆")
       .setBackground("#1a1a1a").setFontColor("#ffcc00").setFontWeight("bold")
       .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(24);

  // Шапка
  resultsSheet.getRange(2, 1, 1, 2).merge().setValue("КОМАНДА");
  if (roundHeaders.length > 0) {
    resultsSheet.getRange(2, 3, 1, roundHeaders.length).setValues([roundHeaders]);
  }
  resultsSheet.getRange(2, lastCol).setValue("ИТОГО").setFontColor("#ffcc00");
  resultsSheet.getRange(2, 1, 1, lastCol).setBackground("#2d2d2d").setFontColor("#00e5ff").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontSize(12);

  // Стили строк и Медали
  for (let i = 0; i < teamCount; i++) {
    const row = 3 + i;
    const rowRange = resultsSheet.getRange(row, 1, 1, lastCol);
    let medal = "";
    let bgColor = "#1a1a1a";
    let nameColor = "#ffffff";

    if (i === 0) { medal = "🥇"; bgColor = "#3d3211"; nameColor = "#ffd700"; }
    else if (i === 1) { medal = "🥈"; bgColor = "#2f3136"; nameColor = "#e0e0e0"; }
    else if (i === 2) { medal = "🥉"; bgColor = "#32231a"; nameColor = "#cd7f32"; }

    resultsSheet.getRange(row, 1).setValue(medal).setFontSize(20).setHorizontalAlignment("center");
    rowRange.setBackground(bgColor);
    resultsSheet.getRange(row, 2).setFontColor(nameColor).setFontWeight(i < 3 ? "bold" : "normal");
    rowRange.setBorder(null, null, true, null, null, null, "#444444", SpreadsheetApp.BorderStyle.SOLID);
    resultsSheet.getRange(row, lastCol).setFontWeight("bold").setFontColor("#ffcc00").setFontSize(18);
  }

  const allDataRange = resultsSheet.getRange(3, 1, teamCount, lastCol);
  allDataRange.setVerticalAlignment("middle").setFontSize(16);
  resultsSheet.getRange(3, 3, teamCount, lastCol - 2).setHorizontalAlignment("center");
  resultsSheet.getRange(3, 2, teamCount, 1).setHorizontalAlignment("left");
  resultsSheet.getRange(2, lastCol, teamCount + 1, 1).setBorder(true, true, true, true, null, null, "#ffcc00", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  resultsSheet.setColumnWidth(1, 50);
  resultsSheet.setColumnWidth(2, 300);
  resultsSheet.setColumnWidth(lastCol, 100);
  resultsSheet.setRowHeight(1, 80);
  resultsSheet.setRowHeights(3, teamCount, 55);
  if (roundHeaders.length > 0) resultsSheet.setColumnWidths(3, roundHeaders.length, 60);

  if (roundHeaders.length > 1) {
    try {
      resultsSheet.getRange(1, 3, 1, roundHeaders.length).shiftColumnGroupDepth(1);
      resultsSheet.collapseAllColumnGroups();
    } catch(e) {}
  }
}

function applySummaryStyles(summarySheet, lastRow, lastCol, roundCount) {
  summarySheet.getRange(1, 1, 1, lastCol).setBackground("#444444").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  summarySheet.getRange(2, 1, lastRow - 1, 1).setHorizontalAlignment("center").setFontColor("#999999");
  summarySheet.getRange(2, 2, lastRow - 1, 1).setFontWeight("bold");
  summarySheet.getRange(2, 3, lastRow - 1, lastCol - 2).setHorizontalAlignment("center");
  summarySheet.getRange(2, lastCol, lastRow - 1, 1).setBackground("#ffffff").setFontColor("#000000").setFontWeight("bold").setHorizontalAlignment("center").setBorder(null, true, null, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  summarySheet.setColumnWidth(1, 35);
  summarySheet.setColumnWidth(2, 250);
  if (roundCount > 0) summarySheet.setColumnWidths(3, roundCount, 85);
  summarySheet.setColumnWidth(lastCol, 100);
  summarySheet.setFrozenRows(1);
  summarySheet.setFrozenColumns(2);
}

function columnToLetter(column) {
  let temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}