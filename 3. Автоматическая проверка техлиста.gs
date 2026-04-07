/**
 * Файл: auto_sync.gs
 * Ответственный за отслеживание изменений в составе листов, обновление Техлиста и Сводной таблицы.
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

  // Создаем триггер на редактирование (для списка команд)
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
    
  // Сразу обновляем технические данные при настройке
  refreshAllData();
    
  SpreadsheetApp.getUi().alert("✅ Автообновление включено. Техлист и Сводная таблица обновляются автоматически. Дашборд обновляется ТОЛЬКО вручную.");
}

/**
 * Автоматическое обновление при редактировании ячеек (например, в списке команд).
 */
function onEdit(e) {
  if (!e) return;
  const sheetName = e.range.getSheet().getName();
  
  // Если изменили список команд, обновляем сводную таблицу автоматически
  if (sheetName === "Список команд") {
    refreshAllData();
  }
}

/**
 * Обработчик изменений структуры (добавление/удаление/переименование листов)
 */
function autoUpdateDashboard(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Обновляем технический список листов
  updateTechSheetList();
  
  // 2. Обновляем только Сводную таблицу
  refreshAllData();
  
  console.log("Структура изменилась: Техлист и Сводная таблица обновлены.");
}

/**
 * Обновление списка листов в Техлисте (Диапазон A2:A)
 */
function updateTechSheetList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let techSheet = ss.getSheetByName("Техлист");
  
  if (!techSheet) {
    techSheet = ss.insertSheet("Техлист");
    techSheet.hideSheet();
  }
  
  const sheetNames = ss.getSheets().map(s => [s.getName()]);
  techSheet.clearContents();
  techSheet.getRange(1, 1).setValue("Список листов (служебный)").setFontWeight("bold");
  
  if (sheetNames.length > 0) {
    techSheet.getRange(2, 1, sheetNames.length, 1).setValues(sheetNames);
  }
}

/**
 * Функция для пересчета Сводной таблицы
 */
function refreshAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamListSheet = ss.getSheetByName("Список команд");
  const roundSheets = ss.getSheets().filter(s => s.getName().startsWith("Раунд "));
  
  if (teamListSheet && roundSheets.length > 0) {
    // Берем команды из столбца B
    const teams = teamListSheet.getRange("B2:B").getValues().filter(r => r[0] !== "");
    
    // Обновляем матрицу формул в Сводной таблице
    updateSummaryMatrix(teams, roundSheets);
    
    // ВАЖНО: Вызов createResultsDashboard() УДАЛЕН отсюда, 
    // чтобы дашборд не обновлялся автоматически.
  }
}

/**
 * Умное обновление "Сводной таблицы"
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

  // Сортируем раунды по номеру в названии
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
    const teamColLetter = columnToLetter(5 + tIdx); // Данные команд в раундах начинаются с E (5-й столбец)
    
    sortedRoundSheets.forEach(s => {
      // Динамическая формула для подтягивания ИТОГО из листа каждого раунда
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