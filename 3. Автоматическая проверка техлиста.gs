/**
 * Файл: auto_sync.gs
 * Ответственный за отслеживание изменений в составе листов и обновление Сводной таблицы.
 */

function setupAutoUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  
  // Очистка старых триггеров
  triggers.forEach(t => { 
    const func = t.getHandlerFunction();
    if (func === 'autoUpdateDashboard' || func === 'onEdit') ScriptApp.deleteTrigger(t); 
  });
  
  // Создаем триггер на изменение структуры таблиц
  ScriptApp.newTrigger('autoUpdateDashboard')
    .forSpreadsheet(ss)
    .onChange()
    .create();
    
  SpreadsheetApp.getUi().alert("✅ Автообновление включено.");
}

/**
 * Стандартная функция Google Apps Script, срабатывает при редактировании ячеек.
 */
function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const sheetName = range.getSheet().getName();
  
  // 1. Если правим данные в раундах — пересчитываем результаты
  if (sheetName.startsWith("Раунд ")) {
    if (typeof createResultsDashboard === 'function') {
      createResultsDashboard();
    }
  }
  
  // 2. Если правим "Список команд" (столбец B, начиная со 2 строки)
  // Это решит проблему, когда команды удалены, но сводная таблица не обновилась.
  if (sheetName === "Список команд" && range.getColumn() === 2 && range.getRow() >= 2) {
    refreshAllData();
  }
}

/**
 * Обработчик изменений структуры (добавление/удаление/переименование листов)
 */
function autoUpdateDashboard(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let techSheet = ss.getSheetByName("Техлист");
  
  // Если техлиста нет, создаем его и запускаем обновление
  if (!techSheet) {
    updateTechSheetList();
    refreshAllData();
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
    updateTechSheetList();
    refreshAllData();
  }
}

/**
 * Вспомогательная функция для полного обновления данных
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

  // Сортировка раундов по номеру
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
    const teamColLetter = columnToLetter(5 + tIdx); // Данные в раундах начинаются с E (5)
    
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
  
  // Очистка старой области перед вставкой новых данных
  summarySheet.clear(); 
  summarySheet.getRange(1, 1, rows, cols).setValues(newData);
  
  // Применяем стили
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