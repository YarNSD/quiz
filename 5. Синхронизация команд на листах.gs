/**
 * Файл: sync_logic.gs
 * Описание: Синхронизация команд и управление автоматизацией.
 */

/**
 * Основная функция синхронизации команд.
 * Обновляет имена, сбрасывает пустые чекбоксы и доп. очки, не скрывая столбцы.
 */
function syncTeamsAcrossRounds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const rounds = allSheets.filter(s => s.getName().includes("Раунд"));
  const teamListSheet = ss.getSheetByName("Список команд");
  
  if (!teamListSheet) {
    SpreadsheetApp.getUi().alert("Ошибка: Лист 'Список команд' не найден.");
    return;
  }

  // Получаем актуальный список команд (столбец B)
  const teams = teamListSheet.getRange("B2:B21").getValues().map(r => r[0]);
  const maxTeams = 20;

  rounds.forEach(sheet => {
    const maxRows = sheet.getMaxRows();
    if (maxRows < 2) return;

    // 1. Поиск строки бонусов для очистки
    const searchRangeHeight = Math.min(sheet.getLastRow(), 100);
    const colAValues = sheet.getRange(1, 1, searchRangeHeight, 1).getValues();
    let bonusRowIndex = -1;
    for (let i = 0; i < colAValues.length; i++) {
      if (colAValues[i][0].toString().includes("Доп. баллы")) {
        bonusRowIndex = i + 1;
        break;
      }
    }

    // Подготовка массивов для пакетной обработки
    const rangesToReset = [];
    const newHeaders = [teams]; // Для обновления имен команд в первой строке

    // 2. Сбор ячеек для сброса (чекбоксы + бонусы)
    for (let t = 0; t < maxTeams; t++) {
      const colIndex = 5 + t; // Столбцы E-X
      const teamName = teams[t] ? teams[t].toString().trim() : "";

      if (teamName === "") {
        // Если команды нет — сбрасываем чекбоксы (строка 2 до строки бонусов)
        const endRow = bonusRowIndex !== -1 ? bonusRowIndex - 1 : maxRows;
        if (endRow >= 2) {
          rangesToReset.push(sheet.getRange(2, colIndex, endRow - 1, 1).getA1Notation());
        }
        
        if (bonusRowIndex !== -1) {
          rangesToReset.push(sheet.getRange(bonusRowIndex, colIndex).getA1Notation());
        }
      }
    }

    // 3. Исполнение изменений
    sheet.getRange(1, 5, 1, maxTeams).setValues(newHeaders);
    sheet.showColumns(5, maxTeams);

    if (rangesToReset.length > 0) {
      sheet.getRangeList(rangesToReset).uncheck().clearContent();
    }
  });
  
  // Принудительное обновление всех данных после синхронизации
  refreshAllDataManual();
  
  SpreadsheetApp.getUi().alert("✅ Команды синхронизированы. Пустые слоты очищены.");
}

/**
 * ФУНКЦИИ УПРАВЛЕНИЯ ВЕБ-ПРИЛОЖЕНИЕМ
 */

/**
 * Полное ручное обновление всех связанных данных (Сводная + Дашборд)
 */
function refreshAllDataManual() {
  if (typeof refreshAllData === 'function') {
    refreshAllData();
    SpreadsheetApp.getActiveSpreadsheet().toast("Данные веб-приложения обновлены", "🔄 Обновление", 3);
  } else {
    // Если функция из других файлов не доступна, вызываем то, что есть
    if (typeof createResultsDashboard === 'function') createResultsDashboard();
  }
}

/**
 * Включение автоматического обновления (создание триггеров)
 */
function enableAutoUpdate() {
  disableAutoUpdate(true); // Сначала удаляем старые, чтобы не дублировать
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Триггер на редактирование (для баллов)
  ScriptApp.newTrigger('onEditAutoHandler')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
    
  // Триггер на изменение структуры (для листов)
  ScriptApp.newTrigger('onChangeAutoHandler')
    .forSpreadsheet(ss)
    .onChange()
    .create();
    
  SpreadsheetApp.getUi().alert("🚀 Автообновление ВКЛЮЧЕНО. Результаты будут обновляться мгновенно.");
}

/**
 * Выключение автоматического обновления (удаление триггеров)
 */
function disableAutoUpdate(silent) {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  
  triggers.forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'onEditAutoHandler' || fn === 'onChangeAutoHandler' || fn === 'autoUpdateDashboard') {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  });
  
  if (!silent) {
    SpreadsheetApp.getUi().alert("📴 Автообновление ВЫКЛЮЧЕНО. Удалено триггеров: " + count);
  }
}

/**
 * Прослойки для обработки событий триггерами
 * Вызывают функции из auto_sync.gs или main.gs
 */
function onEditAutoHandler(e) {
  if (typeof onEdit === 'function') onEdit(e);
}

function onChangeAutoHandler(e) {
  if (typeof autoUpdateDashboard === 'function') autoUpdateDashboard(e);
}