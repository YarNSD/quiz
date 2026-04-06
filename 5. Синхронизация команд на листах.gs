/**
 * Файл: sync_logic.gs
 * Описание: Синхронизация команд и очистка данных для удаленных команд.
 */

/**
 * Основная функция синхронизации команд.
 * Обновляет имена (если вдруг формулы слетели) и СБРАСЫВАЕТ ЧЕКБОКСЫ в пустых столбцах.
 */
function syncTeamsAcrossRounds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const rounds = allSheets.filter(s => s.getName().startsWith("Раунд "));
  const teamListSheet = ss.getSheetByName("Список команд");
  
  if (!teamListSheet) {
    SpreadsheetApp.getUi().alert("Ошибка: Лист 'Список команд' не найден.");
    return;
  }

  const maxTeams = 20;
  const startCol = 5; // Столбец E

  rounds.forEach(sheet => {
    // 1. Проверяем/обновляем формулы названий в первой строке (на всякий случай)
    const headerRange = sheet.getRange(1, startCol, 1, maxTeams);
    const currentFormulas = headerRange.getFormulas()[0];
    
    // Если формул нет или они не ссылаются на список команд — восстанавливаем
    for (let i = 0; i < maxTeams; i++) {
      const expectedFormula = `='Список команд'!B${i + 2}`;
      if (currentFormulas[i] !== expectedFormula) {
        sheet.getRange(1, startCol + i).setFormula(expectedFormula);
      }
    }

    // 2. Сброс чекбоксов для удаленных команд (Функционал №2)
    // Читаем значения заголовков (результаты формул)
    const teamNames = headerRange.getValues()[0];
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      for (let j = 0; j < maxTeams; j++) {
        // Если имя команды пустое (команда удалена из списка)
        if (!teamNames[j] || teamNames[j].toString().trim() === "") {
          const colIdx = startCol + j;
          // Находим диапазон под заголовком до последней строки (где чекбоксы)
          const dataRange = sheet.getRange(2, colIdx, lastRow - 1, 1);
          const values = dataRange.getValues();
          let needsClear = false;

          // Проверяем, есть ли там включенные чекбоксы
          const clearedValues = values.map(row => {
            if (row[0] === true) {
              needsClear = true;
              return [false]; // Сбрасываем в false
            }
            return row;
          });

          if (needsClear) {
            dataRange.setValues(clearedValues);
          }
        }
      }
    }
  });

  // Обновляем дашборд, если функция существует
  if (typeof createResultsDashboard === 'function') {
    createResultsDashboard();
  }

  SpreadsheetApp.getUi().alert("✅ Синхронизация завершена. Чекбоксы в пустых столбцах сброшены.");
}

/**
 * Обработчик автоматического обновления
 */
function onEditAutoHandler(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  const name = sheet.getName();
  
  if (name.startsWith("Раунд ") || name === "Список команд") {
    if (typeof createResultsDashboard === 'function') createResultsDashboard();
  }
}