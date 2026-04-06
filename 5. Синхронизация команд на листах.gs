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

    // 2. Сброс чекбоксов и очистка ячеек над ИТОГО для удаленных команд
    const teamNames = headerRange.getValues()[0];
    const lastRow = sheet.getLastRow();
    
    // Пытаемся найти строку "ИТОГО" в колонке A или B
    const searchLimit = Math.min(lastRow, 100);
    const firstColsValues = sheet.getRange(1, 1, searchLimit, 2).getValues();
    let totalRowIndex = -1;
    
    for (let r = 0; r < firstColsValues.length; r++) {
      if (firstColsValues[r][0].toString().includes("ИТОГО") || 
          firstColsValues[r][1].toString().includes("ИТОГО")) {
        totalRowIndex = r + 1;
        break;
      }
    }

    if (lastRow > 1) {
      for (let j = 0; j < maxTeams; j++) {
        // Если имя команды пустое (команда удалена из списка)
        if (!teamNames[j] || teamNames[j].toString().trim() === "") {
          const colIdx = startCol + j;
          
          // А) Определяем границы области чекбоксов
          // Предполагаем, что чекбоксы идут со 2-й строки до (totalRowIndex - 2)
          // А ячейка непосредственно НАД ИТОГО — это (totalRowIndex - 1)
          const endCheckboxRow = totalRowIndex > 0 ? totalRowIndex - 2 : lastRow - 2;
          
          if (endCheckboxRow >= 2) {
            const checkboxRange = sheet.getRange(2, colIdx, endCheckboxRow - 1, 1);
            const values = checkboxRange.getValues();
            let needsClear = false;

            const clearedValues = values.map(row => {
              if (row[0] === true) {
                needsClear = true;
                return [false]; 
              }
              return row;
            });

            if (needsClear) {
              checkboxRange.setValues(clearedValues);
            }
          }

          // Б) Очищаем ячейку НАД СТРОКОЙ ИТОГО (там обычно "Доп. баллы" или ручной ввод)
          if (totalRowIndex > 1) {
            const cellAboveTotal = sheet.getRange(totalRowIndex - 1, colIdx);
            if (cellAboveTotal.getValue() !== "") {
              cellAboveTotal.clearContent();
            }
          }
          
          // В) Саму ячейку ИТОГО (с формулой) НЕ ТРОГАЕМ, как вы и просили.
          // Если там формула, она сама покажет 0 после очистки чекбоксов и доп. баллов.
        }
      }
    }
  });

  // Обновляем дашборд, если функция существует
  if (typeof createResultsDashboard === 'function') {
    createResultsDashboard();
  }

  SpreadsheetApp.getUi().alert("✅ Синхронизация завершена. Чекбоксы и дополнительные баллы очищены.");
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