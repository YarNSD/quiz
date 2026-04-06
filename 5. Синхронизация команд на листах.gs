/**
 * Файл: sync_logic.gs
 * Описание: Синхронизация команд и очистка данных для удаленных команд.
 */

/**
 * Основная функция синхронизации команд.
 * Оптимизирована для сверхбыстрого выполнения (Batch operations).
 * Добавлено логирование для анализа производительности.
 */
function syncTeamsAcrossRounds() {
  console.time("Общее время выполнения");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.time("Поиск листов");
  const allSheets = ss.getSheets();
  const rounds = allSheets.filter(s => s.getName().startsWith("Раунд "));
  const teamListSheet = ss.getSheetByName("Список команд");
  console.timeEnd("Поиск листов");
  
  if (!teamListSheet) {
    SpreadsheetApp.getUi().alert("Ошибка: Лист 'Список команд' не найден.");
    return;
  }

  const maxTeams = 20;
  const startCol = 5; // Столбец E

  console.log("Найдено листов раундов: " + rounds.length);

  rounds.forEach((sheet) => {
    const sheetName = sheet.getName();
    console.time("Обработка листа " + sheetName);
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) {
      console.log("Лист " + sheetName + " пустой, пропуск.");
      console.timeEnd("Обработка листа " + sheetName);
      return;
    }

    // 1. Проверка и обновление формул заголовков (Пакетно)
    const headerRange = sheet.getRange(1, startCol, 1, maxTeams);
    const currentFormulas = headerRange.getFormulas()[0];
    const newFormulas = [];
    let needsFormulaUpdate = false;

    for (let i = 0; i < maxTeams; i++) {
      const expectedFormula = `='Список команд'!B${i + 2}`;
      newFormulas.push(expectedFormula);
      if (currentFormulas[i] !== expectedFormula) needsFormulaUpdate = true;
    }

    if (needsFormulaUpdate) {
      headerRange.setFormulas([newFormulas]);
    }

    // 2. Поиск строки ИТОГО и получение данных листа в память
    // Считываем область A:B и всю строку заголовков один раз
    const teamNames = headerRange.getValues()[0];
    const firstColsValues = sheet.getRange(1, 1, Math.min(lastRow, 100), 2).getValues();
    
    let totalRowIndex = -1;
    for (let r = 0; r < firstColsValues.length; r++) {
      if (firstColsValues[r][0].toString().includes("ИТОГО") || 
          firstColsValues[r][1].toString().includes("ИТОГО")) {
        totalRowIndex = r + 1;
        break;
      }
    }

    if (totalRowIndex === -1) {
      console.warn("Строка ИТОГО не найдена на листе " + sheetName);
      console.timeEnd("Обработка листа " + sheetName);
      return;
    }

    // 3. Оптимизированная очистка (Сбор диапазонов)
    console.time("Сбор данных и очистка на " + sheetName);
    
    // Считываем значения всего рабочего диапазона команд один раз, чтобы проверить, нужно ли что-то чистить
    const dataRange = sheet.getRange(1, startCol, totalRowIndex, maxTeams);
    const dataValues = dataRange.getValues();
    
    let rangesToUncheck = [];
    let rangesToClear = [];

    for (let j = 0; j < maxTeams; j++) {
      // Если имя команды пустое
      if (!teamNames[j] || teamNames[j].toString().trim() === "") {
        const colIdx = startCol + j;
        
        // Проверяем ячейку над ИТОГО (строка totalRowIndex - 1)
        // Индекс в массиве dataValues будет [totalRowIndex - 2][j] (так как массив с 0, а заголовок в 1 строке)
        if (dataValues[totalRowIndex - 2][j] !== "") {
          rangesToClear.push(sheet.getRange(totalRowIndex - 1, colIdx).getA1Notation());
        }

        // Проверяем область чекбоксов (строки со 2 по totalRowIndex - 2)
        // Если хотя бы один чекбокс нажат, добавляем весь диапазон в список на очистку
        let hasChecked = false;
        for (let r = 1; r < totalRowIndex - 2; r++) { // от 2-й строки до строки над "доп баллами"
          if (dataValues[r][j] === true) {
            hasChecked = true;
            break;
          }
        }
        
        if (hasChecked) {
          rangesToUncheck.push(sheet.getRange(2, colIdx, totalRowIndex - 3, 1).getA1Notation());
        }
      }
    }

    // Выполняем очистку пакетно через RangeList (самый быстрый способ в GAS)
    if (rangesToClear.length > 0) {
      sheet.getRangeList(rangesToClear).clearContent();
    }
    if (rangesToUncheck.length > 0) {
      sheet.getRangeList(rangesToUncheck).uncheck();
    }
    
    console.timeEnd("Сбор данных и очистка на " + sheetName);
    console.timeEnd("Обработка листа " + sheetName);
  });

  // 4. Обновление дашборда
  if (typeof createResultsDashboard === 'function') {
    console.time("Обновление Дашборда");
    createResultsDashboard();
    console.timeEnd("Обновление Дашборда");
  }

  console.timeEnd("Общее время выполнения");
  SpreadsheetApp.getUi().alert("✅ Синхронизация завершена.");
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