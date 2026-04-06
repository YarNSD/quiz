/**
 * Файл: sync_logic.gs
 * Описание: Синхронизация команд и очистка данных для удаленных команд.
 *
 * ИСПРАВЛЕНИЯ:
 * - Поиск строки ИТОГО искал по двум столбцам ($A и $B), тогда как "ИТОГО:"
 *   записывается ТОЛЬКО в объединённую ячейку столбца A. Убрана проверка столбца B
 *   — это исключает ложные срабатывания, если в столбце B есть ячейки со словом ИТОГО.
 * - Math.min(lastRow, 100): ограничение 100 строк при поиске ИТОГО может не хватить,
 *   если в раунде > ~97 вопросов. Заменено на lastRow без ограничения (безопасно,
 *   т.к. это чтение из памяти, а не запрос к API).
 * - onEditAutoHandler объявлен, но нигде не регистрируется как триггер.
 *   Добавлен комментарий. Функция оставлена для совместимости.
 * - Очистка "доп. баллов" (rangesToClear) проверяла totalRowIndex - 2 в массиве,
 *   что соответствует строке "Доп. баллы" (bonusRow). Логика корректна, добавлен комментарий.
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
    if (lastRow < 2) {
      console.log("Лист " + sheetName + " пустой, пропуск.");
      console.timeEnd("Обработка листа " + sheetName);
      return;
    }

    // 1. Обновление формул заголовков (пакетно)
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

    // 2. Поиск строки ИТОГО
    const teamNames = headerRange.getValues()[0];
    // ИСПРАВЛЕНИЕ: читаем только столбец A (не A:B) и без ограничения 100 строк
    const firstColValues = sheet.getRange(1, 1, lastRow, 1).getValues();

    let totalRowIndex = -1;
    for (let r = 0; r < firstColValues.length; r++) {
      // ИСПРАВЛЕНИЕ: проверяем только столбец A, т.к. "ИТОГО:" пишется туда
      if (firstColValues[r][0].toString().includes("ИТОГО:")) {
        totalRowIndex = r + 1; // 1-based номер строки
        break;
      }
    }

    if (totalRowIndex === -1) {
      console.warn("Строка ИТОГО не найдена на листе " + sheetName);
      console.timeEnd("Обработка листа " + sheetName);
      return;
    }

    // 3. Очистка данных удалённых команд
    console.time("Сбор данных и очистка на " + sheetName);

    const dataRange = sheet.getRange(1, startCol, totalRowIndex, maxTeams);
    const dataValues = dataRange.getValues();

    let rangesToUncheck = [];
    let rangesToClear = [];

    for (let j = 0; j < maxTeams; j++) {
      if (!teamNames[j] || teamNames[j].toString().trim() === "") {
        const colIdx = startCol + j;

        // Строка "Доп. баллы" — это totalRowIndex - 1 (индекс в массиве: totalRowIndex - 2)
        if (dataValues[totalRowIndex - 2] && dataValues[totalRowIndex - 2][j] !== "") {
          rangesToClear.push(sheet.getRange(totalRowIndex - 1, colIdx).getA1Notation());
        }

        // Область чекбоксов: строки 2..(totalRowIndex - 2)
        let hasChecked = false;
        for (let r = 1; r < totalRowIndex - 2; r++) {
          if (dataValues[r] && dataValues[r][j] === true) {
            hasChecked = true;
            break;
          }
        }

        if (hasChecked) {
          rangesToUncheck.push(sheet.getRange(2, colIdx, totalRowIndex - 3, 1).getA1Notation());
        }
      }
    }

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
 * ПРИМЕЧАНИЕ: эта функция не регистрируется как триггер автоматически.
 * Если нужно — зарегистрируйте её через setupAutoUpdate() или Apps Script UI.
 * Логика дублирует onEdit() из файла 3, оставлена для совместимости.
 */
function onEditAutoHandler(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  const name = sheet.getName();

  if (name.startsWith("Раунд ") || name === "Список команд") {
    if (typeof createResultsDashboard === 'function') createResultsDashboard();
  }
}
