/**
 * Скрипт для автоматического создания и стилизации листов раундов квиза.
 * Резервирует место под 20 команд (столбцы E-X).
 * Исправлено: пустые столбцы команд больше не скрываются.
 */

function createQuizRound() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setupSheet = ss.getSheetByName("Создание раунда") || ss.getSheetByName("Создание раундов");
  const teamListSheet = ss.getSheetByName("Список команд");
  
  if (!setupSheet || !teamListSheet) {
    SpreadsheetApp.getUi().alert("Ошибка: Проверьте наличие листов 'Создание раунда' и 'Список команд'.");
    return;
  }

  const roundNum = setupSheet.getRange("B2").getValue();
  const questCount = setupSheet.getRange("B3").getValue();
  const pointsPerQuest = setupSheet.getRange("B4").getValue() || 1;
  
  if (!roundNum || !questCount) {
    SpreadsheetApp.getUi().alert("Заполните номер раунда и количество вопросов!");
    return;
  }

  const sheetName = "Раунд " + roundNum;
  if (ss.getSheetByName(sheetName)) {
    SpreadsheetApp.getUi().alert("Лист '" + sheetName + "' уже существует!");
    return;
  }
  
  const sheet = ss.insertSheet(sheetName);
  const maxTeams = 20; 
  const teams = teamListSheet.getRange("B2:B21").getValues();
  
  sheet.getRange("A:Z").setFontFamily("Rubik");
  sheet.setHiddenGridlines(true);
  
  // --- 1. ОФОРМЛЕНИЕ ТАБЛИЦЫ ВОПРОСОВ ---
  const headers = [["№", "Баллы", "Отв.", "Верный ответ"]];
  sheet.getRange(1, 1, 1, 4).setValues(headers)
       .setBackground("#444444").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("B1").setTextRotation(90);

  let questionData = [];
  for (let i = 1; i <= questCount; i++) {
    questionData.push([i, pointsPerQuest, "", ""]);
  }
  sheet.getRange(2, 1, questCount, 4).setValues(questionData).setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Зебра для строк
  for (let i = 0; i < questCount; i++) {
    const rowColor = (i % 2 === 0) ? "#ffffff" : "#f9f9f9";
    sheet.getRange(2 + i, 1, 1, 4 + maxTeams).setBackground(rowColor);
  }

  // --- 2. КОМАНДЫ (СТОЛБЦЫ E-X) ---
  const teamHeaderRange = sheet.getRange(1, 5, 1, maxTeams);
  teamHeaderRange.setBackground("#e8f0fe").setFontColor("#1a73e8").setFontWeight("bold")
                 .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  
  for (let i = 0; i < maxTeams; i++) {
    const teamCell = teamHeaderRange.getCell(1, i + 1);
    // Привязываем названия команд формулой к списку
    teamCell.setFormula("='Список команд'!B" + (i + 2));
    // Блок скрытия столбцов удален по запросу
  }
  
  sheet.getRange(2, 5, questCount, maxTeams).insertCheckboxes();
  sheet.getRange(2, 3, questCount, 1).setBackground("#fff2cc").setFontWeight("bold");
  sheet.getRange(2, 4, questCount, 1).setBackground("#d9ead3").setFontWeight("bold").setHorizontalAlignment("left");
  
  // --- 3. ИТОГИ ---
  const bonusRow = questCount + 2;
  const totalRow = questCount + 3;

  sheet.getRange(bonusRow, 1, 1, 4).merge().setValue("Доп. баллы:").setFontWeight("bold").setHorizontalAlignment("right").setBackground("#fff2cc");
  sheet.getRange(totalRow, 1, 1, 4).merge().setValue("ИТОГО:").setFontWeight("bold").setHorizontalAlignment("right").setBackground("#f1f3f4");

  for (let j = 0; j < maxTeams; j++) {
    const colLetter = columnToLetter(5 + j);
    sheet.getRange(bonusRow, 5 + j).setBackground("#fffef3").setHorizontalAlignment("center");
    const formula = `=SUMPRODUCT($B$2:$B$${questCount + 1}; ${colLetter}2:${colLetter}${questCount + 1}) + N(${colLetter}${bonusRow})`;
    sheet.getRange(totalRow, 5 + j).setFormula(formula)
         .setFontWeight("bold").setBackground("#34a853").setFontColor("#ffffff").setHorizontalAlignment("center").setFontSize(12);
  }

  // --- 4. ТЕКУЩИЙ РЕЙТИНГ ---
  const dashStartRow = totalRow + 2; 
  
  sheet.getRange(dashStartRow, 3, 1, 2).merge()
       .setValue("🏆 РЕЙТИНГ РАУНДА")
       .setBackground("#444444").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange(dashStartRow + 1, 3).setValue("Баллы").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#f1f3f4");
  sheet.getRange(dashStartRow + 1, 4).setValue("Команда").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#f1f3f4");

  const teamsRange = `${columnToLetter(5)}1:${columnToLetter(5 + maxTeams - 1)}1`;
  const totalsRange = `${columnToLetter(5)}${totalRow}:${columnToLetter(5 + maxTeams - 1)}${totalRow}`;
  
  const sortFormula = `=SORT(FILTER({TRANSPOSE(${totalsRange})\\ TRANSPOSE(${teamsRange})}; TRANSPOSE(${teamsRange})<>\"\"); 1; FALSE)`;
  
  sheet.getRange(dashStartRow + 2, 3).setFormula(sortFormula);
  
  const resultsRange = sheet.getRange(dashStartRow + 2, 3, maxTeams, 2);
  resultsRange.setFontSize(12).setVerticalAlignment("middle");
  sheet.getRange(dashStartRow + 2, 3, maxTeams, 1).setHorizontalAlignment("center").setFontWeight("bold").setFontColor("#34a853");
  sheet.getRange(dashStartRow + 2, 4, maxTeams, 1).setHorizontalAlignment("left");

  // Настройки размеров
  sheet.setColumnWidth(1, 25); 
  sheet.setColumnWidth(2, 25); 
  sheet.setColumnWidth(3, 60); 
  sheet.setColumnWidth(4, 240); 
  sheet.setColumnWidths(5, maxTeams, 110);
  
  sheet.setRowHeight(1, 60);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4);

  SpreadsheetApp.getUi().alert("Раунд создан!");
}

/**
 * Вспомогательная функция для конвертации индекса колонки в букву
 */
function columnToLetter(column) {
  let temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}