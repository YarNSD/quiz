/**
 * Скрипт для автоматического создания и стилизации листов раундов квиза.
 * Резервирует место под 20 команд (столбцы E-X).
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
  
  // УСТАНОВКА ССЫЛОК НА КОМАНДЫ (Функционал №1)
  // Записываем формулы ссылок на лист 'Список команд' (B2:B21)
  const teamLinks = [];
  for (let i = 2; i <= 21; i++) {
    teamLinks.push([`='Список команд'!B${i}`]);
  }
  // Транспонируем и вставляем в первую строку начиная с 5-го столбца (E1)
  sheet.getRange(1, 5, 1, maxTeams).setFormulas([teamLinks.map(r => r[0])]);
  
  sheet.getRange("A:Z").setFontFamily("Ubuntu").setVerticalAlignment("middle");

  // Заголовки вопросов
  const headerData = [["№", "Вопрос", "Баллы", "Верно"]];
  sheet.getRange(1, 1, 1, 4).setValues(headerData).setBackground("#444444").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  
  // Стили для названий команд в шапке
  const teamHeaderRange = sheet.getRange(1, 5, 1, maxTeams);
  teamHeaderRange.setBackground("#fbbc04").setFontColor("#000000").setFontWeight("bold").setHorizontalAlignment("center").setFontSize(11);

  // Создание строк вопросов
  const questData = [];
  for (let i = 1; i <= questCount; i++) {
    questData.push([i, "Вопрос " + i, pointsPerQuest, 0]);
  }
  sheet.getRange(2, 1, questCount, 4).setValues(questData).setHorizontalAlignment("center");
  sheet.getRange(2, 2, questCount, 1).setHorizontalAlignment("left");

  // Сетка чекбоксов
  const checkboxRange = sheet.getRange(2, 5, questCount, maxTeams);
  checkboxRange.insertCheckboxes().setHorizontalAlignment("center");

  // Формулы подсчета (Верно в строке)
  for (let i = 0; i < questCount; i++) {
    const row = i + 2;
    const range = `${columnToLetter(5)}${row}:${columnToLetter(5 + maxTeams - 1)}${row}`;
    sheet.getRange(row, 4).setFormula(`=COUNTIF(${range}; TRUE)`);
  }

  // Строка ИТОГО
  const totalRow = questCount + 2;
  sheet.getRange(totalRow, 1, 1, 4).setBackground("#eeeeee").setFontWeight("bold");
  sheet.getRange(totalRow, 2).setValue("ИТОГО БАЛЛОВ:");
  
  // Формулы суммы по столбцам команд
  for (let j = 0; j < maxTeams; j++) {
    const colIdx = 5 + j;
    const colLetter = columnToLetter(colIdx);
    const sumFormula = `=SUMPRODUCT(${colLetter}2:${colLetter}${questCount + 1}; $C$2:$C${questCount + 1})`;
    sheet.getRange(totalRow, colIdx).setFormula(sumFormula).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#fff2cc");
  }

  // Область результатов (Дашборд внутри листа)
  const dashStartRow = totalRow + 2;
  sheet.getRange(dashStartRow, 3, 1, 2).setValues([["Очки", "Команда"]]).setBackground("#34a853").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

  const teamsRange = `${columnToLetter(5)}1:${columnToLetter(5 + maxTeams - 1)}1`;
  const totalsRange = `${columnToLetter(5)}${totalRow}:${columnToLetter(5 + maxTeams - 1)}${totalRow}`;
  
  const sortFormula = `=SORT(FILTER({TRANSPOSE(${totalsRange})\\ TRANSPOSE(${teamsRange})}; TRANSPOSE(${teamsRange})<>""); 1; FALSE)`;
  sheet.getRange(dashStartRow + 1, 3).setFormula(sortFormula);
  
  // Форматирование
  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 60);
  sheet.setColumnWidths(5, maxTeams, 100);

  SpreadsheetApp.getUi().alert("✅ Раунд создан! Названия команд теперь связаны со списком.");
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