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
  
  sheet.getRange("A:Z").setFontFamily("Rubik");
  sheet.setHiddenGridlines(true);
  
  // Шапка вопросов
  const headers = [["№", "Баллы", "Отв.", "Верный ответ"]];
  sheet.getRange(1, 1, 1, 4).setValues(headers)
       .setBackground("#444444").setFontColor("#ffffff").setFontWeight("bold")
       .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("B1").setTextRotation(90);

  // Данные вопросов
  let questionData = [];
  for (let i = 1; i <= questCount; i++) {
    questionData.push([i, pointsPerQuest, "", ""]);
  }
  sheet.getRange(2, 1, questCount, 4).setValues(questionData).setHorizontalAlignment("center");

  // Шапка команд (Е2:Х2)
  const teamHeaderRange = sheet.getRange(1, 5, 1, maxTeams);
  teamHeaderRange.setBackground("#e8f0fe").setFontColor("#1a73e8").setFontWeight("bold")
                 .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  
  for (let i = 0; i < maxTeams; i++) {
    const teamCell = teamHeaderRange.getCell(1, i + 1);
    teamCell.setFormula("='Список команд'!B" + (i + 2));
  }
  
  // Скрываем пустые колонки команд сразу при создании
  const teamsData = teamListSheet.getRange("B2:B21").getValues();
  for (let i = 0; i < maxTeams; i++) {
    if (!teamsData[i] || teamsData[i][0] === "") {
      sheet.hideColumns(5 + i);
    }
  }

  // Чекбоксы
  sheet.getRange(2, 5, questCount, maxTeams).insertCheckboxes();
  
  // Итоги
  const bonusRow = questCount + 2;
  const totalRow = questCount + 3;
  const dashStartRow = questCount + 5;

  sheet.getRange(bonusRow, 1, 1, 4).merge().setValue("Доп. баллы:").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange(totalRow, 1, 1, 4).merge().setValue("ИТОГО:").setFontWeight("bold").setHorizontalAlignment("right");

  for (let j = 0; j < maxTeams; j++) {
    const col = columnToLetter(5 + j);
    sheet.getRange(totalRow, 5 + j).setFormula(`=SUMPRODUCT($B$2:$B$${questCount + 1}; ${col}2:${col}${questCount + 1}) + N(${col}${bonusRow})`)
         .setFontWeight("bold").setBackground("#34a853").setFontColor("#ffffff").setHorizontalAlignment("center");
  }

  // Мини-дашборд на листе раунда
  sheet.getRange(dashStartRow, 1, 1, 4).merge().setValue("РЕЙТИНГ РАУНДА").setBackground("#444444").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(dashStartRow + 1, 3).setValue("Баллы").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(dashStartRow + 1, 4).setValue("Команда").setFontWeight("bold");

  const teamsRange = `$E$1:$${columnToLetter(4 + maxTeams)}$1`;
  const totalsRange = `$E$${totalRow}:$${columnToLetter(4 + maxTeams)}$${totalRow}`;
  const sortFormula = `=SORT(FILTER({TRANSPOSE(${totalsRange})\\ TRANSPOSE(${teamsRange})}; TRANSPOSE(${teamsRange})<>""); 1; FALSE)`;
  
  sheet.getRange(dashStartRow + 2, 3).setFormula(sortFormula);
  
  const resultsRange = sheet.getRange(dashStartRow + 2, 3, maxTeams, 2);
  resultsRange.setFontSize(12).setVerticalAlignment("middle");
  sheet.getRange(dashStartRow + 2, 3, maxTeams, 1).setHorizontalAlignment("center").setFontWeight("bold").setFontColor("#34a853");
  sheet.getRange(dashStartRow + 2, 4, maxTeams, 1).setHorizontalAlignment("left");

  // Размеры
  sheet.setColumnWidth(1, 25); 
  sheet.setColumnWidth(2, 25); 
  sheet.setColumnWidth(3, 60);
  sheet.setColumnWidth(4, 240); 
  sheet.setColumnWidths(5, maxTeams, 110);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4);
}