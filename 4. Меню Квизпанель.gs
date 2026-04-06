/**
 * Файл: 4__Меню_Квизпанель.gs
 * Обновлено меню для управления ручным и автоматическим режимом.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Квиз-Панель')
    .addItem('➕ Создать новый раунд', 'createQuizRound')
    .addItem('👥 Синхронизировать команды', 'syncTeamsAcrossRounds')
    .addSeparator()

    // Секция Дашборда
    .addSubMenu(ui.createMenu('📈 Лист РЕЗУЛЬТАТЫ')
      .addItem('🔄 ОБНОВИТЬ СЕЙЧАС (Вручную)', 'createResultsDashboard'))


    .addSeparator()
    .addItem('📝 Обновить Техлист (Список листов)', 'updateTechSheetList')
    .addToUi();
}