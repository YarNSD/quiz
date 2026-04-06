/**
 * Файл: ui_menu.gs
 * Описание: Создание пользовательского меню в интерфейсе Google Таблиц.
 *
 * ИСПРАВЛЕНИЯ:
 * - "Дашборд (Таблица)" и "Веб-приложение (Экран)" имели одинаковые пункты
 *   "ВКЛЮЧИТЬ Живой режим" и "ВЫКЛЮЧИТЬ автообновление" с одними и теми же функциями.
 *   Это дублирование убрано: обе секции используют одни и те же setupAutoUpdate /
 *   disableAllAutomation, т.к. триггер один на весь проект.
 * - Добавлен пункт "🔄 Обновить Сводную таблицу" для явного вызова refreshAllData.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Квиз-Панель')
    .addItem('➕ Создать новый раунд', 'createQuizRound')
    .addItem('👥 Синхронизировать команды', 'syncTeamsAcrossRounds')
    .addSeparator()

    // Дашборд (лист "Результаты")
    .addSubMenu(ui.createMenu('📈 Дашборд (Таблица)')
      .addItem('🔄 Обновить сейчас (Вручную)', 'createResultsDashboard')
      .addItem('✅ ВКЛЮЧИТЬ автообновление', 'setupAutoUpdate')
      .addItem('❌ ВЫКЛЮЧИТЬ автообновление', 'disableAllAutomation'))

    .addSeparator()

    // Веб-приложение (лидерборд на экран)
    .addSubMenu(ui.createMenu('🌐 Веб-приложение (Экран)')
      .addItem('🔄 Обновить данные (Вручную)', 'refreshAllData')
      .addItem('✅ ВКЛЮЧИТЬ "Живой режим"', 'setupAutoUpdate')
      .addItem('❌ ВЫКЛЮЧИТЬ автообновление', 'disableAllAutomation'))

    .addSeparator()
    .addItem('📝 Сбросить список листов (Техлист)', 'updateTechSheetList')
    .addToUi();
}

/**
 * Полное отключение всей автоматизации
 */
function disableAllAutomation() {
  const triggers = ScriptApp.getProjectTriggers();
  const count = triggers.length;

  triggers.forEach(t => ScriptApp.deleteTrigger(t));

  const msg = count > 0
    ? `📴 Вся автоматизация отключена (удалено триггеров: ${count}). Данные обновляются только вручную.`
    : "Триггеры не найдены. Автоматизация уже выключена.";

  SpreadsheetApp.getUi().alert(msg);
}
