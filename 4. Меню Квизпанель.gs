/**
 * Файл: menu_manager.gs
 * Создает пользовательское меню "🏆 КВИЗ-ПАНЕЛЬ" для управления таблицей.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏆 КВИЗ-ПАНЕЛЬ')
   
    .addItem('🔄 Обновить Дашборд (Ручной запуск)', 'createResultsDashboard')
    .addSeparator()
    // Блок 1: Основные действия с таблицей
    .addSubMenu(ui.createMenu('⚙️ Подготовка нового квиза')
    .addItem('➕ Создать новый раунд', 'createQuizRound') 
    .addItem('👥 Сброс ответов на раундах там. где не указана команда', 'syncTeamsAcrossRounds')
    .addItem('📊 Пересчитать Сводную таблицу', 'refreshAllData'))
    .addSeparator()

    // Блок 2: Веб-интерфейс
    .addSubMenu(ui.createMenu('🌐 Веб-интерфейс')
      .addItem('✅ Включить автообновление LIVE', 'enableWebRefresh')
      .addItem('❌ Выключить автообновление LIVE', 'disableWebRefresh'))
    .addSeparator()

    // Блок 3: Настройки и автоматизация
    .addSubMenu(ui.createMenu('⚙️ Настройки')
      .addItem('✅ Включить автообновление (Техлист/Сводная)', 'setupAutoUpdate')
      .addItem('❌ Отключить автообновление', 'disableAutoUpdate')
      .addItem('📜 Обновить список листов в Техлисте', 'updateTechSheetList'))
    
    .addToUi();
}

/**
 * Включает автообновление для веб-морды (устанавливает флаг в свойствах документа)
 */
function enableWebRefresh() {
  PropertiesService.getDocumentProperties().setProperty('WEB_REFRESH_ENABLED', 'true');
  SpreadsheetApp.getUi().alert("✅ Автообновление веб-интерфейса включено. Теперь у пользователей данные будут обновляться автоматически.");
}

/**
 * Выключает автообновление для веб-морды
 */
function disableWebRefresh() {
  PropertiesService.getDocumentProperties().setProperty('WEB_REFRESH_ENABLED', 'false');
  SpreadsheetApp.getUi().alert("❌ Автообновление веб-интерфейса выключено.");
}
