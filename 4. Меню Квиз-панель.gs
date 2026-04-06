
  Файл ui_menu.gs
  Описание Создание пользовательского меню в интерфейсе Google Таблиц.
 

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Квиз-Панель')
    .addItem('➕ Создать новый раунд', 'createQuizRound')
    .addItem('👥 Синхронизировать команды', 'syncTeamsAcrossRounds')
    .addSeparator()
    
     Раздел управления Дашбордом (Лист Результаты)
    .addSubMenu(ui.createMenu('📈 Дашборд (Таблица)')
      .addItem('🔄 Обновить сейчас (Вручную)', 'createResultsDashboard')
      .addItem('✅ ВКЛЮЧИТЬ автообновление', 'setupAutoUpdate')
      .addItem('❌ ВЫКЛЮЧИТЬ автообновление', 'disableAllAutomation'))
    
    .addSeparator()
    
     Раздел управления Веб-приложением (Внешняя ссылка)
    .addSubMenu(ui.createMenu('🌐 Веб-приложение (Экран)')
      .addItem('🔄 Обновить данные (Вручную)', 'refreshAllData')
      .addItem('✅ ВКЛЮЧИТЬ Живой режим', 'setupAutoUpdate')
      .addItem('❌ ВЫКЛЮЧИТЬ автообновление', 'disableAllAutomation'))
    
    .addSeparator()
    .addItem('📝 Сбросить список листов (Техлист)', 'updateTechSheetList')
    .addToUi();
}


  Функция для полного отключения всей автоматизации и триггеров
 
function disableAllAutomation() {
  const triggers = ScriptApp.getProjectTriggers();
  const count = triggers.length;
  
  triggers.forEach(t = ScriptApp.deleteTrigger(t));
  
  const msg = count  0 
     `📴 Вся автоматизация отключена (удалено триггеров ${count}). Данные обновляются только вручную.`
     Триггеры не найдены. Автоматизация уже выключена.;
    
  SpreadsheetApp.getUi().alert(msg);
}