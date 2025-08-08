/**
 * Restaurant Analytics System - Main Module
 * Главный модуль, который объединяет все функции системы
 * Автор: Restaurant Analytics  
 * Версия: 3.2 RU
 */

/**
 * Основная точка входа в систему
 * Запускает полный цикл сбора данных, обработки, анализа и генерации отчетов
 */
function main() {
  const startTime = new Date();
  Logger.log('=== ЗАПУСК ПОЛНОЙ АНАЛИТИКИ РЕСТОРАНА ===');
  Logger.log('Время начала: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    // 1. Запуск полного процесса аналитики
    Logger.log('1/3 - Запуск полного процесса аналитики...');
    runFullAnalyticsProcess();
    Logger.log('✓ Полный процесс аналитики завершен');
    
    // 2. UTM анализ
    Logger.log('2/3 - Запуск UTM анализа...');
    runUtmAnalysis();
    Logger.log('✓ UTM анализ завершен');
    
    // 3. Создание итоговых отчетов
    Logger.log('3/3 - Создание итоговых отчетов...');
    createAmoSummaryTable();
    Logger.log('✓ Итоговые отчеты созданы');
    
    const executionTime = (new Date() - startTime) / 1000;
    const minutes = Math.floor(executionTime / 60);
    const seconds = Math.floor(executionTime % 60);
    
    Logger.log('=== ПОЛНАЯ АНАЛИТИКА ЗАВЕРШЕНА ===');
    Logger.log(`Время выполнения: ${minutes}м ${seconds}с`);
    Logger.log('Все данные обновлены и отчеты созданы');
    
    return {
      success: true,
      executionTime: executionTime,
      message: `Аналитика завершена за ${minutes}м ${seconds}с`
    };
    
  } catch (error) {
    Logger.log('КРИТИЧЕСКАЯ ОШИБКА в main(): ' + error.toString());
    Logger.log('Стек ошибки: ' + error.stack);
    
    return {
      success: false,
      error: error.toString(),
      message: 'Произошла ошибка при выполнении аналитики'
    };
  }
}

/**
 * Быстрый запуск только сбора данных
 */
function quickDataCollection() {
  const startTime = new Date();
  Logger.log('=== БЫСТРЫЙ СБОР ДАННЫХ ===');
  Logger.log('Время начала: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    hourlyDataCollection();
    
    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== СБОР ДАННЫХ ЗАВЕРШЕН ===');
    Logger.log(`Время выполнения: ${executionTime} секунд`);
    
    return {
      success: true,
      executionTime: executionTime,
      message: `Сбор данных завершен за ${executionTime} секунд`
    };
    
  } catch (error) {
    Logger.log('ОШИБКА в quickDataCollection(): ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      message: 'Произошла ошибка при сборе данных'
    };
  }
}

/**
 * Быстрый запуск только анализа
 */
function quickAnalysis() {
  const startTime = new Date();
  Logger.log('=== БЫСТРЫЙ АНАЛИЗ ===');
  Logger.log('Время начала: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    analyzeData();
    
    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== АНАЛИЗ ЗАВЕРШЕН ===');
    Logger.log(`Время выполнения: ${executionTime} секунд`);
    
    return {
      success: true,
      executionTime: executionTime,
      message: `Анализ завершен за ${executionTime} секунд`
    };
    
  } catch (error) {
    Logger.log('ОШИБКА в quickAnalysis(): ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      message: 'Произошла ошибка при анализе данных'
    };
  }
}

/**
 * Создает все недостающие листы в таблице
 */
function createMissingSheets() {
  const startTime = new Date();
  Logger.log('=== СОЗДАНИЕ НЕДОСТАЮЩИХ ЛИСТОВ ===');
  Logger.log('Время начала: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const existingSheets = spreadsheet.getSheets().map(s => s.getName());
    const requiredSheets = Object.values(CONFIG.SHEETS);
    
    let createdCount = 0;
    
    requiredSheets.forEach(sheetName => {
      // Пропускаем листы с префиксами (они создаются динамически)
      if (sheetName.includes('АНАЛИТИКА ЕВГЕНИЧЬ СПБ')) return;
      
      if (!existingSheets.includes(sheetName)) {
        Logger.log('Создание листа: ' + sheetName);
        const newSheet = spreadsheet.insertSheet(sheetName);
        
        // Добавляем базовые заголовки для основных листов
        if (sheetName === 'ЕДИНАЯ_БАЗА_КЛИЕНТОВ') {
          newSheet.getRange(1, 1, 1, 15).setValues([[
            'ID (Телефон)', 'Имя', 'Email', 'Первый визит', 'Общая сумма', 
            'Кол-во визитов', 'Средний чек', 'Последний визит', 'Первый источник',
            'UTM Source', 'UTM Medium', 'UTM Campaign', 'UTM Content', 'UTM Term', 'Статус'
          ]]);
          newSheet.getRange(1, 1, 1, 15).setFontWeight('bold').setBackground('#d0d0d0');
        } else if (sheetName === 'ПУТЬ_КЛИЕНТА') {
          newSheet.getRange(1, 1, 1, 8).setValues([[
            'ID Клиента', 'Дата события', 'Тип события', 'Источник', 
            'Сумма', 'Описание', 'UTM данные', 'Timestamp'
          ]]);
          newSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#d0d0d0');
        } else if (sheetName === 'КАЧЕСТВО_ДАННЫХ') {
          newSheet.getRange(1, 1, 1, 6).setValues([[
            'Источник данных', 'Всего записей', 'Валидных записей', 
            'Процент качества', 'Последняя проверка', 'Статус'
          ]]);
          newSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#d0d0d0');
        } else if (sheetName === '_CACHE_ALL_DATA') {
          newSheet.getRange(1, 1, 1, 4).setValues([[
            'Источник', 'Данные', 'Timestamp', 'Размер'
          ]]);
          newSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');
        }
        
        createdCount++;
        Logger.log('✓ Лист создан: ' + sheetName);
      }
    });
    
    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== СОЗДАНИЕ ЛИСТОВ ЗАВЕРШЕНО ===');
    Logger.log(`Создано листов: ${createdCount}`);
    Logger.log(`Время выполнения: ${executionTime} секунд`);
    
    return {
      success: true,
      createdCount: createdCount,
      executionTime: executionTime,
      message: `Создано ${createdCount} новых листов`
    };
    
  } catch (error) {
    Logger.log('ОШИБКА в createMissingSheets: ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      message: 'Не удалось создать листы'
    };
  }
}

/**
 * Инициализация системы (первый запуск)
 */
function initializeSystem() {
  const startTime = new Date();
  Logger.log('=== ИНИЦИАЛИЗАЦИЯ СИСТЕМЫ ===');
  Logger.log('Время начала: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    // Проверяем доступность основной таблицы
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    Logger.log('✓ Основная таблица доступна: ' + spreadsheet.getName());
    
    // Проверяем наличие всех необходимых листов
    const requiredSheets = Object.values(CONFIG.SHEETS);
    const existingSheets = spreadsheet.getSheets().map(s => s.getName());
    
    Logger.log('Проверка листов:');
    requiredSheets.forEach(sheetName => {
      if (existingSheets.includes(sheetName)) {
        Logger.log(`✓ ${sheetName} - найден`);
      } else {
        Logger.log(`✗ ${sheetName} - отсутствует`);
      }
    });
    
    // Создаем недостающие листы
    Logger.log('Создание недостающих листов...');
    const sheetCreation = createMissingSheets();
    if (sheetCreation.success) {
      Logger.log(`✓ Создано листов: ${sheetCreation.createdCount}`);
    } else {
      Logger.log('✗ Ошибка создания листов: ' + sheetCreation.error);
    }
    
    // Настраиваем автоматические триггеры
    Logger.log('Настройка автоматических триггеров...');
    setupAutomaticTriggers();
    Logger.log('✓ Триггеры настроены');
    
    // Создаем первый бэкап
    Logger.log('Создание первоначального бэкапа...');
    const backupInfo = createDataBackup();
    Logger.log('✓ Бэкап создан: ' + backupInfo.backupUrl);
    
    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== ИНИЦИАЛИЗАЦИЯ ЗАВЕРШЕНА ===');
    Logger.log(`Время выполнения: ${executionTime} секунд`);
    
    return {
      success: true,
      executionTime: executionTime,
      backupUrl: backupInfo.backupUrl,
      message: 'Система успешно инициализирована'
    };
    
  } catch (error) {
    Logger.log('ОШИБКА инициализации: ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      message: 'Произошла ошибка при инициализации системы'
    };
  }
}

/**
 * Получает текущий статус системы
 */
function getSystemStatus() {
  const status = {
    timestamp: new Date().toLocaleString('ru-RU'),
    spreadsheetAccess: false,
    lastDataCollection: null,
    lastAnalysis: null,
    lastBackup: null,
    activeTriggers: [],
    errors: []
  };
  
  try {
    // Проверяем доступ к основной таблице
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    status.spreadsheetAccess = true;
    status.spreadsheetName = spreadsheet.getName();
    
    // Проверяем последние операции
    const properties = PropertiesService.getScriptProperties();
    
    const lastDataCollectionStr = properties.getProperty('lastDataCollection');
    if (lastDataCollectionStr) {
      status.lastDataCollection = JSON.parse(lastDataCollectionStr);
    }
    
    const lastAnalysisStr = properties.getProperty('analyticsData');
    if (lastAnalysisStr) {
      const analyticsData = JSON.parse(lastAnalysisStr);
      status.lastAnalysis = analyticsData.lastUpdate;
    }
    
    const lastBackupStr = properties.getProperty('lastBackup');
    if (lastBackupStr) {
      status.lastBackup = JSON.parse(lastBackupStr);
    }
    
    // Проверяем активные триггеры
    status.activeTriggers = getTriggerStatus();
    
  } catch (error) {
    status.errors.push(error.toString());
  }
  
  Logger.log('Статус системы: ' + JSON.stringify(status, null, 2));
  return status;
}

/**
 * Очистка и обслуживание системы
 */
function maintenanceCleanup() {
  const startTime = new Date();
  Logger.log('=== ОБСЛУЖИВАНИЕ СИСТЕМЫ ===');
  Logger.log('Время начала: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    // Создаем бэкап перед очисткой
    Logger.log('Создание бэкапа перед очисткой...');
    const backupInfo = createDataBackup();
    Logger.log('✓ Бэкап создан');
    
    // Очищаем старые данные (старше 1 года)
    Logger.log('Очистка данных старше 1 года...');
    const deletedRecords = cleanOldData(365);
    Logger.log(`✓ Удалено записей: ${deletedRecords}`);
    
    // Очищаем кэш
    Logger.log('Очистка кэша...');
    const cache = CacheService.getScriptCache();
    cache.removeAll(['collectedData', 'processedData']);
    Logger.log('✓ Кэш очищен');
    
    // Проверяем и восстанавливаем триггеры
    Logger.log('Проверка триггеров...');
    setupAutomaticTriggers();
    Logger.log('✓ Триггеры проверены и обновлены');
    
    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== ОБСЛУЖИВАНИЕ ЗАВЕРШЕНО ===');
    Logger.log(`Время выполнения: ${executionTime} секунд`);
    
    return {
      success: true,
      executionTime: executionTime,
      deletedRecords: deletedRecords,
      backupUrl: backupInfo.backupUrl,
      message: `Обслуживание завершено. Удалено ${deletedRecords} записей.`
    };
    
  } catch (error) {
    Logger.log('ОШИБКА обслуживания: ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      message: 'Произошла ошибка при обслуживании системы'
    };
  }
}

// ==================== СПЕЦИАЛЬНЫЕ ФУНКЦИИ ЗАПУСКА ====================

/**
 * Функции для запуска отдельных модулей
 */

// Модуль сбора данных
function runDataCollectionModule() {
  return hourlyDataCollection();
}

// Модуль обработки данных
function runDataProcessingModule() {
  return processAndLinkData();
}

// Модуль аналитики
function runAnalyticsModule() {
  return analyzeData();
}

// Модуль отчетности
function runReportingModule() {
  return generateAllReports();
}

// UTM анализ
function runUtmAnalysisModule() {
  return runUtmAnalysis();
}

// Еженедельные отчеты
function runWeeklyReportsModule() {
  return weeklyReport();
}

// AMO отчеты
function runAmoReportsModule() {
  return createAmoSummaryTable();
}

/**
 * Запуск процедуры восстановления после ошибки
 */
function recoverFromError() {
  const startTime = new Date();
  Logger.log('=== ВОССТАНОВЛЕНИЕ ПОСЛЕ ОШИБКИ ===');
  Logger.log('Время начала: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    // Проверяем состояние системы
    Logger.log('1. Проверка состояния системы...');
    const systemStatus = getSystemStatus();
    
    if (systemStatus.errors.length > 0) {
      Logger.log('Найдены ошибки: ' + JSON.stringify(systemStatus.errors));
    }
    
    // Проверяем доступность данных
    Logger.log('2. Проверка доступности данных...');
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const sheets = spreadsheet.getSheets();
    Logger.log('Доступно листов: ' + sheets.length);
    
    // Восстанавливаем триггеры
    Logger.log('3. Восстановление триггеров...');
    setupAutomaticTriggers();
    
    // Пытаемся выполнить минимальный сбор данных
    Logger.log('4. Тестовый сбор данных...');
    try {
      quickDataCollection();
      Logger.log('✓ Сбор данных работает');
    } catch (collectionError) {
      Logger.log('✗ Ошибка сбора данных: ' + collectionError.toString());
    }
    
    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== ВОССТАНОВЛЕНИЕ ЗАВЕРШЕНО ===');
    Logger.log(`Время выполнения: ${executionTime} секунд`);
    
    return {
      success: true,
      executionTime: executionTime,
      systemStatus: systemStatus,
      message: 'Восстановление завершено'
    };
    
  } catch (error) {
    Logger.log('КРИТИЧЕСКАЯ ОШИБКА восстановления: ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      message: 'Не удалось восстановить систему'
    };
  }
}

/**
 * Функции для тестирования отдельных компонентов
 */

// Тест конфигурации
function testConfig() {
  Logger.log('=== ТЕСТ КОНФИГУРАЦИИ ===');
  try {
    Logger.log('CONFIG.MAIN_SPREADSHEET_ID: ' + CONFIG.MAIN_SPREADSHEET_ID);
    Logger.log('CONFIG.SHEETS: ' + JSON.stringify(CONFIG.SHEETS, null, 2));
    Logger.log('✓ Конфигурация загружена корректно');
    return { success: true };
  } catch (error) {
    Logger.log('✗ Ошибка конфигурации: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// Тест утилит
function testUtils() {
  Logger.log('=== ТЕСТ УТИЛИТ ===');
  try {
    Logger.log('cleanPhone тест: ' + cleanPhone('+7 (123) 456-78-90'));
    Logger.log('formatDate тест: ' + formatDate(new Date()));
    Logger.log('parseNumber тест: ' + parseNumber('1 234,56'));
    Logger.log('✓ Утилиты работают корректно');
    return { success: true };
  } catch (error) {
    Logger.log('✗ Ошибка утилит: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// Тест доступа к таблице
function testSpreadsheetAccess() {
  Logger.log('=== ТЕСТ ДОСТУПА К ТАБЛИЦЕ ===');
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    Logger.log('Название таблицы: ' + spreadsheet.getName());
    
    const sheets = spreadsheet.getSheets();
    Logger.log('Количество листов: ' + sheets.length);
    
    sheets.forEach(sheet => {
      Logger.log(`- ${sheet.getName()}: ${sheet.getMaxRows()} строк`);
    });
    
    Logger.log('✓ Доступ к таблице работает');
    return { success: true, sheetsCount: sheets.length };
  } catch (error) {
    Logger.log('✗ Ошибка доступа к таблице: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Комплексный тест всей системы
 */
function runSystemTest() {
  const startTime = new Date();
  Logger.log('=== КОМПЛЕКСНЫЙ ТЕСТ СИСТЕМЫ ===');
  Logger.log('Время начала: ' + startTime.toLocaleString('ru-RU'));
  
  const results = {
    config: testConfig(),
    utils: testUtils(),
    spreadsheetAccess: testSpreadsheetAccess(),
    systemStatus: getSystemStatus()
  };
  
  const allSuccess = Object.values(results).every(result => 
    result && (result.success === true || !result.hasOwnProperty('success'))
  );
  
  const executionTime = (new Date() - startTime) / 1000;
  Logger.log('=== ТЕСТ ЗАВЕРШЕН ===');
  Logger.log(`Время выполнения: ${executionTime} секунд`);
  Logger.log('Результат: ' + (allSuccess ? '✓ ВСЕ ТЕСТЫ ПРОЙДЕНЫ' : '✗ ЕСТЬ ОШИБКИ'));
  
  return {
    success: allSuccess,
    executionTime: executionTime,
    results: results,
    message: allSuccess ? 'Все тесты пройдены успешно' : 'Обнаружены ошибки в системе'
  };
}

// ==================== СПРАВОЧНАЯ ИНФОРМАЦИЯ ====================

/**
 * Получает справочную информацию о системе
 */
function getSystemInfo() {
  return {
    name: 'Restaurant Analytics System',
    version: '3.2 RU',
    description: 'Комплексная система аналитики ресторана Евгеничь СПб',
    author: 'Restaurant Analytics',
    modules: [
      'Сбор данных (Data Collection)',
      'Обработка данных (Data Processing)', 
      'Аналитика (Analytics)',
      'Отчетность (Reporting)',
      'UTM анализ (UTM Analysis)',
      'Утилиты (Utils)',
      'Конфигурация (Config)'
    ],
    dataSources: [
      'AMO CRM',
      'Reserves RP',
      'Guests RP',
      'Site Requests',
      'Budgets'
    ],
    reports: [
      'Анализ клиентской базы',
      'Анализ трендов и прогнозы',
      'Воронка продаж',
      'Эффективность маркетинговых каналов',
      'UTM анализ кампаний',
      'Еженедельные отчеты',
      'AMO итоговые таблицы'
    ],
    features: [
      'Автоматический сбор данных',
      'Объединение клиентов по телефону',
      'Построение customer journey',
      'ROI анализ маркетинговых каналов',
      'Прогнозирование',
      'Автоматические бэкапы',
      'Система триггеров'
    ],
    lastUpdate: new Date().toLocaleString('ru-RU')
  };
}

/**
 * Получает список всех доступных функций
 */
function getFunctionsList() {
  return {
    mainFunctions: [
      'main() - Полный цикл аналитики',
      'initializeSystem() - Первоначальная настройка',
      'getSystemStatus() - Статус системы',
      'runSystemTest() - Комплексное тестирование'
    ],
    quickFunctions: [
      'quickDataCollection() - Быстрый сбор данных',
      'quickAnalysis() - Быстрый анализ'
    ],
    moduleFunctions: [
      'runDataCollectionModule() - Модуль сбора данных',
      'runDataProcessingModule() - Модуль обработки данных',
      'runAnalyticsModule() - Модуль аналитики',
      'runReportingModule() - Модуль отчетности',
      'runUtmAnalysisModule() - UTM анализ',
      'runWeeklyReportsModule() - Еженедельные отчеты',
      'runAmoReportsModule() - AMO отчеты'
    ],
    setupFunctions: [
      'createMissingSheets() - Создать недостающие листы',
      'initializeSystem() - Полная инициализация системы'
    ],
    maintenanceFunctions: [
      'maintenanceCleanup() - Обслуживание системы',
      'recoverFromError() - Восстановление после ошибки',
      'createDataBackup() - Создание бэкапа',
      'cleanOldData() - Очистка старых данных'
    ],
    testFunctions: [
      'testConfig() - Тест конфигурации',
      'testUtils() - Тест утилит',
      'testSpreadsheetAccess() - Тест доступа к таблице'
    ],
    infoFunctions: [
      'getSystemInfo() - Информация о системе',
      'getFunctionsList() - Список функций'
    ]
  };
}
