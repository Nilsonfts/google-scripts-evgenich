/**
 * Restaurant Analytics System - Module 1: Data Collection
 * Модуль сбора данных из всех источников
 * Автор: Restaurant Analytics
 * Версия: 3.2
 */

// ==================== ОСНОВНЫЕ ФУНКЦИИ ====================

/**
 * Главная функция для hourly триггера
 * Собирает данные из всех источников и кеширует
 */
function hourlyDataCollection() {
  const startTime = new Date();
  Logger.log('=== Начало сбора данных: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    
    // Собираем данные из всех источников
    const allData = {
      timestamp: startTime.toISOString(),
      workingAmo: collectWorkingAmoData(spreadsheet),
      reserves: collectReservesData(spreadsheet),
      guests: collectGuestsData(spreadsheet),
      siteRequests: collectSiteRequestsData(spreadsheet),
      budgets: collectBudgetsData(spreadsheet)
    };

    // Сохраняем статистику
    const stats = {
      workingAmo: allData.workingAmo.data.length - 1, // минус заголовок
      reserves: allData.reserves.length - 1,
      guests: allData.guests.length - 1,
      siteRequests: allData.siteRequests.length - 1,
      budgets: Object.keys(allData.budgets).length,
      executionTime: (new Date() - startTime) / 1000
    };

    // Кешируем данные
    cacheAllData(spreadsheet, allData, stats);

    Logger.log('=== Сбор данных завершен за ' + stats.executionTime + ' секунд');
    Logger.log('Собрано записей: РАБОЧИЙ АМО=' + stats.workingAmo + ', Reserves=' + stats.reserves + ', Guests=' + stats.guests + ', Site=' + stats.siteRequests);

    // Сохраняем время последнего обновления
    PropertiesService.getScriptProperties().setProperty('lastDataCollection', startTime.toISOString());
  } catch (error) {
    Logger.log('ОШИБКА в hourlyDataCollection: ' + error.toString());
    throw error;
  }
}

/**
 * Сбор данных из таблицы РАБОЧИЙ АМО
 */
function collectWorkingAmoData(spreadsheet) {
  try {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.WORKING_AMO);
    if (!sheet) {
      Logger.log('Лист РАБОЧИЙ АМО не найден');
      return { headers: [], data: [[]], columnMap: {} };
    }

    const data = sheet.getDataRange().getValues();
    
    // Первые две строки - заголовки блоков и названия столбцов
    const blockHeaders = data[0] || [];
    const columnHeaders = data[1] || [];

    // Создаём маппинг названий столбцов к их индексам
    const columnMap = {};
    columnHeaders.forEach((header, index) => {
      if (header) {
        columnMap[header.toString().trim()] = index;
      }
    });

    // Данные начинаются с 3-й строки (индекс 2)
    const rowData = data.slice(2);

    Logger.log('РАБОЧИЙ АМО данные: ' + rowData.length + ' записей');
    return {
      blockHeaders: blockHeaders,
      columnHeaders: columnHeaders,
      data: [columnHeaders].concat(rowData), // Добавляем заголовки обратно для совместимости
      columnMap: columnMap
    };
  } catch (error) {
    Logger.log('Ошибка при сборе РАБОЧИЙ АМО данных: ' + error.toString());
    return { headers: [], data: [[]], columnMap: {} };
  }
}

/**
 * Сбор данных из Reserves RP
 */
function collectReservesData(spreadsheet) {
  try {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.RESERVES);
    if (!sheet) {
      Logger.log('Лист RESERVES не найден');
      return [[]];
    }

    const data = sheet.getDataRange().getValues();
    Logger.log('Reserves данные: ' + (data.length - 1) + ' записей');
    return data;
  } catch (error) {
    Logger.log('Ошибка при сборе Reserves данных: ' + error.toString());
    return [[]];
  }
}

/**
 * Сбор данных из Guests RP
 */
function collectGuestsData(spreadsheet) {
  try {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.GUESTS);
    if (!sheet) {
      Logger.log('Лист GUESTS не найден');
      return [[]];
    }

    const data = sheet.getDataRange().getValues();
    Logger.log('Guests данные: ' + (data.length - 1) + ' записей');
    return data;
  } catch (error) {
    Logger.log('Ошибка при сборе Guests данных: ' + error.toString());
    return [[]];
  }
}

/**
 * Сбор данных из заявок с сайта
 */
function collectSiteRequestsData(spreadsheet) {
  try {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SITE_REQUESTS);
    if (!sheet) {
      Logger.log('Лист SITE_REQUESTS не найден');
      return [[]];
    }

    const data = sheet.getDataRange().getValues();
    Logger.log('Site Requests данные: ' + (data.length - 1) + ' записей');
    return data;
  } catch (error) {
    Logger.log('Ошибка при сборе Site Requests данных: ' + error.toString());
    return [[]];
  }
}

/**
 * Сбор данных о бюджетах
 */
function collectBudgetsData(spreadsheet) {
  try {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGETS);
    if (!sheet) {
      Logger.log('Лист BUDGETS не найден');
      return {};
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0]; // Первая строка - заголовки с месяцами
    const budgets = {};

    // Парсим месяцы из заголовков (начиная с колонки C)
    const months = [];
    for (let i = CONFIG.BUDGET_COLUMNS.MONTHS_START; i < headers.length; i++) {
      if (headers[i]) {
        months.push(headers[i]);
      }
    }

    // Собираем данные по каналам
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const channel = row[CONFIG.BUDGET_COLUMNS.CHANNEL];
      const tags = row[CONFIG.BUDGET_COLUMNS.TAGS];

      if (channel) {
        budgets[channel] = {
          tags: tags || '',
          monthly: {}
        };

        // Собираем бюджеты по месяцам
        months.forEach((month, index) => {
          const value = row[CONFIG.BUDGET_COLUMNS.MONTHS_START + index];
          if (value) {
            budgets[channel].monthly[month] = parseNumber(value) || 0;
          }
        });
      }
    }

    Logger.log('Budgets данные: ' + Object.keys(budgets).length + ' каналов');
    return budgets;
  } catch (error) {
    Logger.log('Ошибка при сборе Budgets данных: ' + error.toString());
    return {};
  }
}

/**
 * Кеширование всех данных в служебный лист
 */
function cacheAllData(spreadsheet, allData, stats) {
  try {
    // Создаем или очищаем лист метаданных
    let metaSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.CACHE_METADATA);
    if (!metaSheet) {
      metaSheet = spreadsheet.insertSheet(CONFIG.SHEETS.CACHE_METADATA);
      metaSheet.hideSheet(); // Скрываем служебный лист
    } else {
      metaSheet.clear();
    }

    // Сохраняем метаданные
    const metadata = [
      ['Параметр', 'Значение'],
      ['Время обновления', allData.timestamp],
      ['РАБОЧИЙ АМО записей', stats.workingAmo],
      ['Reserves записей', stats.reserves],
      ['Guests записей', stats.guests],
      ['Site Requests записей', stats.siteRequests],
      ['Каналов с бюджетами', stats.budgets],
      ['Время выполнения (сек)', stats.executionTime]
    ];
    metaSheet.getRange(1, 1, metadata.length, 2).setValues(metadata);

    // Сохраняем данные в Properties для быстрого доступа другими модулями
    const properties = PropertiesService.getScriptProperties();
    // Сохраняем статистику
    properties.setProperty('dataStats', JSON.stringify(stats));

    // Для больших данных используем Cache Service
    const cache = CacheService.getScriptCache();
    // Кешируем бюджеты (они небольшие)
    cache.put('budgetsData', JSON.stringify(allData.budgets), 3600); // 1 час

    // Кешируем маппинг колонок для быстрого доступа к данным
    cache.put('amoColumnMap', JSON.stringify(allData.workingAmo.columnMap), 3600);

    Logger.log('Данные успешно закешированы');
  } catch (error) {
    Logger.log('Ошибка при кешировании: ' + error.toString());
  }
}

/**
 * Получение значения ячейки из строки данных РАБОЧИЙ АМО
 * Использует маппинг блоков для доступа к данным
 */
function getAmoValue(row, fieldName, columnMap) {
  // Проверяем, есть ли такое поле в нашей конфигурации блоков
  for (const blockKey in CONFIG.WORKING_AMO_BLOCKS) {
    const block = CONFIG.WORKING_AMO_BLOCKS[blockKey];
    for (const key in block) {
      if (block[key] === fieldName) {
        // Нашли поле в конфигурации, теперь получаем его индекс из маппинга колонок
        const colIndex = columnMap[fieldName];
        if (colIndex !== undefined && row[colIndex] !== undefined) {
          return row[colIndex];
        }
        break;
      }
    }
  }
  return ''; // Если поле не найдено
}

// ==================== ФУНКЦИИ УСТАНОВКИ И УПРАВЛЕНИЯ ====================

/**
 * Установка hourly триггера
 */
function setupHourlyTrigger() {
  // Удаляем существующие триггеры
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'hourlyDataCollection') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Создаем новый триггер
  ScriptApp.newTrigger('hourlyDataCollection')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('Hourly триггер установлен для hourlyDataCollection');

  // Запускаем первый сбор данных
  hourlyDataCollection();
}

/**
 * Ручной запуск сбора данных
 */
function manualDataCollection() {
  hourlyDataCollection();
  SpreadsheetApp.getActiveSpreadsheet().toast('Сбор данных завершен', 'Успех', 3);
}

/**
 * Получение статистики последнего сбора
 */
function getLastCollectionStats() {
  const properties = PropertiesService.getScriptProperties();
  const lastRun = properties.getProperty('lastDataCollection');
  const stats = properties.getProperty('dataStats');

  if (lastRun && stats) {
    const lastRunDate = new Date(lastRun);
    const statsObj = JSON.parse(stats);

    Logger.log('=== Последний сбор данных ===');
    Logger.log('Время: ' + lastRunDate.toLocaleString('ru-RU'));
    Logger.log('РАБОЧИЙ АМО записей: ' + statsObj.workingAmo);
    Logger.log('Reserves записей: ' + statsObj.reserves);
    Logger.log('Guests записей: ' + statsObj.guests);
    Logger.log('Site Requests записей: ' + statsObj.siteRequests);
    Logger.log('Каналов с бюджетами: ' + statsObj.budgets);
    Logger.log('Время выполнения: ' + statsObj.executionTime + ' сек');

    return {
      lastRun: lastRunDate,
      stats: statsObj
    };
  } else {
    Logger.log('Данные еще не собирались');
    return null;
  }
}
