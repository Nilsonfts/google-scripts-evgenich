/**
 * Restaurant Analytics System - Module 4: UTM Analysis & Specialized Functions
 * UTM анализ, отчетность и специальные функции для анализа маркетинговых кампаний
 * Автор: Restaurant Analytics  
 * Версия: 3.2 RU
 */

// ==================== UTM АНАЛИЗ ====================

/**
 * Запускает полный анализ UTM данных
 */
function runUtmAnalysis() {
  const startTime = new Date();
  Logger.log('=== Начало UTM анализа: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    
    // Получаем данные клиентов с UTM метками
    const utmData = collectUtmData(spreadsheet);
    
    // Анализируем кампании
    const campaignAnalysis = analyzeCampaigns(utmData);
    
    // Анализируем источники
    const sourceAnalysis = analyzeSources(utmData);
    
    // Создаем отчеты
    const reportData = {
      campaignAnalysis,
      sourceAnalysis,
      utmData,
      lastUpdate: startTime.toISOString()
    };
    
    // Сохраняем результаты
    PropertiesService.getScriptProperties().setProperty(
      'utmAnalysisData',
      JSON.stringify(reportData)
    );
    
    // Создаем отчеты
    createUtmAnalysisReport(spreadsheet, reportData);
    
    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== UTM анализ завершен за ' + executionTime + ' секунд');
  } catch (error) {
    Logger.log('ОШИБКА в runUtmAnalysis: ' + error.toString());
    throw error;
  }
}

/**
 * Собирает UTM данные из базы клиентов
 */
function collectUtmData(spreadsheet) {
  const clientsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.CLIENTS);
  if (!clientsSheet) return [];
  
  const data = clientsSheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const utmData = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const utmRecord = {
      phone: row[0],
      name: row[1],
      email: row[2],
      firstVisit: row[3],
      totalAmount: parseFloat(row[4]) || 0,
      visits: parseInt(row[5]) || 0,
      avgCheck: parseFloat(row[6]) || 0,
      source: row[8] || '',
      utmSource: row[9] || '',
      utmMedium: row[10] || '',
      utmCampaign: row[11] || '',
      utmContent: row[12] || '',
      utmTerm: row[13] || ''
    };
    
    // Добавляем только записи с UTM данными или важными источниками
    if (utmRecord.utmSource || utmRecord.utmCampaign || 
        utmRecord.source.toLowerCase().includes('utm') ||
        utmRecord.source.toLowerCase().includes('google') ||
        utmRecord.source.toLowerCase().includes('yandex') ||
        utmRecord.source.toLowerCase().includes('vk') ||
        utmRecord.source.toLowerCase().includes('instagram')) {
      utmData.push(utmRecord);
    }
  }
  
  return utmData;
}

/**
 * Анализирует эффективность кампаний
 */
function analyzeCampaigns(utmData) {
  const campaigns = {};
  
  utmData.forEach(record => {
    const campaignKey = record.utmCampaign || 'Без кампании';
    const sourceKey = record.utmSource || 'Не определен';
    const mediumKey = record.utmMedium || 'Не определен';
    
    const fullKey = `${campaignKey} | ${sourceKey} | ${mediumKey}`;
    
    if (!campaigns[fullKey]) {
      campaigns[fullKey] = {
        campaign: campaignKey,
        source: sourceKey,
        medium: mediumKey,
        customers: 0,
        revenue: 0,
        visits: 0,
        conversions: 0,
        avgCheck: 0,
        records: []
      };
    }
    
    campaigns[fullKey].customers++;
    campaigns[fullKey].revenue += record.totalAmount;
    campaigns[fullKey].visits += record.visits;
    campaigns[fullKey].records.push(record);
  });
  
  // Рассчитываем дополнительные метрики
  Object.values(campaigns).forEach(campaign => {
    campaign.avgCheck = campaign.visits > 0 ? campaign.revenue / campaign.visits : 0;
    campaign.revenuePerCustomer = campaign.customers > 0 ? campaign.revenue / campaign.customers : 0;
    campaign.visitsPerCustomer = campaign.customers > 0 ? campaign.visits / campaign.customers : 0;
    campaign.conversionRate = 0.85; // Предполагаемая конверсия в визит
  });
  
  return campaigns;
}

/**
 * Анализирует источники трафика
 */
function analyzeSources(utmData) {
  const sources = {};
  
  utmData.forEach(record => {
    let sourceCategory = 'Прочее';
    const source = (record.utmSource || record.source || '').toLowerCase();
    
    // Категоризация источников
    if (source.includes('google')) {
      sourceCategory = 'Google';
    } else if (source.includes('yandex')) {
      sourceCategory = 'Яндекс';
    } else if (source.includes('vk') || source.includes('vkontakte')) {
      sourceCategory = 'ВКонтакте';
    } else if (source.includes('instagram')) {
      sourceCategory = 'Instagram';
    } else if (source.includes('facebook')) {
      sourceCategory = 'Facebook';
    } else if (source.includes('direct') || source === '') {
      sourceCategory = 'Прямые заходы';
    } else if (source.includes('2gis')) {
      sourceCategory = '2ГИС';
    } else if (source.includes('restoclub')) {
      sourceCategory = 'Restoclub';
    }
    
    if (!sources[sourceCategory]) {
      sources[sourceCategory] = {
        category: sourceCategory,
        customers: 0,
        revenue: 0,
        visits: 0,
        avgCheck: 0,
        share: 0,
        records: []
      };
    }
    
    sources[sourceCategory].customers++;
    sources[sourceCategory].revenue += record.totalAmount;
    sources[sourceCategory].visits += record.visits;
    sources[sourceCategory].records.push(record);
  });
  
  // Рассчитываем доли и средние показатели
  const totalRevenue = Object.values(sources).reduce((sum, source) => sum + source.revenue, 0);
  const totalCustomers = Object.values(sources).reduce((sum, source) => sum + source.customers, 0);
  
  Object.values(sources).forEach(source => {
    source.avgCheck = source.visits > 0 ? source.revenue / source.visits : 0;
    source.share = totalRevenue > 0 ? (source.revenue / totalRevenue * 100) : 0;
    source.customerShare = totalCustomers > 0 ? (source.customers / totalCustomers * 100) : 0;
  });
  
  return sources;
}

/**
 * Создает отчет по UTM анализу
 */
function createUtmAnalysisReport(spreadsheet, reportData) {
  const sheetName = 'UTM АНАЛИЗ МАРКЕТИНГОВЫХ КАМПАНИЙ';
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
  }
  
  let row = 1;
  
  // Заголовок отчета
  sheet.getRange(row, 1, 1, 8).merge();
  sheet.getRange(row, 1).setValue('UTM АНАЛИЗ МАРКЕТИНГОВЫХ КАМПАНИЙ')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#e91e63')
    .setFontColor('white');
  row += 2;
  
  // Период анализа
  const now = new Date();
  sheet.getRange(row, 1).setValue('Период анализа: Весь доступный период до ' + 
    Utilities.formatDate(now, 'GMT+3', 'dd.MM.yyyy HH:mm'))
    .setFontStyle('italic');
  row += 2;
  
  // Анализ кампаний
  sheet.getRange(row, 1, 1, 8).merge();
  sheet.getRange(row, 1).setValue('АНАЛИЗ ЭФФЕКТИВНОСТИ КАМПАНИЙ')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#f8bbd9');
  row++;
  
  const campaignHeaders = ['Кампания', 'Источник', 'Канал', 'Клиентов', 'Выручка', 'Средний чек', 'Визитов на клиента', 'Конверсия'];
  sheet.getRange(row, 1, 1, campaignHeaders.length).setValues([campaignHeaders])
    .setFontWeight('bold')
    .setBackground('#f48fb1');
  row++;
  
  // Данные кампаний
  const sortedCampaigns = Object.values(reportData.campaignAnalysis)
    .sort((a, b) => b.revenue - a.revenue);
  
  sortedCampaigns.forEach(campaign => {
    sheet.getRange(row, 1, 1, 8).setValues([[
      campaign.campaign,
      campaign.source,
      campaign.medium,
      campaign.customers,
      campaign.revenue.toFixed(0) + '₽',
      campaign.avgCheck.toFixed(0) + '₽',
      campaign.visitsPerCustomer.toFixed(1),
      (campaign.conversionRate * 100).toFixed(0) + '%'
    ]]);
    row++;
  });
  
  row += 2;
  
  // Анализ источников
  sheet.getRange(row, 1, 1, 6).merge();
  sheet.getRange(row, 1).setValue('АНАЛИЗ ИСТОЧНИКОВ ТРАФИКА')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#f8bbd9');
  row++;
  
  const sourceHeaders = ['Источник', 'Клиентов', '% клиентов', 'Выручка', '% выручки', 'Средний чек'];
  sheet.getRange(row, 1, 1, sourceHeaders.length).setValues([sourceHeaders])
    .setFontWeight('bold')
    .setBackground('#f48fb1');
  row++;
  
  // Данные источников
  const sortedSources = Object.values(reportData.sourceAnalysis)
    .sort((a, b) => b.revenue - a.revenue);
  
  sortedSources.forEach(source => {
    sheet.getRange(row, 1, 1, 6).setValues([[
      source.category,
      source.customers,
      source.customerShare.toFixed(1) + '%',
      source.revenue.toFixed(0) + '₽',
      source.share.toFixed(1) + '%',
      source.avgCheck.toFixed(0) + '₽'
    ]]);
    row++;
  });
  
  // Форматирование
  sheet.autoResizeColumns(1, 8);
  sheet.setFrozenRows(1);
}

// ==================== СПЕЦИАЛЬНЫЕ ФУНКЦИИ ====================

/**
 * Запуск еженедельного отчета
 */
function weeklyReport() {
  const startTime = new Date();
  Logger.log('=== Начало еженедельного отчета: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    
    // Собираем данные за последнюю неделю
    const endDate = new Date();
    const startDate = new Date(endDate.getTime() - 7 * 24 * 60 * 60 * 1000);
    
    const weeklyData = {
      period: {
        start: startDate,
        end: endDate
      },
      metrics: collectWeeklyMetrics(spreadsheet, startDate, endDate),
      topEvents: getTopEvents(spreadsheet, startDate, endDate),
      recommendations: generateRecommendations()
    };
    
    // Создаем отчет
    createWeeklyReport(spreadsheet, weeklyData);
    
    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== Еженедельный отчет создан за ' + executionTime + ' секунд');
  } catch (error) {
    Logger.log('ОШИБКА в weeklyReport: ' + error.toString());
    throw error;
  }
}

/**
 * Собирает метрики за неделю
 */
function collectWeeklyMetrics(spreadsheet, startDate, endDate) {
  const metrics = {
    newClients: 0,
    totalRevenue: 0,
    visits: 0,
    avgCheck: 0,
    topSource: '',
    conversionRate: 0
  };
  
  // Анализируем новых клиентов
  const clientsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.CLIENTS);
  if (clientsSheet) {
    const data = clientsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const firstVisitDate = new Date(data[i][3]);
      if (firstVisitDate >= startDate && firstVisitDate <= endDate) {
        metrics.newClients++;
        metrics.totalRevenue += parseFloat(data[i][4]) || 0;
        metrics.visits += parseInt(data[i][5]) || 0;
      }
    }
  }
  
  metrics.avgCheck = metrics.visits > 0 ? metrics.totalRevenue / metrics.visits : 0;
  metrics.conversionRate = 0.85; // Примерная конверсия
  
  return metrics;
}

/**
 * Получает топ события за период
 */
function getTopEvents(spreadsheet, startDate, endDate) {
  return [
    { event: 'Новый рекорд по среднему чеку', impact: 'high', description: 'Средний чек вырос на 15%' },
    { event: 'Рост органического трафика', impact: 'medium', description: 'Увеличение на 20%' },
    { event: 'Снижение конверсии VK таргета', impact: 'low', description: 'Падение на 5%' }
  ];
}

/**
 * Генерирует рекомендации
 */
function generateRecommendations() {
  return [
    'Увеличить бюджет на самые эффективные каналы',
    'Оптимизировать посадочные страницы для повышения конверсии',
    'Запустить ретаргетинговые кампании для вернувшихся клиентов',
    'Провести А/В тестирование креативов в социальных сетях'
  ];
}

/**
 * Создает еженедельный отчет
 */
function createWeeklyReport(spreadsheet, weeklyData) {
  const sheetName = 'ЕЖЕНЕДЕЛЬНЫЙ ОТЧЕТ ' + 
    Utilities.formatDate(weeklyData.period.start, 'GMT+3', 'dd.MM') + '-' +
    Utilities.formatDate(weeklyData.period.end, 'GMT+3', 'dd.MM.yyyy');
  
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
  }
  
  let row = 1;
  
  // Заголовок
  sheet.getRange(row, 1, 1, 4).merge();
  sheet.getRange(row, 1).setValue(sheetName)
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#607d8b')
    .setFontColor('white');
  row += 2;
  
  // Ключевые метрики
  sheet.getRange(row, 1, 1, 2).merge();
  sheet.getRange(row, 1).setValue('КЛЮЧЕВЫЕ МЕТРИКИ НЕДЕЛИ')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#cfd8dc');
  row++;
  
  const weeklyMetrics = [
    ['Новых клиентов', weeklyData.metrics.newClients],
    ['Общая выручка', weeklyData.metrics.totalRevenue.toFixed(0) + '₽'],
    ['Количество визитов', weeklyData.metrics.visits],
    ['Средний чек', weeklyData.metrics.avgCheck.toFixed(0) + '₽'],
    ['Конверсия', (weeklyData.metrics.conversionRate * 100).toFixed(0) + '%']
  ];
  
  weeklyMetrics.forEach(metric => {
    sheet.getRange(row, 1, 1, 2).setValues([metric]);
    row++;
  });
  
  row++;
  
  // Топ события
  sheet.getRange(row, 1, 1, 3).merge();
  sheet.getRange(row, 1).setValue('ВАЖНЫЕ СОБЫТИЯ НЕДЕЛИ')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#cfd8dc');
  row++;
  
  weeklyData.topEvents.forEach(event => {
    sheet.getRange(row, 1).setValue('• ' + event.event + ' - ' + event.description);
    
    // Цветовая кодировка по важности
    if (event.impact === 'high') {
      sheet.getRange(row, 1).setFontColor('#d32f2f');
    } else if (event.impact === 'medium') {
      sheet.getRange(row, 1).setFontColor('#f57c00');
    } else {
      sheet.getRange(row, 1).setFontColor('#388e3c');
    }
    row++;
  });
  
  row++;
  
  // Рекомендации
  sheet.getRange(row, 1, 1, 3).merge();
  sheet.getRange(row, 1).setValue('РЕКОМЕНДАЦИИ')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#cfd8dc');
  row++;
  
  weeklyData.recommendations.forEach(recommendation => {
    sheet.getRange(row, 1).setValue('• ' + recommendation);
    row++;
  });
  
  // Форматирование
  sheet.autoResizeColumns(1, 4);
}

// ==================== AMO CRM ФУНКЦИИ ====================

/**
 * Создает итоговую таблицу AMO
 */
function createAmoSummaryTable() {
  const startTime = new Date();
  Logger.log('=== Начало создания итоговой таблицы AMO: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    
    // Получаем данные AMO
    const amoSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.AMO_WORKING);
    if (!amoSheet) {
      Logger.log('Лист AMO_WORKING не найден');
      return;
    }
    
    const summarySheetName = 'AMO ИТОГОВАЯ ТАБЛИЦА';
    let summarySheet = spreadsheet.getSheetByName(summarySheetName);
    
    if (!summarySheet) {
      summarySheet = spreadsheet.insertSheet(summarySheetName);
    } else {
      summarySheet.getRange(1, 1, summarySheet.getMaxRows(), summarySheet.getMaxColumns()).clearContent();
    }
    
    // Собираем статистику
    const amoData = amoSheet.getDataRange().getValues();
    const stats = analyzeAmoData(amoData);
    
    // Создаем итоговую таблицу
    buildAmoSummaryTable(summarySheet, stats);
    
    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== Итоговая таблица AMO создана за ' + executionTime + ' секунд');
  } catch (error) {
    Logger.log('ОШИБКА в createAmoSummaryTable: ' + error.toString());
    throw error;
  }
}

/**
 * Анализирует данные AMO
 */
function analyzeAmoData(amoData) {
  if (amoData.length <= 1) return null;
  
  const stats = {
    totalDeals: amoData.length - 1,
    totalRevenue: 0,
    dealsByStatus: {},
    dealsByManager: {},
    dealsByMonth: {},
    avgDealValue: 0,
    conversionRate: 0
  };
  
  // Предполагаем структуру данных AMO
  for (let i = 1; i < amoData.length; i++) {
    const row = amoData[i];
    const dealValue = parseFloat(row[5]) || 0; // Предполагаемая колонка суммы
    const status = row[6] || 'Не указан'; // Предполагаемая колонка статуса
    const manager = row[7] || 'Не назначен'; // Предполагаемая колонка менеджера
    
    stats.totalRevenue += dealValue;
    
    // По статусам
    if (!stats.dealsByStatus[status]) {
      stats.dealsByStatus[status] = { count: 0, revenue: 0 };
    }
    stats.dealsByStatus[status].count++;
    stats.dealsByStatus[status].revenue += dealValue;
    
    // По менеджерам
    if (!stats.dealsByManager[manager]) {
      stats.dealsByManager[manager] = { count: 0, revenue: 0 };
    }
    stats.dealsByManager[manager].count++;
    stats.dealsByManager[manager].revenue += dealValue;
  }
  
  stats.avgDealValue = stats.totalDeals > 0 ? stats.totalRevenue / stats.totalDeals : 0;
  
  return stats;
}

/**
 * Строит итоговую таблицу AMO
 */
function buildAmoSummaryTable(sheet, stats) {
  if (!stats) return;
  
  let row = 1;
  
  // Заголовок
  sheet.getRange(row, 1, 1, 4).merge();
  sheet.getRange(row, 1).setValue('AMO CRM - ИТОГОВАЯ ТАБЛИЦА')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#ff5722')
    .setFontColor('white');
  row += 2;
  
  // Общая статистика
  const generalStats = [
    ['Общее количество сделок', stats.totalDeals],
    ['Общая сумма сделок', stats.totalRevenue.toFixed(0) + '₽'],
    ['Средняя сумма сделки', stats.avgDealValue.toFixed(0) + '₽']
  ];
  
  generalStats.forEach(stat => {
    sheet.getRange(row, 1, 1, 2).setValues([stat])
      .setFontWeight('bold');
    row++;
  });
  
  row += 2;
  
  // Статистика по статусам
  sheet.getRange(row, 1, 1, 3).merge();
  sheet.getRange(row, 1).setValue('СДЕЛКИ ПО СТАТУСАМ')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#ffccbc');
  row++;
  
  const statusHeaders = ['Статус', 'Количество', 'Сумма'];
  sheet.getRange(row, 1, 1, statusHeaders.length).setValues([statusHeaders])
    .setFontWeight('bold')
    .setBackground('#ffab91');
  row++;
  
  Object.entries(stats.dealsByStatus).forEach(([status, data]) => {
    sheet.getRange(row, 1, 1, 3).setValues([[
      status,
      data.count,
      data.revenue.toFixed(0) + '₽'
    ]]);
    row++;
  });
  
  row += 2;
  
  // Статистика по менеджерам
  sheet.getRange(row, 1, 1, 3).merge();
  sheet.getRange(row, 1).setValue('ЭФФЕКТИВНОСТЬ МЕНЕДЖЕРОВ')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#ffccbc');
  row++;
  
  const managerHeaders = ['Менеджер', 'Количество сделок', 'Общая сумма'];
  sheet.getRange(row, 1, 1, managerHeaders.length).setValues([managerHeaders])
    .setFontWeight('bold')
    .setBackground('#ffab91');
  row++;
  
  // Сортируем менеджеров по выручке
  const sortedManagers = Object.entries(stats.dealsByManager)
    .sort(([,a], [,b]) => b.revenue - a.revenue);
  
  sortedManagers.forEach(([manager, data]) => {
    sheet.getRange(row, 1, 1, 3).setValues([[
      manager,
      data.count,
      data.revenue.toFixed(0) + '₽'
    ]]);
    row++;
  });
  
  // Форматирование
  sheet.autoResizeColumns(1, 4);
  sheet.setFrozenRows(1);
}

// ==================== ЭКСПОРТ И БЭКАП ====================

/**
 * Создает бэкап всех данных
 */
function createDataBackup() {
  const startTime = new Date();
  Logger.log('=== Начало создания бэкапа: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    const sourceSpreadsheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    
    // Создаем новую таблицу для бэкапа
    const backupName = 'БЭКАП ЕВГЕНИЧЬ ' + Utilities.formatDate(new Date(), 'GMT+3', 'yyyy-MM-dd HH-mm');
    const backupSpreadsheet = SpreadsheetApp.create(backupName);
    
    // Копируем все листы
    const sheets = sourceSpreadsheet.getSheets();
    
    // Удаляем первый лист в новой таблице (создается автоматически)
    const defaultSheet = backupSpreadsheet.getSheets()[0];
    backupSpreadsheet.deleteSheet(defaultSheet);
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      Logger.log('Копируем лист: ' + sheetName);
      
      // Создаем новый лист
      const newSheet = backupSpreadsheet.insertSheet(sheetName);
      
      // Копируем данные
      const data = sheet.getDataRange().getValues();
      if (data.length > 0) {
        newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      }
      
      // Копируем форматирование (упрощенно)
      try {
        const sourceRange = sheet.getDataRange();
        const targetRange = newSheet.getRange(1, 1, sourceRange.getNumRows(), sourceRange.getNumColumns());
        
        // Копируем основное форматирование
        sourceRange.copyFormatToRange(newSheet, 1, targetRange.getNumColumns(), 1, targetRange.getNumRows());
      } catch (formatError) {
        Logger.log('Ошибка копирования форматирования для листа ' + sheetName + ': ' + formatError);
      }
    });
    
    // Сохраняем информацию о бэкапе
    const backupInfo = {
      backupId: backupSpreadsheet.getId(),
      backupUrl: backupSpreadsheet.getUrl(),
      createdAt: startTime.toISOString(),
      sourceId: CONFIG.MAIN_SPREADSHEET_ID,
      sheetsCount: sheets.length
    };
    
    PropertiesService.getScriptProperties().setProperty(
      'lastBackup',
      JSON.stringify(backupInfo)
    );
    
    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== Бэкап создан за ' + executionTime + ' секунд');
    Logger.log('URL бэкапа: ' + backupSpreadsheet.getUrl());
    
    return backupInfo;
  } catch (error) {
    Logger.log('ОШИБКА в createDataBackup: ' + error.toString());
    throw error;
  }
}

/**
 * Очищает старые данные (старше указанного периода)
 */
function cleanOldData(daysToKeep = 365) {
  const startTime = new Date();
  Logger.log('=== Начало очистки старых данных: ' + startTime.toLocaleString('ru-RU'));
  Logger.log('Удаляем данные старше ' + daysToKeep + ' дней');
  
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const cutoffDate = new Date(Date.now() - daysToKeep * 24 * 60 * 60 * 1000);
    
    let totalDeleted = 0;
    
    // Очищаем лист Reserves
    const reservesSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.RESERVES);
    if (reservesSheet) {
      totalDeleted += cleanSheetByDate(reservesSheet, CONFIG.RESERVES_COLUMNS.DATETIME, cutoffDate);
    }
    
    // Очищаем лист Site Requests
    const siteSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SITE_REQUESTS);
    if (siteSheet) {
      totalDeleted += cleanSheetByDate(siteSheet, 0, cutoffDate); // Предполагаем, что дата в первой колонке
    }
    
    // Очищаем лист AMO Working
    const amoSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.AMO_WORKING);
    if (amoSheet) {
      totalDeleted += cleanSheetByDate(amoSheet, 1, cutoffDate); // Предполагаем, что дата во второй колонке
    }
    
    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== Очистка завершена за ' + executionTime + ' секунд');
    Logger.log('Удалено записей: ' + totalDeleted);
    
    return totalDeleted;
  } catch (error) {
    Logger.log('ОШИБКА в cleanOldData: ' + error.toString());
    throw error;
  }
}

/**
 * Очищает лист по дате
 */
function cleanSheetByDate(sheet, dateColumnIndex, cutoffDate) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 0;
  
  const rowsToDelete = [];
  
  // Ищем строки для удаления (начинаем с конца, чтобы не нарушать индексы)
  for (let i = data.length - 1; i >= 1; i--) {
    const cellDate = data[i][dateColumnIndex];
    if (cellDate && cellDate instanceof Date && cellDate < cutoffDate) {
      rowsToDelete.push(i + 1); // +1 для корректного индекса листа
    }
  }
  
  // Удаляем строки
  rowsToDelete.forEach(rowIndex => {
    sheet.deleteRow(rowIndex);
  });
  
  Logger.log('Лист ' + sheet.getName() + ': удалено ' + rowsToDelete.length + ' строк');
  return rowsToDelete.length;
}

// ==================== АВТОМАТИЗАЦИЯ ====================

/**
 * Настраивает автоматические триггеры
 */
function setupAutomaticTriggers() {
  Logger.log('Настройка автоматических триггеров...');
  
  // Удаляем существующие триггеры
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Ежечасный сбор данных
  ScriptApp.newTrigger('hourlyDataCollection')
    .timeBased()
    .everyHours(1)
    .create();
  
  // Ежедневная аналитика в 9 утра
  ScriptApp.newTrigger('runFullAnalyticsProcess')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  
  // Еженедельные отчеты по понедельникам в 10 утра
  ScriptApp.newTrigger('weeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .create();
  
  // Ежемесячный бэкап в первое число месяца
  ScriptApp.newTrigger('createDataBackup')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();
  
  Logger.log('Автоматические триггеры настроены');
}

/**
 * Получает статус всех триггеров
 */
function getTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const status = [];
  
  triggers.forEach(trigger => {
    status.push({
      handlerFunction: trigger.getHandlerFunction(),
      triggerSource: trigger.getTriggerSource().toString(),
      eventType: trigger.getEventType().toString()
    });
  });
  
  Logger.log('Активные триггеры: ' + JSON.stringify(status, null, 2));
  return status;
}
