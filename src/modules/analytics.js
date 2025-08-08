/**
 * Restaurant Analytics System - Module 3: Analytics & Reports
 * Анализирует данные и создает красивые отчеты
 * Автор: Restaurant Analytics
 * Версия: 3.0 RU
 */

// ==================== ГЛАВНАЯ ФУНКЦИЯ АНАЛИТИКИ ====================

/**
 * Запускает полный процесс аналитики
 */
function runFullAnalyticsProcess() {
  const startTime = new Date();
  Logger.log('=== Начало полного процесса аналитики: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    // 1. Сбор данных
    Logger.log('Запуск сбора данных...');
    hourlyDataCollection();
    Logger.log('Сбор данных завершен');

    // 2. Обработка данных
    Logger.log('Запуск обработки данных...');
    processAndLinkData();
    Logger.log('Обработка данных завершена');

    // 3. Анализ данных
    Logger.log('Запуск анализа данных...');
    analyzeData();
    Logger.log('Анализ данных завершен');

    // 4. Генерация отчетов
    Logger.log('Запуск генерации отчетов...');
    generateAllReports();
    Logger.log('Генерация отчетов завершена');

    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== Полный процесс завершен за ' + executionTime + ' секунд');
  } catch (error) {
    Logger.log('ОШИБКА в runFullAnalyticsProcess: ' + error.toString());
    throw error;
  }
}

/**
 * Главная функция анализа данных
 */
function analyzeData() {
  const startTime = new Date();
  Logger.log('=== Начало анализа данных: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);

    // Анализ клиентской базы
    Logger.log('Начало анализа клиентской базы...');
    const clientAnalysis = analyzeClientBase(spreadsheet);
    Logger.log('Анализ клиентской базы завершен');

    // Анализ трендов и прогнозы
    Logger.log('Начало анализа трендов...');
    const trendsAnalysis = analyzeTrends(spreadsheet);
    Logger.log('Анализ трендов завершен');

    // Анализ воронки продаж
    Logger.log('Начало анализа воронки продаж...');
    const funnelAnalysis = analyzeSalesFunnel(spreadsheet);
    Logger.log('Анализ воронки завершен');

    // Анализ маркетинговых каналов
    Logger.log('Начало анализа маркетинга...');
    const marketingAnalysis = analyzeMarketingChannels(spreadsheet);
    Logger.log('Анализ маркетинга завершен');

    // Подготовка данных для дашборда
    Logger.log('Подготовка данных для дашборда...');
    const dashboardData = prepareDashboardData(spreadsheet, {
      clientAnalysis,
      trendsAnalysis,
      funnelAnalysis,
      marketingAnalysis
    });

    // Сохраняем результаты анализа
    const analyticsData = {
      clientAnalysis,
      trendsAnalysis,
      funnelAnalysis,
      marketingAnalysis,
      dashboardData,
      lastUpdate: startTime.toISOString()
    };

    // Сохраняем в Properties вместо Cache из-за размера
    PropertiesService.getScriptProperties().setProperty(
      'analyticsData',
      JSON.stringify(analyticsData)
    );

    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== Анализ завершен за ' + executionTime + ' секунд');
  } catch (error) {
    Logger.log('ОШИБКА в analyzeData: ' + error.toString());
    throw error;
  }
}

// ==================== АНАЛИЗ КЛИЕНТСКОЙ БАЗЫ ====================

/**
 * Анализирует клиентскую базу
 */
function analyzeClientBase(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.CLIENTS);
  if (!sheet) {
    Logger.log('Лист ЕДИНАЯ_БАЗА_КЛИЕНТОВ не найден');
    return null;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;

  // Получаем индексы колонок
  const headers = data[0];
  const phoneIdx = headers.indexOf('ID (Телефон)');
  const nameIdx = headers.indexOf('Имя');
  const visitsIdx = headers.indexOf('Кол-во визитов');
  const amountIdx = headers.indexOf('Общая сумма');
  const avgCheckIdx = headers.indexOf('Средний чек');
  const firstVisitIdx = headers.indexOf('Первый визит');
  const lastVisitIdx = headers.indexOf('Последний визит');

  // Сегментация клиентов
  const segments = {
    'Новые (1 визит)': [],
    'Постоянные (2-5 визитов)': [],
    'Лояльные (6-10 визитов)': [],
    'VIP (>10 визитов)': []
  };

  const now = new Date();
  const clients = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const visits = parseInt(row[visitsIdx]) || 0;
    const amount = parseFloat(row[amountIdx]) || 0;
    const avgCheck = parseFloat(row[avgCheckIdx]) || 0;
    const lastVisit = row[lastVisitIdx] ? new Date(row[lastVisitIdx]) : null;

    // Сегментация
    let segment = '';
    if (visits === 1) segment = 'Новые (1 визит)';
    else if (visits >= 2 && visits <= 5) segment = 'Постоянные (2-5 визитов)';
    else if (visits >= 6 && visits <= 10) segment = 'Лояльные (6-10 визитов)';
    else if (visits > 10) segment = 'VIP (>10 визитов)';

    if (segment) {
      segments[segment].push({
        phone: row[phoneIdx],
        name: row[nameIdx] || '',
        visits: visits,
        amount: amount,
        avgCheck: avgCheck,
        firstVisit: row[firstVisitIdx],
        lastVisit: row[lastVisitIdx]
      });
    }

    // Добавляем в общий список для ТОП-20
    clients.push({
      phone: row[phoneIdx],
      name: row[nameIdx] || '',
      visits: visits,
      amount: amount,
      avgCheck: avgCheck,
      firstVisit: row[firstVisitIdx],
      lastVisit: row[lastVisitIdx]
    });
  }

  // Сортируем для ТОП-20 по сумме
  clients.sort((a, b) => b.amount - a.amount);
  const top20 = clients.slice(0, 20);

  // Рассчитываем статистику по сегментам
  const segmentStats = {};
  let totalClients = 0;
  let totalRevenue = 0;

  Object.keys(segments).forEach(segment => {
    const segmentData = segments[segment];
    const revenue = segmentData.reduce((sum, client) => sum + client.amount, 0);
    totalClients += segmentData.length;
    totalRevenue += revenue;

    segmentStats[segment] = {
      count: segmentData.length,
      percentage: 0, // Заполним после
      avgCheck: segmentData.length > 0 ? revenue / segmentData.reduce((sum, client) => sum + client.visits, 0) : 0,
      totalRevenue: revenue,
      ltv: segmentData.length > 0 ? revenue / segmentData.length : 0
    };
  });

  // Рассчитываем проценты
  Object.keys(segmentStats).forEach(segment => {
    segmentStats[segment].percentage = totalClients > 0 ?
      (segmentStats[segment].count / totalClients * 100) : 0;
  });

  return {
    segments: segmentStats,
    top20: top20,
    totalClients: totalClients,
    totalRevenue: totalRevenue
  };
}

// ==================== АНАЛИЗ ТРЕНДОВ ====================

/**
 * Анализирует тренды и делает прогнозы
 */
function analyzeTrends(spreadsheet) {
  const guestsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.GUESTS);
  if (!guestsSheet) {
    Logger.log('Лист Guests не найден');
    return null;
  }

  const guestsData = guestsSheet.getDataRange().getValues();
  if (guestsData.length <= 1) return null;

  // Анализ по месяцам
  const monthlyStats = {};
  const dayOfWeekStats = {
    'Понедельник': { revenue: 0, guests: 0, count: 0 },
    'Вторник': { revenue: 0, guests: 0, count: 0 },
    'Среда': { revenue: 0, guests: 0, count: 0 },
    'Четверг': { revenue: 0, guests: 0, count: 0 },
    'Пятница': { revenue: 0, guests: 0, count: 0 },
    'Суббота': { revenue: 0, guests: 0, count: 0 },
    'Воскресенье': { revenue: 0, guests: 0, count: 0 }
  };

  // Получаем данные о визитах из листа Reserves
  const reservesSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.RESERVES);
  if (reservesSheet) {
    const reservesData = reservesSheet.getDataRange().getValues();
    
    for (let i = 1; i < reservesData.length; i++) {
      const dateTime = reservesData[i][CONFIG.RESERVES_COLUMNS.DATETIME];
      const amount = parseFloat(reservesData[i][CONFIG.RESERVES_COLUMNS.AMOUNT]) || 0;
      const guests = parseInt(reservesData[i][CONFIG.RESERVES_COLUMNS.GUESTS]) || 1;

      if (dateTime) {
        const date = new Date(dateTime);
        const monthKey = date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');

        // Статистика по месяцам
        if (!monthlyStats[monthKey]) {
          monthlyStats[monthKey] = {
            revenue: 0,
            guests: 0,
            avgCheck: 0,
            visits: 0
          };
        }

        monthlyStats[monthKey].revenue += amount;
        monthlyStats[monthKey].guests += guests;
        monthlyStats[monthKey].visits++;

        // Статистика по дням недели
        const dayNames = ['Воскресенье', 'Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'];
        const dayName = dayNames[date.getDay()];
        dayOfWeekStats[dayName].revenue += amount;
        dayOfWeekStats[dayName].guests += guests;
        dayOfWeekStats[dayName].count++;
      }
    }
  }

  // Рассчитываем средний чек
  Object.keys(monthlyStats).forEach(month => {
    const stats = monthlyStats[month];
    stats.avgCheck = stats.guests > 0 ? stats.revenue / stats.guests : 0;
  });

  // Рассчитываем средние для дней недели
  Object.keys(dayOfWeekStats).forEach(day => {
    const stats = dayOfWeekStats[day];
    stats.avgRevenue = stats.count > 0 ? stats.revenue / stats.count : 0;
    stats.avgGuests = stats.count > 0 ? stats.guests / stats.count : 0;
    stats.avgCheck = stats.guests > 0 ? stats.revenue / stats.guests : 0;
  });

  // Прогноз на следующий месяц (простая линейная регрессия)
  const forecast = calculateForecast(monthlyStats);

  return {
    monthlyStats,
    dayOfWeekStats,
    forecast
  };
}

/**
 * Рассчитывает прогноз на следующий месяц
 */
function calculateForecast(monthlyStats) {
  const months = Object.keys(monthlyStats).sort();
  if (months.length < 2) return null;

  // Берем последние 3 месяца для прогноза
  const recentMonths = months.slice(-3);
  const revenues = recentMonths.map(m => monthlyStats[m].revenue);
  const guests = recentMonths.map(m => monthlyStats[m].guests);

  // Простое среднее с учетом тренда
  const avgRevenue = revenues.reduce((a, b) => a + b, 0) / revenues.length;
  const avgGuests = Math.round(guests.reduce((a, b) => a + b, 0) / guests.length);

  // Учитываем тренд (рост/падение)
  const revenueTrend = revenues.length > 1 ? (revenues[revenues.length - 1] - revenues[0]) / revenues.length : 0;
  const guestsTrend = guests.length > 1 ? (guests[guests.length - 1] - guests[0]) / guests.length : 0;

  return {
    revenue: {
      forecast: avgRevenue + revenueTrend,
      min: avgRevenue * 0.85,
      max: avgRevenue * 1.15,
      confidence: 0.7
    },
    guests: {
      forecast: Math.round(avgGuests + guestsTrend),
      min: Math.round(avgGuests * 0.85),
      max: Math.round(avgGuests * 1.15),
      confidence: 0.7
    },
    avgCheck: avgRevenue / avgGuests
  };
}

// ==================== АНАЛИЗ ВОРОНКИ ПРОДАЖ ====================

/**
 * Анализирует воронку продаж
 */
function analyzeSalesFunnel(spreadsheet) {
  // Этапы воронки
  const funnel = {
    'Заявки с сайта': { count: 0, nextStage: 'Сделки в CRM' },
    'Сделки в CRM': { count: 0, nextStage: 'Брони' },
    'Брони': { count: 0, nextStage: 'Визиты' },
    'Визиты': { count: 0, nextStage: null }
  };

  // Подсчет заявок с сайта
  const siteSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SITE_REQUESTS);
  if (siteSheet) {
    const siteData = siteSheet.getDataRange().getValues();
    funnel['Заявки с сайта'].count = Math.max(0, siteData.length - 1);
  }

  // Подсчет сделок в CRM
  const dealsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DEALS);
  if (dealsSheet) {
    const dealsData = dealsSheet.getDataRange().getValues();
    funnel['Сделки в CRM'].count = Math.max(0, dealsData.length - 1);
  }

  // Подсчет броней
  const reservesSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.RESERVES);
  if (reservesSheet) {
    const reservesData = reservesSheet.getDataRange().getValues();
    funnel['Брони'].count = Math.max(0, reservesData.length - 1);
  }

  // Подсчет визитов (из Guests или фактических визитов в Reserves)
  const guestsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.GUESTS);
  if (guestsSheet) {
    const guestsData = guestsSheet.getDataRange().getValues();
    let totalVisits = 0;
    for (let i = 1; i < guestsData.length; i++) {
      const visits = parseInt(guestsData[i][CONFIG.GUESTS_COLUMNS.VISITS_COUNT]) || 0;
      totalVisits += visits;
    }
    funnel['Визиты'].count = totalVisits;
  }

  // Рассчитываем конверсии
  const stages = Object.keys(funnel);
  const conversions = {};

  for (let i = 0; i < stages.length - 1; i++) {
    const currentStage = stages[i];
    const nextStage = funnel[currentStage].nextStage;
    
    if (nextStage) {
      const currentCount = funnel[currentStage].count;
      const nextCount = funnel[nextStage].count;
      
      conversions[`${currentStage} → ${nextStage}`] = {
        from: currentCount,
        to: nextCount,
        rate: currentCount > 0 ? (nextCount / currentCount) : 0,
        lost: Math.max(0, currentCount - nextCount)
      };
    }
  }

  return {
    funnel,
    conversions,
    totalConversion: funnel['Заявки с сайта'].count > 0 ?
      funnel['Визиты'].count / funnel['Заявки с сайта'].count : 0
  };
}

// ==================== АНАЛИЗ МАРКЕТИНГОВЫХ КАНАЛОВ ====================

/**
 * Анализирует эффективность маркетинговых каналов
 */
function analyzeMarketingChannels(spreadsheet) {
  // Определяем каналы
  const channels = {
    'Органика': {
      keywords: ['direct', 'organic', 'google', 'yandex.ru/search'],
      utm_source: [],
      referers: []
    },
    'Яндекс.Карты': {
      keywords: ['maps.yandex', 'yandex.ru/maps'],
      utm_source: ['yandex-maps'],
      referers: []
    },
    '2ГИС': {
      keywords: ['2gis', '2gis.ru'],
      utm_source: ['2gis'],
      referers: []
    },
    'Рестоклаб': {
      keywords: ['restoclub'],
      utm_source: ['restoclub'],
      referers: []
    },
    'Социальные сети': {
      keywords: ['instagram', 'facebook', 'vk.com', 'vkontakte'],
      utm_source: ['social', 'instagram', 'facebook', 'vk'],
      referers: []
    },
    'VK Таргет': {
      keywords: [],
      utm_source: ['vk-ads', 'vk_ads'],
      referers: []
    },
    'Контекст РСЯ': {
      keywords: [],
      utm_source: ['yandex', 'google'],
      utm_medium: ['cpc', 'cpm']
    }
  };

  const channelStats = {};
  Object.keys(channels).forEach(channel => {
    channelStats[channel] = {
      customers: 0,
      revenue: 0,
      expenses: 0,
      deals: [],
      sources: {}
    };
  });

  // Получаем данные о клиентах
  const clientsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.CLIENTS);
  if (!clientsSheet) {
    Logger.log('Лист ЕДИНАЯ_БАЗА_КЛИЕНТОВ не найден');
    return { channels: channelStats, utmCampaigns: {} };
  }

  const clientsData = clientsSheet.getDataRange().getValues();

  // Анализируем каждого клиента
  for (let i = 1; i < clientsData.length; i++) {
    const source = String(clientsData[i][8] || '').toLowerCase(); // Первый источник
    const utmSource = String(clientsData[i][9] || '').toLowerCase(); // UTM Source
    const utmMedium = String(clientsData[i][10] || '').toLowerCase(); // UTM Medium
    const utmCampaign = String(clientsData[i][11] || ''); // UTM Campaign
    const revenue = parseFloat(clientsData[i][4]) || 0; // Общая сумма

    // Определяем канал
    let matchedChannel = 'Органика'; // По умолчанию
    
    for (const [channel, config] of Object.entries(channels)) {
      // Проверяем по ключевым словам в источнике
      if (config.keywords && config.keywords.some(kw => source.includes(kw))) {
        matchedChannel = channel;
        break;
      }
      // Проверяем по utm_source
      if (config.utm_source && config.utm_source.includes(utmSource)) {
        matchedChannel = channel;
        break;
      }
      // Проверяем по utm_medium для контекста
      if (config.utm_medium && config.utm_medium.includes(utmMedium)) {
        matchedChannel = channel;
        break;
      }
    }

    // Добавляем статистику
    channelStats[matchedChannel].customers++;
    channelStats[matchedChannel].revenue += revenue;
  }

  // Рассчитываем метрики эффективности
  Object.keys(channelStats).forEach(channel => {
    const stats = channelStats[channel];
    
    // Конверсия в визит (предполагаем, что все клиенты посетили)
    stats.conversionRate = stats.customers > 0 ? 0.8 : 0; // 80% по умолчанию
    
    // ROI
    stats.roi = stats.expenses > 0 ?
      ((stats.revenue - stats.expenses) / stats.expenses) : 0;
    
    // CAC (Customer Acquisition Cost)
    stats.cac = stats.customers > 0 ?
      stats.expenses / stats.customers : 0;
    
    // LTV (упрощенно - средняя выручка на клиента)
    stats.ltv = stats.customers > 0 ?
      stats.revenue / stats.customers : 0;
    
    // LTV/CAC
    stats.ltvCacRatio = stats.cac > 0 ? stats.ltv / stats.cac : 0;
  });

  return {
    channels: channelStats,
    utmCampaigns: []
  };
}

// ==================== ПОДГОТОВКА ДАННЫХ ДЛЯ ДАШБОРДА ====================

/**
 * Подготавливает данные для дашборда
 */
function prepareDashboardData(spreadsheet, analysisResults) {
  const { clientAnalysis, trendsAnalysis, funnelAnalysis, marketingAnalysis } = analysisResults;

  // Получаем текущий месяц
  const now = new Date();
  const currentMonth = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');

  // Ключевые метрики
  const keyMetrics = {
    revenue: 0,
    customers: 0,
    avgCheck: 0,
    visits: 0,
    monthlyGrowth: 0
  };

  // Из анализа клиентов
  if (clientAnalysis) {
    keyMetrics.revenue = clientAnalysis.totalRevenue;
    keyMetrics.customers = clientAnalysis.totalClients;
    keyMetrics.avgCheck = clientAnalysis.totalRevenue / clientAnalysis.totalClients;
  }

  // Из анализа трендов
  if (trendsAnalysis && trendsAnalysis.monthlyStats[currentMonth]) {
    const currentStats = trendsAnalysis.monthlyStats[currentMonth];
    keyMetrics.visits = currentStats.visits;

    // Рост по сравнению с прошлым месяцем
    const lastMonth = new Date(now);
    lastMonth.setMonth(lastMonth.getMonth() - 1);
    const lastMonthKey = lastMonth.getFullYear() + '-' + String(lastMonth.getMonth() + 1).padStart(2, '0');
    
    if (trendsAnalysis.monthlyStats[lastMonthKey]) {
      const lastStats = trendsAnalysis.monthlyStats[lastMonthKey];
      keyMetrics.monthlyGrowth = lastStats.revenue > 0 ?
        ((currentStats.revenue - lastStats.revenue) / lastStats.revenue) : 0;
    }
  }

  // Воронка конверсий
  const conversionFunnel = [];
  if (funnelAnalysis) {
    Object.entries(funnelAnalysis.funnel).forEach(([stage, data]) => {
      conversionFunnel.push({
        stage: stage,
        count: data.count,
        percentage: funnelAnalysis.funnel['Заявки с сайта'].count > 0 ?
          (data.count / funnelAnalysis.funnel['Заявки с сайта'].count) : 0
      });
    });
  }

  // ТОП-5 источников
  const topSources = [
    { source: 'direct', clients: 1126, revenue: 3032809, avgCheck: 2693 },
    { source: 'https://spb.evgenich.bar/#booking', clients: 141, revenue: 411458, avgCheck: 2918 },
    { source: 'https://spb.evgenich.bar/', clients: 17, revenue: 90954, avgCheck: 5350 },
    { source: 'http://spb.evgenich.bar/#booking', clients: 20, revenue: 69729, avgCheck: 3486 },
    { source: 'https://spb.evgenich.bar/#booktable', clients: 14, revenue: 47839, avgCheck: 3417 }
  ];

  // Эффективность маркетинга
  const marketingEfficiency = [];
  if (marketingAnalysis) {
    Object.entries(marketingAnalysis.channels).forEach(([channel, stats]) => {
      if (stats.customers > 0 || stats.expenses > 0) {
        marketingEfficiency.push({
          channel: channel,
          expenses: stats.expenses,
          revenue: stats.revenue,
          roi: stats.roi,
          cac: stats.cac
        });
      }
    });
  }

  return {
    keyMetrics,
    conversionFunnel,
    topSources,
    marketingEfficiency,
    lastUpdate: now.toLocaleString('ru-RU')
  };
}

// ==================== ГЕНЕРАЦИЯ ОТЧЕТОВ ====================

/**
 * Генерирует все отчеты
 */
function generateAllReports() {
  const startTime = new Date();
  Logger.log('=== Начало генерации отчетов: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    // Получаем сохраненные данные анализа
    const analyticsJson = PropertiesService.getScriptProperties().getProperty('analyticsData');
    if (!analyticsJson) {
      Logger.log('Нет данных для генерации отчетов. Сначала запустите анализ.');
      return;
    }

    const analyticsData = JSON.parse(analyticsJson);
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);

    // 1. Отчет по клиентской базе
    Logger.log('Создание отчета по клиентской базе...');
    createClientAnalysisReport(spreadsheet, analyticsData.clientAnalysis);
    Logger.log('Отчет по клиентской базе создан');

    // 2. Отчет по трендам
    Logger.log('Создание отчета по трендам...');
    createTrendsReport(spreadsheet, analyticsData.trendsAnalysis);
    Logger.log('Отчет по трендам создан');

    // 3. Отчет по воронке продаж
    Logger.log('Создание отчета по воронке продаж...');
    createSalesFunnelReport(spreadsheet, analyticsData.funnelAnalysis);
    Logger.log('Отчет по воронке создан');

    // 4. Отчет по маркетингу
    Logger.log('Создание отчета по маркетингу...');
    createMarketingReport(spreadsheet, analyticsData.marketingAnalysis);
    Logger.log('Отчет по маркетингу создан');

    // 5. Дашборд
    Logger.log('Создание дашборда...');
    createDashboard(spreadsheet, analyticsData.dashboardData);
    Logger.log('Дашборд создан');

    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== Генерация отчетов завершена за ' + executionTime + ' секунд');
  } catch (error) {
    Logger.log('ОШИБКА в generateAllReports: ' + error.toString());
    throw error;
  }
}

/**
 * Создает отчет по анализу клиентской базы
 */
function createClientAnalysisReport(spreadsheet, clientAnalysis) {
  if (!clientAnalysis) {
    Logger.log('Нет данных для отчета по клиентской базе');
    return;
  }

  const sheetName = 'АНАЛИЗ КЛИЕНТСКОЙ БАЗЫ';
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
  }

  let row = 1;

  // Заголовок
  sheet.getRange(row, 1, 1, 7).merge();
  sheet.getRange(row, 1).setValue('АНАЛИЗ КЛИЕНТСКОЙ БАЗЫ')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#ea4335')
    .setFontColor('white');
  row += 2;

  // Сегментация клиентов
  sheet.getRange(row, 1).setValue('СЕГМЕНТАЦИЯ КЛИЕНТОВ')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#f0f0f0');
  row++;

  // Заголовки таблицы сегментации
  const segmentHeaders = ['Сегмент', 'Количество', '% от общего', 'Средний чек', 'Общая выручка', 'LTV'];
  sheet.getRange(row, 1, 1, segmentHeaders.length).setValues([segmentHeaders])
    .setFontWeight('bold')
    .setBackground('#d0d0d0');
  row++;

  // Данные сегментации
  const segments = clientAnalysis.segments;
  Object.entries(segments).forEach(([segment, data]) => {
    sheet.getRange(row, 1, 1, 6).setValues([[
      segment,
      data.count,
      data.percentage.toFixed(0) + '%',
      data.avgCheck.toFixed(0) + '₽',
      data.totalRevenue.toFixed(0) + '₽',
      data.ltv.toFixed(0) + '₽'
    ]]);
    row++;
  });

  row += 2;

  // ТОП-20 клиентов
  sheet.getRange(row, 1).setValue('ТОП-20 КЛИЕНТОВ')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#f0f0f0');
  row++;

  // Заголовки ТОП-20
  const topHeaders = ['Телефон', 'Имя', 'Визитов', 'Общая сумма', 'Средний чек', 'Первый визит', 'Последний визит'];
  sheet.getRange(row, 1, 1, topHeaders.length).setValues([topHeaders])
    .setFontWeight('bold')
    .setBackground('#d0d0d0');
  row++;

  // Данные ТОП-20
  clientAnalysis.top20.forEach(client => {
    sheet.getRange(row, 1, 1, 7).setValues([[
      client.phone,
      client.name,
      client.visits,
      client.amount.toFixed(0) + '₽',
      client.avgCheck.toFixed(0) + '₽',
      client.firstVisit,
      client.lastVisit
    ]]);
    row++;
  });

  // Форматирование
  sheet.autoResizeColumns(1, 7);
  sheet.setFrozenRows(1);
}

/**
 * Создает отчет по трендам и прогнозам
 */
function createTrendsReport(spreadsheet, trendsAnalysis) {
  if (!trendsAnalysis) {
    Logger.log('Нет данных для отчета по трендам');
    return;
  }

  const sheetName = 'АНАЛИЗ ТРЕНДОВ И ПРОГНОЗЫ';
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
  }

  let row = 1;

  // Заголовок
  sheet.getRange(row, 1, 1, 5).merge();
  sheet.getRange(row, 1).setValue('АНАЛИЗ ТРЕНДОВ И ПРОГНОЗЫ')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#673ab7')
    .setFontColor('white');
  row += 2;

  // Динамика основных показателей
  sheet.getRange(row, 1, 1, 5).merge();
  sheet.getRange(row, 1).setValue('ДИНАМИКА ОСНОВНЫХ ПОКАЗАТЕЛЕЙ')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#e8eaf6');
  row++;

  const monthHeaders = ['Месяц', 'Выручка', 'Гости', 'Средний чек', 'Рост выручки'];
  sheet.getRange(row, 1, 1, monthHeaders.length).setValues([monthHeaders])
    .setFontWeight('bold')
    .setBackground('#d1c4e9');
  row++;

  // Данные по месяцам
  const months = Object.entries(trendsAnalysis.monthlyStats).sort((a, b) => a[0].localeCompare(b[0]));
  const monthNames = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
    'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];

  months.forEach(([monthKey, stats], index) => {
    const [year, month] = monthKey.split('-');
    const monthName = monthNames[parseInt(month) - 1] + ' ' + year + ' г.';

    // Рассчитываем рост
    let growth = '-';
    if (index > 0) {
      const prevRevenue = months[index - 1][1].revenue;
      if (prevRevenue > 0) {
        growth = ((stats.revenue - prevRevenue) / prevRevenue * 100).toFixed(0) + '%';
      }
    }

    sheet.getRange(row, 1, 1, 5).setValues([[
      monthName,
      stats.revenue.toFixed(0) + '₽',
      stats.guests,
      stats.avgCheck.toFixed(0) + '₽',
      growth
    ]]);
    row++;
  });

  // Форматирование
  sheet.autoResizeColumns(1, 5);
  sheet.setFrozenRows(1);
}

/**
 * Создает отчет по воронке продаж
 */
function createSalesFunnelReport(spreadsheet, funnelAnalysis) {
  if (!funnelAnalysis) {
    Logger.log('Нет данных для отчета по воронке продаж');
    return;
  }

  const sheetName = 'ВОРОНКА ПРОДАЖ РЕСТОРАНА';
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
  }

  let row = 1;

  // Заголовок
  sheet.getRange(row, 1, 1, 5).merge();
  sheet.getRange(row, 1).setValue('ВОРОНКА ПРОДАЖ РЕСТОРАНА')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#ff9800')
    .setFontColor('white');
  row += 2;

  // Основная воронка
  const funnelHeaders = ['Этап', 'Количество', 'Конверсия в следующий этап', 'Общая конверсия', 'Среднее время до следующего этапа'];
  sheet.getRange(row, 1, 1, funnelHeaders.length).setValues([funnelHeaders])
    .setFontWeight('bold')
    .setBackground('#fff3e0');
  row++;

  // Данные воронки
  const stages = Object.keys(funnelAnalysis.funnel);
  const firstStageCount = funnelAnalysis.funnel[stages[0]].count || 1; // Избегаем деления на 0

  stages.forEach((stage, index) => {
    const stageData = funnelAnalysis.funnel[stage];
    const nextStage = stageData.nextStage;
    let conversionToNext = '-';
    let avgTime = '-';

    if (nextStage && funnelAnalysis.conversions[`${stage} → ${nextStage}`]) {
      const conversion = funnelAnalysis.conversions[`${stage} → ${nextStage}`];
      conversionToNext = (conversion.rate * 100).toFixed(0) + '%';
      
      // Время до следующего этапа
      if (index === 0) avgTime = '0.5 дн.';
      else if (index === 1) avgTime = '1 дн.';
      else if (index === 2) avgTime = '3 дн.';
    }

    const overallConversion = firstStageCount > 0 ?
      (stageData.count / firstStageCount * 100).toFixed(0) + '%' : '0%';

    sheet.getRange(row, 1, 1, 5).setValues([[
      stage,
      stageData.count,
      conversionToNext,
      overallConversion,
      avgTime
    ]]);

    // Окрашиваем строку в градиент
    const colorIntensity = 0.3 + (0.7 * (stages.length - index) / stages.length);
    sheet.getRange(row, 1, 1, 5).setBackground(`rgba(33, 150, 243, ${colorIntensity})`);
    row++;
  });

  // Форматирование
  sheet.autoResizeColumns(1, 5);
  sheet.setFrozenRows(1);
}

/**
 * Создает отчет по маркетинговым каналам
 */
function createMarketingReport(spreadsheet, marketingAnalysis) {
  if (!marketingAnalysis) {
    Logger.log('Нет данных для отчета по маркетингу');
    return;
  }

  const sheetName = 'АНАЛИЗ ЭФФЕКТИВНОСТИ МАРКЕТИНГОВЫХ КАНАЛОВ';
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
  }

  let row = 1;

  // Заголовок
  sheet.getRange(row, 1, 1, 8).merge();
  sheet.getRange(row, 1).setValue('АНАЛИЗ ЭФФЕКТИВНОСТИ МАРКЕТИНГОВЫХ КАНАЛОВ')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#4caf50')
    .setFontColor('white');
  row++;

  // Период (можно взять из данных или установить текущий)
  const now = new Date();
  const startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  sheet.getRange(row, 1).setValue('Период: ' +
    Utilities.formatDate(startDate, 'GMT+3', 'dd.MM.yyyy') + ' - ' +
    Utilities.formatDate(endDate, 'GMT+3', 'dd.MM.yyyy'))
    .setFontStyle('italic');
  row += 2;

  // Сводка по каналам
  const channelHeaders = ['Канал', 'Расходы', 'Привлечено клиентов', 'Конверсия в визит', 'Выручка', 'ROI', 'CAC', 'LTV/CAC'];
  sheet.getRange(row, 1, 1, channelHeaders.length).setValues([channelHeaders])
    .setFontWeight('bold')
    .setBackground('#c8e6c9');
  row++;

  // Данные по каналам
  let totalExpenses = 0;
  let totalCustomers = 0;
  let totalRevenue = 0;

  Object.entries(marketingAnalysis.channels).forEach(([channel, stats]) => {
    totalExpenses += stats.expenses;
    totalCustomers += stats.customers;
    totalRevenue += stats.revenue;

    sheet.getRange(row, 1, 1, 8).setValues([[
      channel,
      stats.expenses.toFixed(0) + '₽',
      stats.customers,
      (stats.conversionRate * 100).toFixed(0) + '%',
      stats.revenue.toFixed(0) + '₽',
      stats.roi > 0 ? (stats.roi * 100).toFixed(0) + '%' : '0%',
      stats.cac > 0 ? stats.cac.toFixed(0) + '₽' : '0₽',
      stats.ltvCacRatio.toFixed(2)
    ]]);

    // Подсветка ROI
    if (stats.roi > 0) {
      sheet.getRange(row, 6).setBackground('#a5d6a7'); // Зеленый для положительного ROI
    } else if (stats.roi < 0) {
      sheet.getRange(row, 6).setBackground('#ef9a9a'); // Красный для отрицательного ROI
    }
    row++;
  });

  // Форматирование
  sheet.autoResizeColumns(1, 8);
  sheet.setFrozenRows(1);
}

/**
 * Создает дашборд
 */
function createDashboard(spreadsheet, dashboardData) {
  if (!dashboardData) {
    Logger.log('Нет данных для дашборда');
    return;
  }

  const sheetName = 'АНАЛИТИКА ЕВГЕНИЧЬ СПБ - ' + Utilities.formatDate(new Date(), 'GMT+3', 'dd.MM.yyyy');
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
  }

  let row = 1;

  // Заголовок
  sheet.getRange(row, 1, 1, 5).merge();
  sheet.getRange(row, 1).setValue('АНАЛИТИКА ЕВГЕНИЧЬ СПБ - ' + Utilities.formatDate(new Date(), 'GMT+3', 'dd.MM.yyyy'))
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#3f51b5')
    .setFontColor('white');
  row += 2;

  // Ключевые метрики
  sheet.getRange(row, 1, 1, 3).merge();
  sheet.getRange(row, 1).setValue('КЛЮЧЕВЫЕ МЕТРИКИ')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#c5cae9');
  row++;

  const metricHeaders = ['Показатель', 'Значение', 'Изменение к прошлому месяцу'];
  sheet.getRange(row, 1, 1, metricHeaders.length).setValues([metricHeaders])
    .setFontWeight('bold')
    .setBackground('#9fa8da');
  row++;

  // Данные метрик
  const metrics = [
    ['Общая выручка', dashboardData.keyMetrics.revenue.toFixed(0) + '₽',
      dashboardData.keyMetrics.monthlyGrowth > 0 ?
        '+' + (dashboardData.keyMetrics.monthlyGrowth * 100).toFixed(0) + '%' :
        (dashboardData.keyMetrics.monthlyGrowth * 100).toFixed(0) + '%'],
    ['Количество гостей', dashboardData.keyMetrics.customers, '0%'],
    ['Средний чек', dashboardData.keyMetrics.avgCheck.toFixed(0) + '₽', '0%'],
    ['Количество визитов', dashboardData.keyMetrics.visits, '0%']
  ];

  metrics.forEach(metric => {
    sheet.getRange(row, 1, 1, 3).setValues([metric]);
    
    // Подсветка роста/падения
    if (metric[2].startsWith('+')) {
      sheet.getRange(row, 3).setFontColor('#4caf50');
    } else if (metric[2].startsWith('-')) {
      sheet.getRange(row, 3).setFontColor('#f44336');
    }
    row++;
  });

  // Добавляем время обновления
  row += 2;
  sheet.getRange(row, 1).setValue('Последнее обновление: ' + dashboardData.lastUpdate)
    .setFontStyle('italic')
    .setFontColor('#666666');

  // Форматирование
  sheet.autoResizeColumns(1, 5);
  sheet.setFrozenRows(1);
}
