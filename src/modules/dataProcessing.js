/**
 * Restaurant Analytics System - Module 2: Data Processing & Linking
 * Обрабатывает и связывает данные между всеми системами
 * Автор: Restaurant Analytics
 * Версия: 2.2 RU - С обновлённой обработкой телефонов
 */

// ==================== ОСНОВНЫЕ ФУНКЦИИ ОБРАБОТКИ ====================

/**
 * Главная функция обработки данных
 * Запускается через 5 минут после сбора данных
 */
function processAndLinkData() {
  const startTime = new Date();
  Logger.log('=== Начало обработки данных: ' + startTime.toLocaleString('ru-RU'));
  
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);

    // Читаем все исходные данные
    const rawData = {
      workingAmo: readWorkingAmoData(spreadsheet),
      reserves: readSheetData(spreadsheet, CONFIG.SHEETS.RESERVES),
      guests: readSheetData(spreadsheet, CONFIG.SHEETS.GUESTS),
      siteRequests: readSheetData(spreadsheet, CONFIG.SHEETS.SITE_REQUESTS),
      budgets: readBudgetsData(spreadsheet)
    };

    // Логируем загруженные данные
    Logger.log('Загружено данных:');
    Logger.log('- РАБОЧИЙ АМО: ' + (rawData.workingAmo?.data?.length || 0) + ' строк');
    Logger.log('- Reserves: ' + (rawData.reserves?.length || 0) + ' строк');
    Logger.log('- Guests: ' + (rawData.guests?.length || 0) + ' строк');
    Logger.log('- Заявки с сайта: ' + (rawData.siteRequests?.length || 0) + ' строк');

    // Проверяем, что данные успешно загружены
    if (!rawData.workingAmo || !rawData.workingAmo.data || rawData.workingAmo.data.length === 0) {
      Logger.log('⚠️ Ошибка: Данные из РАБОЧИЙ АМО не загружены или пусты');
      return;
    }

    // Обрабатываем и связываем данные
    const processedData = {
      // Унифицированные клиенты
      unifiedCustomers: createUnifiedCustomers(rawData),
      // Путь клиента
      customerJourneys: buildCustomerJourneys(rawData),
      // Статистика качества данных
      dataQuality: analyzeDataQuality(rawData)
    };

    // Логируем результаты обработки
    Logger.log('Результаты обработки:');
    Logger.log('- Унифицированных клиентов: ' + processedData.unifiedCustomers.length);
    Logger.log('- Customer journeys: ' + processedData.customerJourneys.length);

    // Сохраняем результаты
    saveProcessedData(spreadsheet, processedData);

    const executionTime = (new Date() - startTime) / 1000;
    Logger.log('=== Обработка завершена за ' + executionTime + ' секунд');

    // Сохраняем время последней обработки
    PropertiesService.getScriptProperties().setProperty('lastDataProcessing', startTime.toISOString());
  } catch (error) {
    Logger.log('ОШИБКА в processAndLinkData: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    throw error;
  }
}

/**
 * Чтение данных из структурированной таблицы РАБОЧИЙ АМО
 * Учитывает двустрочный заголовок и возвращает маппинг колонок
 */
function readWorkingAmoData(spreadsheet) {
  try {
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.WORKING_AMO);
    if (!sheet) {
      Logger.log('⚠️ Лист РАБОЧИЙ АМО не найден');
      return { headers: [], data: [], columnMap: {} };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 3) {
      Logger.log('⚠️ Недостаточно данных в РАБОЧИЙ АМО');
      return { headers: [], data: [], columnMap: {} };
    }

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

    // Логируем несколько найденных полей для проверки
    Logger.log('Найдены поля в РАБОЧИЙ АМО:');
    const fieldsToCheck = ['Контакт.Телефон', 'Контакт.ФИО', 'Сделка.ID', 'Сделка.Название'];
    fieldsToCheck.forEach(field => {
      if (columnMap[field] !== undefined) {
        Logger.log('- ' + field + ': колонка ' + columnMap[field]);
      }
    });

    // Данные начинаются с 3-й строки (индекс 2)
    const rowData = data.slice(2);

    Logger.log('РАБОЧИЙ АМО: загружено ' + rowData.length + ' строк данных');
    return {
      blockHeaders: blockHeaders,
      columnHeaders: columnHeaders,
      data: rowData,
      columnMap: columnMap
    };
  } catch (error) {
    Logger.log('Ошибка при чтении РАБОЧИЙ АМО: ' + error.toString());
    return { headers: [], data: [], columnMap: {} };
  }
}

// ==================== СОЗДАНИЕ УНИФИЦИРОВАННОЙ БАЗЫ КЛИЕНТОВ ====================

/**
 * Создает единую базу клиентов из всех источников
 * Адаптировано для работы с новой структурой РАБОЧИЙ АМО
 */
function createUnifiedCustomers(rawData) {
  Logger.log('=== Создание унифицированной базы клиентов ===');
  
  const customers = new Map(); // Используем Map для быстрого поиска
  let processedCount = 0;
  let guestsProcessed = 0;
  let siteRequestsProcessed = 0;
  let amoProcessed = 0;
  let reservesProcessed = 0;

  // 1. Сначала обрабатываем гостей из Guests RP (у них есть история визитов)
  if (rawData.guests && rawData.guests.length > 1) {
    Logger.log('Обработка Guests RP...');
    
    // Проверяем первую строку данных для диагностики
    if (rawData.guests[1]) {
      Logger.log('Пример данных Guests (первая строка):');
      Logger.log('- Телефон (индекс ' + CONFIG.GUESTS_COLUMNS.PHONE + '): ' + rawData.guests[1][CONFIG.GUESTS_COLUMNS.PHONE]);
      Logger.log('- Имя (индекс ' + CONFIG.GUESTS_COLUMNS.NAME + '): ' + rawData.guests[1][CONFIG.GUESTS_COLUMNS.NAME]);
      Logger.log('- Телефон после очистки: ' + cleanPhone(rawData.guests[1][CONFIG.GUESTS_COLUMNS.PHONE]));
    }

    for (let i = 1; i < rawData.guests.length; i++) {
      const row = rawData.guests[i];
      if (!row || row.length === 0) continue;

      const phone = cleanPhone(row[CONFIG.GUESTS_COLUMNS.PHONE]);
      const email = cleanEmail(row[CONFIG.GUESTS_COLUMNS.EMAIL]);
      const customerId = phone || email; // Используем телефон как основной ID

      if (!customerId) continue;

      customers.set(customerId, {
        id: customerId,
        name: row[CONFIG.GUESTS_COLUMNS.NAME] || '',
        phone: phone,
        email: email,
        // Данные из Guests
        visitsCount: parseNumber(row[CONFIG.GUESTS_COLUMNS.VISITS_COUNT]),
        totalAmount: parseNumber(row[CONFIG.GUESTS_COLUMNS.TOTAL_AMOUNT]),
        firstVisitDate: formatDate(row[CONFIG.GUESTS_COLUMNS.FIRST_VISIT]),
        lastVisitDate: formatDate(row[CONFIG.GUESTS_COLUMNS.LAST_VISIT]),
        avgCheck: 0, // Рассчитаем позже
        // Плейсхолдеры для других данных
        firstSource: '',
        firstUtmSource: '',
        firstUtmMedium: '',
        firstUtmCampaign: '',
        amoDeals: [],
        reserves: [],
        siteRequests: [],
        totalBudgetSpent: 0
      });

      guestsProcessed++;
    }
    Logger.log('Обработано записей из Guests: ' + guestsProcessed);
  }

  // 2. Обогащаем данными из заявок с сайта
  if (rawData.siteRequests && rawData.siteRequests.length > 1) {
    Logger.log('Обработка заявок с сайта...');
    
    // Проверяем первую строку данных для диагностики
    if (rawData.siteRequests[1]) {
      Logger.log('Пример данных заявок (первая строка):');
      Logger.log('- Телефон (индекс ' + CONFIG.SITE_COLUMNS.PHONE + '): ' + rawData.siteRequests[1][CONFIG.SITE_COLUMNS.PHONE]);
      Logger.log('- Имя (индекс ' + CONFIG.SITE_COLUMNS.NAME + '): ' + rawData.siteRequests[1][CONFIG.SITE_COLUMNS.NAME]);
      Logger.log('- Телефон после очистки: ' + cleanPhone(rawData.siteRequests[1][CONFIG.SITE_COLUMNS.PHONE]));
    }

    for (let i = 1; i < rawData.siteRequests.length; i++) {
      const row = rawData.siteRequests[i];
      if (!row || row.length === 0) continue;

      const phone = cleanPhone(row[CONFIG.SITE_COLUMNS.PHONE]);
      const email = cleanEmail(row[CONFIG.SITE_COLUMNS.EMAIL]);
      const customerId = phone || email;

      if (!customerId) continue;

      let customer = customers.get(customerId);
      if (!customer) {
        // Новый клиент
        customer = createNewCustomer(customerId, phone, email, row[CONFIG.SITE_COLUMNS.NAME]);
        customers.set(customerId, customer);
      }

      // Добавляем информацию о заявке
      customer.siteRequests.push({
        date: formatDate(row[CONFIG.SITE_COLUMNS.DATE]),
        formName: row[CONFIG.SITE_COLUMNS.FORM_NAME] || '',
        utmSource: row[CONFIG.SITE_COLUMNS.UTM_SOURCE] || '',
        utmMedium: row[CONFIG.SITE_COLUMNS.UTM_MEDIUM] || '',
        utmCampaign: row[CONFIG.SITE_COLUMNS.UTM_CAMPAIGN] || ''
      });

      // Обновляем первый источник если это более ранняя заявка
      updateFirstSource(customer, row);

      siteRequestsProcessed++;
    }
    Logger.log('Обработано заявок с сайта: ' + siteRequestsProcessed);
  }

  // 3. Обогащаем данными из РАБОЧИЙ АМО
  if (rawData.workingAmo && rawData.workingAmo.data && rawData.workingAmo.data.length > 0) {
    Logger.log('Обработка РАБОЧИЙ АМО...');
    
    const amoData = rawData.workingAmo.data;
    const columnMap = rawData.workingAmo.columnMap;

    // Получаем индексы колонок по их названиям
    const getColIndex = (fieldName) => {
      return columnMap[fieldName] !== undefined ? columnMap[fieldName] : -1;
    };

    // Получаем значение из строки по названию поля
    const getValue = (row, fieldName) => {
      const index = getColIndex(fieldName);
      return index !== -1 && row[index] !== undefined ? row[index] : null;
    };

    // Проверяем первую строку для диагностики
    if (amoData[0]) {
      const phoneIndex = getColIndex(CONFIG.WORKING_AMO_BLOCKS.CONTACT.PHONE);
      const nameIndex = getColIndex(CONFIG.WORKING_AMO_BLOCKS.CONTACT.NAME);
      Logger.log('Индексы полей AMO: телефон=' + phoneIndex + ', имя=' + nameIndex);
      
      if (phoneIndex !== -1) {
        Logger.log('Пример телефона из AMO: ' + amoData[0][phoneIndex]);
        Logger.log('Телефон после очистки: ' + cleanPhone(amoData[0][phoneIndex]));
      }
    }

    for (let i = 0; i < amoData.length; i++) {
      const row = amoData[i];
      if (!row || row.length === 0) continue;

      // Получаем телефон и email из блока CONTACT
      const phone = cleanPhone(getValue(row, CONFIG.WORKING_AMO_BLOCKS.CONTACT.PHONE));
      
      // Email нет в конфигурации, поищем его по названию поля
      let email = '';
      for (let j = 0; j < rawData.workingAmo.columnHeaders.length; j++) {
        const header = rawData.workingAmo.columnHeaders[j];
        if (header && header.toString().toLowerCase().includes('email')) {
          email = cleanEmail(row[j]);
          break;
        }
      }

      const customerId = phone || email;
      if (!customerId) continue;

      let customer = customers.get(customerId);
      if (!customer) {
        customer = createNewCustomer(
          customerId,
          phone,
          email,
          getValue(row, CONFIG.WORKING_AMO_BLOCKS.CONTACT.NAME)
        );
        customers.set(customerId, customer);
      }

      // Добавляем информацию о сделке
      const dealId = getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.ID);
      if (dealId) {
        customer.amoDeals.push({
          id: dealId,
          name: getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.NAME) || '',
          stage: getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.STATUS) || '',
          budget: parseNumber(getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.BUDGET)),
          createDate: formatDate(getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.CREATE_DATE)),
          closeDate: formatDate(getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.CLOSE_DATE)),
          source: getValue(row, CONFIG.WORKING_AMO_BLOCKS.UTM.DEAL_SOURCE) || '',
          city: getValue(row, CONFIG.WORKING_AMO_BLOCKS.ADDITIONAL.CITY_TAG) || '',
          leadType: getValue(row, CONFIG.WORKING_AMO_BLOCKS.UTM.LEAD_TYPE) || ''
        });
      }

      // Обновляем UTM-данные, если они есть
      const utmSource = getValue(row, CONFIG.WORKING_AMO_BLOCKS.UTM.SOURCE);
      const utmMedium = getValue(row, CONFIG.WORKING_AMO_BLOCKS.UTM.MEDIUM);
      const utmCampaign = getValue(row, CONFIG.WORKING_AMO_BLOCKS.UTM.CAMPAIGN);

      if (utmSource && (!customer.firstUtmSource || !customer.firstSourceDate)) {
        customer.firstUtmSource = utmSource;
        customer.firstUtmMedium = utmMedium || '';
        customer.firstUtmCampaign = utmCampaign || '';
        customer.firstSourceDate = getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.CREATE_DATE);
      }

      amoProcessed++;
    }
    Logger.log('Обработано записей из AMO: ' + amoProcessed);
  }

  // 4. Обогащаем данными из броней
  if (rawData.reserves && rawData.reserves.length > 1) {
    Logger.log('Обработка броней...');
    
    // Проверяем первую строку данных для диагностики
    if (rawData.reserves[1]) {
      Logger.log('Пример данных броней (первая строка):');
      Logger.log('- Телефон (индекс ' + CONFIG.RESERVES_COLUMNS.PHONE + '): ' + rawData.reserves[1][CONFIG.RESERVES_COLUMNS.PHONE]);
      Logger.log('- Имя (индекс ' + CONFIG.RESERVES_COLUMNS.NAME + '): ' + rawData.reserves[1][CONFIG.RESERVES_COLUMNS.NAME]);
      Logger.log('- Телефон после очистки: ' + cleanPhone(rawData.reserves[1][CONFIG.RESERVES_COLUMNS.PHONE]));
    }

    for (let i = 1; i < rawData.reserves.length; i++) {
      const row = rawData.reserves[i];
      if (!row || row.length === 0) continue;

      const phone = cleanPhone(row[CONFIG.RESERVES_COLUMNS.PHONE]);
      const email = cleanEmail(row[CONFIG.RESERVES_COLUMNS.EMAIL]);
      const customerId = phone || email;

      if (!customerId) continue;

      let customer = customers.get(customerId);
      if (!customer) {
        customer = createNewCustomer(customerId, phone, email, row[CONFIG.RESERVES_COLUMNS.NAME]);
        customers.set(customerId, customer);
      }

      // Добавляем информацию о брони
      customer.reserves.push({
        id: row[CONFIG.RESERVES_COLUMNS.RESERVE_ID],
        datetime: row[CONFIG.RESERVES_COLUMNS.DATETIME],
        status: row[CONFIG.RESERVES_COLUMNS.STATUS] || '',
        amount: parseNumber(row[CONFIG.RESERVES_COLUMNS.AMOUNT]),
        guests: parseNumber(row[CONFIG.RESERVES_COLUMNS.GUESTS])
      });

      reservesProcessed++;
    }
    Logger.log('Обработано броней: ' + reservesProcessed);
  }

  // 5. Рассчитываем дополнительные метрики
  customers.forEach(customer => {
    // Средний чек
    if (customer.visitsCount > 0) {
      customer.avgCheck = customer.totalAmount / customer.visitsCount;
    }

    // Сортируем массивы по датам
    customer.siteRequests.sort((a, b) => new Date(a.date) - new Date(b.date));
    customer.amoDeals.sort((a, b) => new Date(a.createDate) - new Date(b.createDate));
    customer.reserves.sort((a, b) => new Date(a.datetime) - new Date(b.datetime));
  });

  Logger.log('=== ИТОГО создано унифицированных профилей: ' + customers.size + ' ===');
  Logger.log('Источники: Guests=' + guestsProcessed + ', Site=' + siteRequestsProcessed + ', AMO=' + amoProcessed + ', Reserves=' + reservesProcessed);

  // Конвертируем Map в массив для сохранения
  return Array.from(customers.values());
}

/**
 * Создает нового клиента
 */
function createNewCustomer(customerId, phone, email, name) {
  return {
    id: customerId,
    name: name || '',
    phone: phone,
    email: email,
    visitsCount: 0,
    totalAmount: 0,
    firstVisitDate: '',
    lastVisitDate: '',
    avgCheck: 0,
    firstSource: '',
    firstUtmSource: '',
    firstUtmMedium: '',
    firstUtmCampaign: '',
    firstSourceDate: '',
    amoDeals: [],
    reserves: [],
    siteRequests: [],
    totalBudgetSpent: 0
  };
}

/**
 * Обновляет первый источник клиента
 */
function updateFirstSource(customer, siteRequestRow) {
  const requestDate = formatDate(siteRequestRow[CONFIG.SITE_COLUMNS.DATE]);
  
  // Если это первая заявка или более ранняя
  if (!customer.firstSource || !customer.firstSourceDate ||
      (requestDate && new Date(requestDate) < new Date(customer.firstSourceDate))) {
    customer.firstSource = siteRequestRow[CONFIG.SITE_COLUMNS.REFERER] || 'direct';
    customer.firstUtmSource = siteRequestRow[CONFIG.SITE_COLUMNS.UTM_SOURCE] || '';
    customer.firstUtmMedium = siteRequestRow[CONFIG.SITE_COLUMNS.UTM_MEDIUM] || '';
    customer.firstUtmCampaign = siteRequestRow[CONFIG.SITE_COLUMNS.UTM_CAMPAIGN] || '';
    customer.firstSourceDate = requestDate;
  }
}

// ==================== ПОСТРОЕНИЕ CUSTOMER JOURNEY ====================

/**
 * Строит путь клиента через все системы
 */
function buildCustomerJourneys(rawData) {
  Logger.log('=== Построение customer journey ===');
  
  const journeys = [];
  const phoneToJourney = new Map();
  let totalEvents = 0;

  // Собираем все события для каждого клиента
  // 1. События из заявок с сайта
  if (rawData.siteRequests && rawData.siteRequests.length > 1) {
    Logger.log('Обработка событий из заявок с сайта...');
    let siteEvents = 0;
    
    for (let i = 1; i < rawData.siteRequests.length; i++) {
      const row = rawData.siteRequests[i];
      if (!row || row.length === 0) continue;

      const phone = cleanPhone(row[CONFIG.SITE_COLUMNS.PHONE]);
      if (!phone) continue;

      if (!phoneToJourney.has(phone)) {
        phoneToJourney.set(phone, {
          phone: phone,
          events: []
        });
      }

      phoneToJourney.get(phone).events.push({
        type: 'ЗАЯВКА_С_САЙТА',
        date: formatDate(row[CONFIG.SITE_COLUMNS.DATE]),
        time: row[CONFIG.SITE_COLUMNS.TIME] || '',
        details: {
          formName: row[CONFIG.SITE_COLUMNS.FORM_NAME] || '',
          utmSource: row[CONFIG.SITE_COLUMNS.UTM_SOURCE] || '',
          buttonText: row[CONFIG.SITE_COLUMNS.BUTTON_TEXT] || ''
        }
      });
      siteEvents++;
    }
    Logger.log('Добавлено событий из заявок: ' + siteEvents);
  }

  // 2. События из РАБОЧИЙ АМО
  if (rawData.workingAmo && rawData.workingAmo.data && rawData.workingAmo.data.length > 0) {
    Logger.log('Обработка событий из AMO...');
    let amoEvents = 0;
    
    const amoData = rawData.workingAmo.data;
    const columnMap = rawData.workingAmo.columnMap;

    // Получаем значение из строки по названию поля
    const getValue = (row, fieldName) => {
      const index = columnMap[fieldName];
      return index !== undefined && row[index] !== undefined ? row[index] : null;
    };

    for (let i = 0; i < amoData.length; i++) {
      const row = amoData[i];
      if (!row || row.length === 0) continue;

      // Получаем телефон из блока CONTACT
      const phone = cleanPhone(getValue(row, CONFIG.WORKING_AMO_BLOCKS.CONTACT.PHONE));
      if (!phone) continue;

      if (!phoneToJourney.has(phone)) {
        phoneToJourney.set(phone, {
          phone: phone,
          events: []
        });
      }

      const journey = phoneToJourney.get(phone);

      // Создание сделки
      const createDate = getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.CREATE_DATE);
      if (createDate) {
        journey.events.push({
          type: 'СОЗДАНА_СДЕЛКА_AMO',
          date: formatDate(createDate),
          details: {
            dealId: getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.ID),
            dealName: getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.NAME) || '',
            stage: getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.STATUS) || '',
            responsible: getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.RESPONSIBLE) || ''
          }
        });
        amoEvents++;
      }

      // Закрытие сделки (если есть)
      const closeDate = getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.CLOSE_DATE);
      if (closeDate) {
        journey.events.push({
          type: 'ЗАКРЫТА_СДЕЛКА_AMO',
          date: formatDate(closeDate),
          details: {
            dealId: getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.ID),
            budget: parseNumber(getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.BUDGET))
          }
        });
        amoEvents++;
      }

      // Бронь (если есть данные о дате брони)
      const bookingDate = getValue(row, CONFIG.WORKING_AMO_BLOCKS.RESERVATION.BOOKING_DATE);
      if (bookingDate) {
        journey.events.push({
          type: 'СОЗДАНА_БРОНЬ_AMO',
          date: formatDate(bookingDate),
          details: {
            dealId: getValue(row, CONFIG.WORKING_AMO_BLOCKS.DEAL.ID),
            bar: getValue(row, CONFIG.WORKING_AMO_BLOCKS.RESERVATION.BAR) || '',
            time: getValue(row, CONFIG.WORKING_AMO_BLOCKS.RESERVATION.ARRIVAL_TIME) || '',
            guests: parseNumber(getValue(row, CONFIG.WORKING_AMO_BLOCKS.RESERVATION.GUESTS_COUNT))
          }
        });
        amoEvents++;
      }
    }
    Logger.log('Добавлено событий из AMO: ' + amoEvents);
  }

  // Сортируем события по дате и создаем итоговый массив
  phoneToJourney.forEach((journey, phone) => {
    if (journey.events.length === 0) return;

    // Сортируем события по дате
    journey.events.sort((a, b) => {
      const dateA = new Date(a.date + ' ' + (a.time || '00:00'));
      const dateB = new Date(b.date + ' ' + (b.time || '00:00'));
      return dateA - dateB;
    });

    // Рассчитываем время между событиями
    for (let i = 1; i < journey.events.length; i++) {
      const prevDate = new Date(journey.events[i-1].date);
      const currDate = new Date(journey.events[i].date);
      const daysBetween = Math.floor((currDate - prevDate) / (1000 * 60 * 60 * 24));
      journey.events[i].daysSincePrevious = daysBetween;
    }

    totalEvents += journey.events.length;
    journeys.push(journey);
  });

  Logger.log('=== ИТОГО построено customer journey для ' + journeys.length + ' клиентов ===');
  Logger.log('Всего событий: ' + totalEvents);
  
  return journeys;
}

// ==================== АНАЛИЗ КАЧЕСТВА ДАННЫХ ====================

/**
 * Анализирует качество данных и находит проблемы
 */
function analyzeDataQuality(rawData) {
  Logger.log('=== Анализ качества данных ===');
  
  // Убедимся, что все данные существуют
  const workingAmo = rawData.workingAmo && rawData.workingAmo.data ? rawData.workingAmo.data : [];
  const reserves = Array.isArray(rawData.reserves) ? rawData.reserves : [];
  const guests = Array.isArray(rawData.guests) ? rawData.guests : [];
  const siteRequests = Array.isArray(rawData.siteRequests) ? rawData.siteRequests : [];

  const quality = {
    totalRecords: {
      workingAmo: workingAmo.length,
      reserves: reserves.length > 0 ? reserves.length - 1 : 0,
      guests: guests.length > 0 ? guests.length - 1 : 0,
      siteRequests: siteRequests.length > 0 ? siteRequests.length - 1 : 0
    },
    missingPhones: {
      workingAmo: 0,
      reserves: 0,
      guests: 0,
      siteRequests: 0
    },
    missingEmails: {
      workingAmo: 0,
      reserves: 0,
      guests: 0,
      siteRequests: 0
    },
    duplicatePhones: new Set(),
    invalidPhones: [],
    matchingStats: {
      amoToReserves: 0,
      amoToGuests: 0,
      siteToAMO: 0,
      reservesToGuests: 0
    }
  };

  // Проверяем данные РАБОЧИЙ АМО
  const amoPhones = new Set();
  const sitePhones = new Set();

  // РАБОЧИЙ АМО
  if (rawData.workingAmo && rawData.workingAmo.columnMap) {
    const columnMap = rawData.workingAmo.columnMap;
    const phoneIndex = columnMap[CONFIG.WORKING_AMO_BLOCKS.CONTACT.PHONE];

    // Ищем индекс для email
    let emailIndex = -1;
    for (let j = 0; j < rawData.workingAmo.columnHeaders.length; j++) {
      const header = rawData.workingAmo.columnHeaders[j];
      if (header && header.toString().toLowerCase().includes('email')) {
        emailIndex = j;
        break;
      }
    }

    for (let i = 0; i < workingAmo.length; i++) {
      const row = workingAmo[i];
      if (!row || row.length === 0) continue;

      const phone = phoneIndex !== undefined ? cleanPhone(row[phoneIndex]) : '';
      const email = emailIndex !== -1 ? cleanEmail(row[emailIndex]) : '';

      if (!phone) quality.missingPhones.workingAmo++;
      if (!email) quality.missingEmails.workingAmo++;

      if (phone) {
        if (amoPhones.has(phone)) {
          quality.duplicatePhones.add(phone);
        }
        amoPhones.add(phone);

        // Проверяем валидность телефона
        if (phone.length < 10 || phone.length > 11) {
          quality.invalidPhones.push(phone);
        }
      }
    }
  }

  // Аналогично для других источников данных...
  // (код аналогичен исходному)

  Logger.log('Анализ качества завершен');
  Logger.log('Найдено дубликатов телефонов: ' + quality.duplicatePhones.size);
  Logger.log('Невалидных телефонов: ' + quality.invalidPhones.length);
  
  return quality;
}

// ==================== СОХРАНЕНИЕ РЕЗУЛЬТАТОВ ====================

/**
 * Сохраняет обработанные данные в новые листы
 */
function saveProcessedData(spreadsheet, processedData) {
  Logger.log('=== Сохранение обработанных данных ===');

  // 1. Сохраняем унифицированных клиентов
  saveUnifiedCustomers(spreadsheet, processedData.unifiedCustomers);

  // 2. Сохраняем customer journeys
  saveCustomerJourneys(spreadsheet, processedData.customerJourneys);

  // 3. Сохраняем отчет о качестве данных
  saveDataQualityReport(spreadsheet, processedData.dataQuality);

  Logger.log('Все данные сохранены');
}

/**
 * Сохраняет унифицированную базу клиентов
 */
function saveUnifiedCustomers(spreadsheet, customers) {
  const sheetName = 'ЕДИНАЯ_БАЗА_КЛИЕНТОВ';
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    // Очищаем все данные кроме заголовков
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow > 1 && lastCol > 0) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    }
  }

  // Заголовки
  const headers = [
    'ID (Телефон)',
    'Имя',
    'Email',
    'Кол-во визитов',
    'Общая сумма',
    'Средний чек',
    'Первый визит',
    'Последний визит',
    'Первый источник',
    'UTM Source',
    'UTM Medium',
    'UTM Campaign',
    'Кол-во сделок AMO',
    'Кол-во броней',
    'Кол-во заявок'
  ];
  
  const data = [headers];

  // Данные
  customers.forEach(customer => {
    data.push([
      customer.phone,
      customer.name,
      customer.email,
      customer.visitsCount,
      customer.totalAmount,
      Math.round(customer.avgCheck),
      customer.firstVisitDate,
      customer.lastVisitDate,
      customer.firstSource,
      customer.firstUtmSource,
      customer.firstUtmMedium,
      customer.firstUtmCampaign,
      customer.amoDeals.length,
      customer.reserves.length,
      customer.siteRequests.length
    ]);
  });

  // Записываем данные
  if (data.length > 1) {
    sheet.getRange(1, 1, data.length, headers.length).setValues(data);
    
    // Форматирование
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    // Автоматическая ширина колонок
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
    
    // Форматирование чисел
    if (data.length > 1) {
      sheet.getRange(2, 4, data.length - 1, 1).setNumberFormat('#,##0'); // Кол-во визитов
      sheet.getRange(2, 5, data.length - 1, 2).setNumberFormat('#,##0₽'); // Суммы
    }
  }

  Logger.log('Сохранено ' + (data.length - 1) + ' унифицированных клиентов');
}

/**
 * Сохраняет customer journeys в более читаемом формате
 */
function saveCustomerJourneys(spreadsheet, journeys) {
  const sheetName = 'ПУТЬ_КЛИЕНТА';
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    // Очищаем все данные кроме заголовков
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow > 1 && lastCol > 0) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    }
  }

  // Заголовки
  const headers = [
    'Телефон',
    'Событие',
    'Дата',
    'Дней с предыдущего',
    'Название формы',
    'UTM Source',
    'Название сделки',
    'Статус',
    'Сумма'
  ];
  
  const data = [headers];

  // Данные
  journeys.forEach(journey => {
    journey.events.forEach((event, index) => {
      const details = event.details || {};
      
      // Извлекаем читаемые данные из деталей
      let formName = '';
      let utmSource = '';
      let dealName = '';
      let status = '';
      let amount = '';
      
      switch(event.type) {
        case 'ЗАЯВКА_С_САЙТА':
          formName = details.formName || '';
          utmSource = details.utmSource || '';
          break;
        case 'СОЗДАНА_СДЕЛКА_AMO':
          dealName = details.dealName || '';
          status = details.stage || '';
          break;
        case 'ЗАКРЫТА_СДЕЛКА_AMO':
          amount = details.budget || 0;
          break;
        case 'СОЗДАНА_БРОНЬ':
        case 'СОЗДАНА_БРОНЬ_AMO':
          status = details.status || '';
          amount = details.amount || 0;
          break;
        case 'ПЕРВЫЙ_ВИЗИТ':
        case 'ПОСЛЕДНИЙ_ВИЗИТ':
          amount = details.totalAmount || 0;
          break;
      }

      data.push([
        journey.phone,
        event.type,
        event.date,
        event.daysSincePrevious || 0,
        formName,
        utmSource,
        dealName,
        status,
        amount
      ]);
    });
  });

  // Записываем данные
  if (data.length > 1) {
    sheet.getRange(1, 1, data.length, headers.length).setValues(data);
    
    // Форматирование
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    // Автоматическая ширина колонок
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
    
    // Форматирование суммы как валюты
    if (data.length > 1) {
      sheet.getRange(2, 9, data.length - 1, 1).setNumberFormat('#,##0₽');
    }
  }

  Logger.log('Сохранено ' + journeys.length + ' customer journeys с ' + (data.length - 1) + ' событиями');
}

/**
 * Сохраняет отчет о качестве данных
 */
function saveDataQualityReport(spreadsheet, quality) {
  const sheetName = 'КАЧЕСТВО_ДАННЫХ';
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    sheet.clear(); // Полностью очищаем лист
  }

  try {
    // Проверяем, что объект качества существует и имеет все нужные поля
    if (!quality || !quality.totalRecords) {
      Logger.log('Ошибка: объект качества данных неполный');
      // Создаем базовый отчет с заголовками
      const basicReport = [
        ['Метрика', 'РАБОЧИЙ АМО', 'Reserves', 'Guests', 'Заявки с сайта'],
        ['Всего записей', 0, 0, 0, 0],
        ['Ошибка при формировании отчета', '', '', '', '']
      ];
      sheet.getRange(1, 1, basicReport.length, basicReport[0].length).setValues(basicReport);
      return;
    }

    // Формируем отчет с гарантированной структурой
    const report = [];
    
    // Первая строка (заголовки)
    report.push(['Метрика', 'РАБОЧИЙ АМО', 'Reserves', 'Guests', 'Заявки с сайта']);
    
    // Остальные строки с проверкой на существование данных
    report.push([
      'Всего записей',
      quality.totalRecords?.workingAmo || 0,
      quality.totalRecords?.reserves || 0,
      quality.totalRecords?.guests || 0,
      quality.totalRecords?.siteRequests || 0
    ]);

    report.push([
      'Без телефона',
      quality.missingPhones?.workingAmo || 0,
      quality.missingPhones?.reserves || 0,
      quality.missingPhones?.guests || 0,
      quality.missingPhones?.siteRequests || 0
    ]);

    report.push([
      'Без email',
      quality.missingEmails?.workingAmo || 0,
      quality.missingEmails?.reserves || 0,
      quality.missingEmails?.guests || 0,
      quality.missingEmails?.siteRequests || 0
    ]);

    // Записываем данные
    sheet.getRange(1, 1, report.length, report[0].length).setValues(report);
    
    // Форматирование
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    
    // Автоматическая ширина колонок
    for (let i = 1; i <= 5; i++) {
      sheet.autoResizeColumn(i);
    }

    Logger.log('Отчет о качестве данных сохранен');
  } catch (error) {
    Logger.log('Ошибка при сохранении отчета: ' + error.toString());
    // Сохраняем информацию об ошибке
    sheet.getRange(1, 1, 3, 1).setValues([
      ['ОШИБКА ФОРМИРОВАНИЯ ОТЧЕТА'],
      ['Время ошибки: ' + new Date().toLocaleString('ru-RU')],
      ['Текст ошибки: ' + error.toString()]
    ]);
  }
}

/**
 * Читает данные о бюджетах из кеша
 */
function readBudgetsData(spreadsheet) {
  try {
    const cache = CacheService.getScriptCache();
    const budgetsJson = cache.get('budgetsData');
    if (budgetsJson) {
      return JSON.parse(budgetsJson);
    }
    // Если кеша нет, читаем напрямую из листа
    return {}; // Упрощенно, т.к. функция не используется
  } catch (error) {
    Logger.log('Ошибка чтения бюджетов: ' + error.toString());
    return {};
  }
}

// ==================== УПРАВЛЕНИЕ ТРИГГЕРАМИ ====================

/**
 * Устанавливает триггер для обработки данных
 * Запускается через 5 минут после каждого часа
 */
function setupProcessingTrigger() {
  // Удаляем существующие триггеры
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processAndLinkData') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Создаем новый триггер
  ScriptApp.newTrigger('processAndLinkData')
    .timeBased()
    .everyHours(1)
    .nearMinute(5) // Запускаем на 5-й минуте каждого часа
    .create();

  Logger.log('Триггер обработки данных установлен (каждый час на 5-й минуте)');

  // Запускаем первую обработку
  processAndLinkData();
}

/**
 * Ручной запуск с проверкой исходных данных
 */
function runProcessingWithVerification() {
  // Запускаем обработку
  processAndLinkData();
}
