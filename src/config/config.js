/**
 * Restaurant Analytics System - Центральная конфигурация
 * Все настройки для всех модулей системы
 * Автор: Restaurant Analytics
 * Версия: 3.2
 */

const CONFIG = {
  // ID основной таблицы
  MAIN_SPREADSHEET_ID: '1tD89CZMI8KqaHBx0gmGsHpc9eKYvpuk3OnCOpDYMDdE',
  
  // Названия листов для данных
  SHEETS: {
    // Исходные данные
    DEALS: 'РАБОЧИЙ АМО',
    WORKING_AMO: 'РАБОЧИЙ АМО', // Дублируем для совместимости
    BOOKINGS: 'Reserves RP',
    RESERVES: 'Reserves RP', // Дублируем для совместимости
    GUESTS: 'Guests RP',
    CLIENTS: 'ЕДИНАЯ_БАЗА_КЛИЕНТОВ',
    CUSTOMER_JOURNEY: 'ПУТЬ_КЛИЕНТА',
    DATA_QUALITY: 'КАЧЕСТВО_ДАННЫХ',
    SITE_REQUESTS: 'Заявки с Сайта',
    BUDGETS: 'Бюджеты',
    
    // Служебные листы для кеширования
    CACHE_ALL_DATA: '_CACHE_ALL_DATA',
    CACHE_METADATA: '_CACHE_METADATA',
    
    // Новые отчеты
    CLIENT_ANALYSIS: 'АНАЛИЗ КЛИЕНТСКОЙ БАЗЫ',
    TRENDS_ANALYSIS: 'АНАЛИЗ ТРЕНДОВ И ПРОГНОЗЫ',
    SALES_FUNNEL: 'ВОРОНКА ПРОДАЖ РЕСТОРАНА',
    MARKETING_ANALYSIS: 'АНАЛИЗ ЭФФЕКТИВНОСТИ МАРКЕТИНГОВЫХ КАНАЛОВ',
    DASHBOARD_PREFIX: 'АНАЛИТИКА ЕВГЕНИЧЬ СПБ' // Префикс для дашборда с датой
  },

  // Маппинг колонок в таблице РАБОЧИЙ АМО по блокам
  WORKING_AMO_BLOCKS: {
    DEAL: {
      ID: 'Сделка.ID',
      NAME: 'Сделка.Название',
      RESPONSIBLE: 'Сделка.Ответственный',
      STATUS: 'Сделка.Статус',
      BUDGET: 'Сделка.Бюджет',
      CREATE_DATE: 'Сделка.Дата создания',
      CLOSE_DATE: 'Сделка.Дата закрытия',
      CREATOR: 'Кем создана',
      TAGS: 'Сделка.Теги',
      STATUS_HISTORY: 'История статусов'
    },
    CONTACT: {
      NAME: 'Контакт.ФИО',
      PHONE: 'Контакт.Телефон',
      MANGO_LINE: 'Контакт.Номер линии MANGO OFFICE',
      DEAL_MANGO_LINE: 'Сделка.Номер линии MANGO OFFICE'
    },
    RESERVATION: {
      BAR: 'Сделка.Бар (deal)',
      BOOKING_DATE: 'Сделка.Дата брони',
      ARRIVAL_TIME: 'Сделка.Время прихода',
      GUESTS_COUNT: 'Сделка.Кол-во гостей',
      COMMENT: 'Сделка.Комментарий МОБ',
      GUEST_STATUS: 'Сделка.R.Статусы гостей'
    },
    UTM: {
      SOURCE: 'Сделка.UTM_SOURCE',
      MEDIUM: 'Сделка.UTM_MEDIUM',
      CAMPAIGN: 'Сделка.UTM_CAMPAIGN',
      CONTENT: 'Сделка.UTM_CONTENT',
      TERM: 'Сделка.UTM_TERM',
      REFERRER: 'Сделка.utm_referrer',
      SOURCE_TYPE: 'Сделка.Источник',
      DEAL_SOURCE: 'Сделка.R.Источник сделки',
      LEAD_TYPE: 'Сделка.Тип лида',
      REFERER: 'Сделка.REFERER'
    },
    ANALYTICS: {
      YM_CLIENT_ID: 'Сделка.YM_CLIENT_ID',
      YM_UID: 'Сделка._ym_uid',
      GA_CLIENT_ID: 'Сделка.GA_CLIENT_ID',
      FORM_ID: 'Сделка.FORMID',
      FORM_NAME: 'Сделка.FORMNAME',
      BUTTON_TEXT: 'Сделка.BUTTON_TEXT',
      DATE: 'Сделка.DATE',
      TIME: 'Сделка.TIME'
    },
    ADDITIONAL: {
      CITY_TAG: 'Сделка.R.Тег города',
      SOFTWARE: 'Сделка.ПО',
      REJECTION_REASON: 'Сделка.Причина отказа (ОБ)',
      NOTE: 'Примечание 1',
      RELATED_DEALS: 'Связанные сделки',
      MERGED: 'Объединено'
    },
    CHANGES: {
      CHANGED_AT: 'Изменено',
      CHANGED_FIELDS: 'Измененные поля'
    }
  },

  // Маппинг колонок Reserves RP - подтверждено диагностикой
  RESERVES_COLUMNS: {
    ID: 0, // [0]: ID
    RESERVE_ID: 1, // [1]: № заявки
    NAME: 2, // [2]: Имя
    PHONE: 3, // [3]: Телефон
    EMAIL: 4, // [4]: Email
    DATETIME: 5, // [5]: Дата/время
    STATUS: 6, // [6]: Статус
    COMMENT: 7, // [7]: Комментарий
    AMOUNT: 8, // [8]: Счёт, ₽
    GUESTS: 9, // [9]: Гостей
    SOURCE: 10 // [10]: Источник
  },

  // Маппинг колонок Guests RP - подтверждено диагностикой
  GUESTS_COLUMNS: {
    NAME: 0, // [0]: Имя
    PHONE: 1, // [1]: Телефон
    EMAIL: 2, // [2]: Email
    VISITS_COUNT: 3, // [3]: Кол-во визитов
    TOTAL_AMOUNT: 4, // [4]: Общая сумма
    FIRST_VISIT: 5, // [5]: Первый визит
    LAST_VISIT: 6 // [6]: Последний визит
  },

  // Маппинг колонок заявок с сайта - подтверждено диагностикой
  SITE_COLUMNS: {
    NAME: 0, // [0]: Name
    PHONE: 1, // [1]: Phone
    EMAIL: 6, // [6]: Email
    DATE: 7, // [7]: Date
    QUANTITY: 8, // [8]: Quantity
    FORM_NAME: 10, // [10]: Form name
    TIME: 11, // [11]: Time
    UTM_SOURCE: 17, // [17]: utm_source
    UTM_MEDIUM: 19, // [19]: utm_medium
    UTM_CAMPAIGN: 16, // [16]: utm_campaign
    REFERER: 2, // [2]: referer
    BUTTON_TEXT: 23 // [23]: button_text
  },

  // Маппинг колонок бюджетов
  BUDGET_COLUMNS: {
    CHANNEL: 0,
    TAGS: 1,
    MONTHS_START: 2
  },

  // Настройки обновления
  UPDATE_SETTINGS: {
    BATCH_SIZE: 1000,
    CACHE_HOURS: 1,
    MAX_EXECUTION_TIME: 300
  },

  // Настройки кеша
  CACHE_DURATION: 21600, // 6 часов в секундах

  // Настройки отчетов
  REPORT_SETTINGS: {
    DEFAULT_CHART_HEIGHT: 400,
    CHART_COLORS: ['#3366CC', '#DC3912', '#FF9900', '#109618', '#990099', '#0099C6', '#DD4477']
  },

  // Настройки анализа
  ANALYSIS_SETTINGS: {
    DAYS_FOR_INACTIVE: 90, // Дней без активности для неактивных клиентов
    VIP_MIN_VISITS: 5, // Минимум визитов для VIP
    VIP_MIN_AMOUNT: 50000, // Минимальная сумма для VIP
    REGULAR_MIN_VISITS: 2 // Минимум визитов для постоянных клиентов
  },

  // Отладочные настройки
  DEBUG: {
    ENABLED: true, // Включить расширенное логирование
    SHOW_PHONE_SAMPLES: true, // Показывать примеры обработанных телефонов
    LOG_DETAIL_LEVEL: 2 // Уровень детализации логов (1-3)
  }
};
