/**
 * Общие функции для всех модулей
 * Restaurant Analytics System - Utilities
 * Версия: 3.2
 */

/**
 * Очистка телефонного номера для унификации
 * Сохраняет стандартный формат с кодом 7
 */
function cleanPhone(phone) {
  if (!phone) return '';
  
  // Преобразуем в строку для обработки
  let phoneStr = String(phone);
  
  // Проверяем, не число ли уже это (Excel может хранить как число)
  if (typeof phone === 'number') {
    phoneStr = phone.toString();
  }
  
  // Удаляем все нецифровые символы
  let cleaned = phoneStr.replace(/\D/g, '');
  
  // Обработка российских номеров
  if (cleaned.length === 11) {
    if (cleaned.startsWith('8') || cleaned.startsWith('7')) {
      // Удаляем первую цифру для стандартизации
      cleaned = cleaned.substring(1);
    }
  }
  
  // Обработка номеров с кодом страны +7
  if (cleaned.length === 12 && cleaned.startsWith('7')) {
    cleaned = cleaned.substring(1);
  }
  
  // Логируем образцы телефонов для диагностики
  if (CONFIG.DEBUG && CONFIG.DEBUG.SHOW_PHONE_SAMPLES && Math.random() < 0.01) { // 1% для диагностики
    Logger.log(`Пример обработки телефона: ${phone} → ${cleaned}`);
  }
  
  return cleaned;
}

/**
 * Очистка email для унификации
 */
function cleanEmail(email) {
  if (!email) return '';
  return String(email).toLowerCase().trim();
}

/**
 * Форматирование даты
 */
function formatDate(date) {
  if (!date) return '';
  
  // Если это уже строка в формате даты, возвращаем как есть
  if (typeof date === 'string' && date.match(/^\d{4}-\d{2}-\d{2}/)) {
    return date.split(' ')[0]; // Берем только дату без времени
  }
  
  // Если это объект Date
  if (date instanceof Date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  
  // Пытаемся преобразовать в дату
  try {
    const dateObj = new Date(date);
    if (!isNaN(dateObj.getTime())) {
      return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  } catch (e) {
    // Игнорируем ошибки парсинга
  }
  
  return String(date);
}

/**
 * Парсинг числа с обработкой ошибок
 */
function parseNumber(value) {
  if (!value) return 0;
  
  // Если уже число, возвращаем как есть
  if (typeof value === 'number') return value;
  
  // Преобразуем в строку и удаляем все кроме цифр, точки и минуса
  const num = parseFloat(String(value).replace(/[^\d.-]/g, ''));
  return isNaN(num) ? 0 : num;
}

/**
 * Форматирование валюты
 */
function formatCurrency(amount) {
  try {
    return Number(amount).toLocaleString('ru-RU') + ' ₽';
  } catch (error) {
    return '0 ₽';
  }
}

/**
 * Форматирование процентов
 */
function formatPercent(value) {
  try {
    return Number(value).toFixed(2) + '%';
  } catch (error) {
    return '0%';
  }
}

/**
 * Форматирование числа
 */
function formatNumber(value, decimals = 0) {
  try {
    return Number(value).toFixed(decimals);
  } catch (error) {
    return '0';
  }
}

/**
 * Проверяет, является ли значение пустым
 */
function isEmpty(value) {
  return value === undefined || value === null || value === '' ||
    (typeof value === 'string' && value.trim() === '');
}

/**
 * Безопасное преобразование даты
 */
function safeParseDate(dateStr) {
  try {
    if (!dateStr) return null;
    if (dateStr instanceof Date) {
      return isNaN(dateStr.getTime()) ? null : dateStr;
    }
    const date = new Date(dateStr);
    return isNaN(date.getTime()) ? null : date;
  } catch (e) {
    return null;
  }
}

/**
 * Находит индекс колонки по названию (нечувствительно к регистру)
 */
function findColumnIndex(headerRow, columnName) {
  columnName = columnName.toString().trim().toLowerCase();
  for (let i = 0; i < headerRow.length; i++) {
    const header = headerRow[i].toString().trim().toLowerCase();
    if (header === columnName) {
      return i;
    }
  }
  return -1;
}

/**
 * Читает данные с листа с обработкой ошибок
 */
function readSheetData(spreadsheet, sheetName) {
  try {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log('Лист ' + sheetName + ' не найден');
      return [];
    }
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) {
      Logger.log('Лист ' + sheetName + ' пустой');
      return [];
    }
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    Logger.log('Прочитано из ' + sheetName + ': ' + data.length + ' строк');
    return data;
  } catch (error) {
    Logger.log('Ошибка чтения ' + sheetName + ': ' + error.toString());
    return [];
  }
}
