/**
 * Скрипт для Google Таблиц - Интеграция с Ozon Seller API
 *
 * Инструкция по установке:
 * 1. Откройте Google Таблицу
 * 2. Расширения → Apps Script
 * 3. Скопируйте этот код
 * 4. Сохраните проект
 * 5. Введите свои Client-ID и API-Key в настройки
 * 6. Запустите функцию setupMenu() или перезагрузите таблицу
 * 7. Используйте меню "Ozon Seller" для загрузки данных
 */

// ============================================================================
// НАСТРОЙКИ API
// ============================================================================

/**
 * Сохранить Client-ID и API-Key в свойствах скрипта
 * Запустите эту функцию один раз для сохранения ваших учетных данных
 */
function saveApiCredentials() {
  var ui = SpreadsheetApp.getUi();

  var clientIdResponse = ui.prompt(
    'Настройка API',
    'Введите ваш Client-ID:',
    ui.ButtonSet.OK_CANCEL
  );

  if (clientIdResponse.getSelectedButton() == ui.Button.OK) {
    var clientId = clientIdResponse.getResponseText();

    var apiKeyResponse = ui.prompt(
      'Настройка API',
      'Введите ваш API-Key:',
      ui.ButtonSet.OK_CANCEL
    );

    if (apiKeyResponse.getSelectedButton() == ui.Button.OK) {
      var apiKey = apiKeyResponse.getResponseText();

      var properties = PropertiesService.getScriptProperties();
      properties.setProperty('OZON_CLIENT_ID', clientId);
      properties.setProperty('OZON_API_KEY', apiKey);

      ui.alert('Успешно!', 'API учетные данные сохранены', ui.ButtonSet.OK);
    }
  }
}

/**
 * Получить сохраненные учетные данные API
 */
function getApiCredentials() {
  var properties = PropertiesService.getScriptProperties();
  return {
    clientId: properties.getProperty('OZON_CLIENT_ID'),
    apiKey: properties.getProperty('OZON_API_KEY')
  };
}

// ============================================================================
// КОНСТАНТЫ API
// ============================================================================

var OZON_API_BASE_URL = 'https://api-seller.ozon.ru';

// Статусы видимости товаров
var VISIBILITY = {
  ALL: 'ALL',
  VISIBLE: 'VISIBLE',
  INVISIBLE: 'INVISIBLE',
  EMPTY_STOCK: 'EMPTY_STOCK',
  IN_SALE: 'IN_SALE',
  REMOVED_FROM_SALE: 'REMOVED_FROM_SALE',
  ARCHIVED: 'ARCHIVED'
};

// ============================================================================
// МЕНЮ GOOGLE ТАБЛИЦ
// ============================================================================

/**
 * Создает пользовательское меню при открытии таблицы
 */
function onOpen() {
  setupMenu();
}

/**
 * Настройка меню
 */
function setupMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Ozon Seller')
    .addItem('⚙️ Настроить API ключи', 'saveApiCredentials')
    .addSeparator()
    .addItem('📦 Загрузить все товары', 'loadAllProducts')
    .addItem('📊 Загрузить товары с остатками и ценами', 'loadProductsWithStocksAndPrices')
    .addSeparator()
    .addItem('🔄 Обновить остатки', 'updateStocks')
    .addItem('💰 Обновить цены', 'updatePrices')
    .addSeparator()
    .addItem('🗑️ Очистить лист', 'clearSheet')
    .addToUi();
}

// ============================================================================
// ФУНКЦИИ API ЗАПРОСОВ
// ============================================================================

/**
 * Универсальная функция для выполнения запросов к API Ozon
 */
function makeOzonRequest(endpoint, payload) {
  var credentials = getApiCredentials();

  if (!credentials.clientId || !credentials.apiKey) {
    throw new Error('API учетные данные не настроены. Используйте меню: Ozon Seller → Настроить API ключи');
  }

  var url = OZON_API_BASE_URL + endpoint;

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Client-Id': credentials.clientId,
      'Api-Key': credentials.apiKey
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();

    if (responseCode !== 200) {
      Logger.log('Ошибка API: ' + responseCode + ' - ' + responseText);
      throw new Error('Ошибка API: ' + responseCode + ' - ' + responseText);
    }

    return JSON.parse(responseText);
  } catch (e) {
    Logger.log('Ошибка запроса: ' + e.toString());
    throw e;
  }
}

/**
 * Получить список товаров (V3 API)
 * @param {Object} filter - Фильтр товаров
 * @param {string} lastId - ID последнего товара для пагинации
 * @param {number} limit - Лимит товаров (1-1000)
 */
function getProductList(filter, lastId, limit) {
  filter = filter || { visibility: VISIBILITY.ALL };
  lastId = lastId || '';
  limit = limit || 1000;

  var payload = {
    filter: filter,
    last_id: lastId,
    limit: limit
  };

  return makeOzonRequest('/v3/product/list', payload);
}

/**
 * Получить все товары (с пагинацией)
 */
function getAllProducts(filter) {
  var allProducts = [];
  var lastId = '';
  var hasMore = true;

  while (hasMore) {
    var response = getProductList(filter, lastId, 1000);

    if (response.result && response.result.items) {
      allProducts = allProducts.concat(response.result.items);
      lastId = response.result.last_id;

      // Если last_id пустой или количество товаров меньше лимита, то это последняя страница
      hasMore = lastId && response.result.items.length === 1000;

      // Пауза между запросами для избежания rate limit
      if (hasMore) {
        Utilities.sleep(500);
      }
    } else {
      hasMore = false;
    }
  }

  return allProducts;
}

/**
 * Получить остатки товаров (V4 API)
 * @param {Array} productIds - Массив ID товаров
 */
function getProductStocks(productIds) {
  var allStocks = [];
  var cursor = '';
  var hasMore = true;

  // Разбиваем на батчи по 100 товаров
  var batchSize = 100;
  for (var i = 0; i < productIds.length; i += batchSize) {
    var batch = productIds.slice(i, i + batchSize);
    cursor = '';
    hasMore = true;

    while (hasMore) {
      var payload = {
        filter: {
          product_id: batch,
          visibility: VISIBILITY.ALL
        },
        cursor: cursor,
        limit: 1000
      };

      var response = makeOzonRequest('/v4/product/info/stocks', payload);

      if (response.result && response.result.items) {
        allStocks = allStocks.concat(response.result.items);
        cursor = response.result.cursor || '';
        hasMore = cursor !== '';

        if (hasMore) {
          Utilities.sleep(500);
        }
      } else {
        hasMore = false;
      }
    }

    // Пауза между батчами
    if (i + batchSize < productIds.length) {
      Utilities.sleep(1000);
    }
  }

  return allStocks;
}

/**
 * Получить цены товаров (V5 API)
 * @param {Array} productIds - Массив ID товаров
 */
function getProductPrices(productIds) {
  var allPrices = [];
  var cursor = '';
  var hasMore = true;

  // Разбиваем на батчи по 100 товаров
  var batchSize = 100;
  for (var i = 0; i < productIds.length; i += batchSize) {
    var batch = productIds.slice(i, i + batchSize);
    cursor = '';
    hasMore = true;

    while (hasMore) {
      // V5 API требует product_id как массив строк
      var stringIds = batch.map(function(id) { return String(id); });

      var payload = {
        filter: {
          product_id: stringIds,
          visibility: VISIBILITY.ALL
        },
        cursor: cursor,
        limit: 1000
      };

      var response = makeOzonRequest('/v5/product/info/prices', payload);

      if (response.result && response.result.items) {
        allPrices = allPrices.concat(response.result.items);
        cursor = response.result.cursor || '';
        hasMore = cursor !== '';

        if (hasMore) {
          Utilities.sleep(500);
        }
      } else {
        hasMore = false;
      }
    }

    // Пауза между батчами
    if (i + batchSize < productIds.length) {
      Utilities.sleep(1000);
    }
  }

  return allPrices;
}

// ============================================================================
// ФУНКЦИИ ОБРАБОТКИ И ЗАПИСИ ДАННЫХ
// ============================================================================

/**
 * Загрузить все товары в таблицу
 */
function loadAllProducts() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Создаем или получаем лист
  var sheet = ss.getSheetByName('Товары') || ss.insertSheet('Товары');
  sheet.clear();

  ui.alert('Загрузка данных', 'Начинаем загрузку товаров...', ui.ButtonSet.OK);

  try {
    // Получаем все товары
    var products = getAllProducts({ visibility: VISIBILITY.ALL });

    if (products.length === 0) {
      ui.alert('Внимание', 'Товары не найдены', ui.ButtonSet.OK);
      return;
    }

    // Заголовки
    var headers = [
      'Product ID',
      'Offer ID',
      'Статус архивации',
      'Есть остатки FBO',
      'Есть остатки FBS',
      'Есть скидка'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Данные
    var data = products.map(function(product) {
      return [
        product.product_id || '',
        product.offer_id || '',
        product.archived ? 'Да' : 'Нет',
        product.has_fbo_stocks ? 'Да' : 'Нет',
        product.has_fbs_stocks ? 'Да' : 'Нет',
        product.is_discounted ? 'Да' : 'Нет'
      ];
    });

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    }

    // Автоподбор ширины столбцов
    for (var i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }

    ui.alert('Готово!', 'Загружено товаров: ' + products.length, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString(), ui.ButtonSet.OK);
    Logger.log('Ошибка: ' + e.toString());
  }
}

/**
 * Загрузить товары с остатками и ценами
 */
function loadProductsWithStocksAndPrices() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Создаем или получаем лист
  var sheet = ss.getSheetByName('Товары (полные данные)') || ss.insertSheet('Товары (полные данные)');
  sheet.clear();

  ui.alert('Загрузка данных', 'Начинаем загрузку товаров с остатками и ценами...\nЭто может занять несколько минут.', ui.ButtonSet.OK);

  try {
    // 1. Получаем все товары
    Logger.log('Загрузка списка товаров...');
    var products = getAllProducts({ visibility: VISIBILITY.ALL });

    if (products.length === 0) {
      ui.alert('Внимание', 'Товары не найдены', ui.ButtonSet.OK);
      return;
    }

    Logger.log('Загружено товаров: ' + products.length);

    // 2. Получаем ID товаров
    var productIds = products.map(function(p) { return p.product_id; });

    // 3. Получаем остатки
    Logger.log('Загрузка остатков...');
    var stocks = getProductStocks(productIds);
    Logger.log('Загружено остатков для ' + stocks.length + ' товаров');

    // 4. Получаем цены
    Logger.log('Загрузка цен...');
    var prices = getProductPrices(productIds);
    Logger.log('Загружено цен для ' + prices.length + ' товаров');

    // 5. Создаем индексы для быстрого поиска
    var stocksMap = {};
    stocks.forEach(function(item) {
      stocksMap[item.product_id] = item.stocks;
    });

    var pricesMap = {};
    prices.forEach(function(item) {
      pricesMap[item.product_id] = item.price;
    });

    // 6. Заголовки
    var headers = [
      'Product ID',
      'Offer ID',
      'Статус',
      'FBO остаток',
      'FBS остаток',
      'Цена',
      'Старая цена',
      'Цена с премиум',
      'Рекомендованная цена',
      'Минимальная цена',
      'Маркетинговая цена',
      'Маркетинговая акция'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);

    // 7. Данные
    var data = products.map(function(product) {
      var productId = product.product_id;

      // Получаем остатки
      var fboStock = 0;
      var fbsStock = 0;

      if (stocksMap[productId]) {
        stocksMap[productId].forEach(function(stock) {
          if (stock.type === 'fbo') {
            fboStock += (stock.present || 0);
          } else if (stock.type === 'fbs') {
            fbsStock += (stock.present || 0);
          }
        });
      }

      // Получаем цены
      var priceData = pricesMap[productId] || {};
      var price = priceData.price || '';
      var oldPrice = priceData.old_price || '';
      var premiumPrice = priceData.premium_price || '';
      var recommendedPrice = priceData.recommended_price || '';
      var minPrice = priceData.min_price || '';
      var marketingPrice = priceData.marketing_price || '';
      var marketingAction = priceData.marketing_seller_price ? 'Да' : 'Нет';

      // Определяем статус
      var status = 'Неизвестно';
      if (product.archived) {
        status = 'В архиве';
      } else if (fboStock > 0 || fbsStock > 0) {
        status = 'В продаже';
      } else {
        status = 'Нет остатков';
      }

      return [
        productId,
        product.offer_id || '',
        status,
        fboStock,
        fbsStock,
        price,
        oldPrice,
        premiumPrice,
        recommendedPrice,
        minPrice,
        marketingPrice,
        marketingAction
      ];
    });

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    }

    // 8. Форматирование
    // Автоподбор ширины столбцов
    for (var i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }

    // Форматирование числовых столбцов (остатки)
    if (data.length > 0) {
      sheet.getRange(2, 4, data.length, 2).setNumberFormat('#,##0');
      // Форматирование цен
      sheet.getRange(2, 6, data.length, 6).setNumberFormat('#,##0.00 ₽');
    }

    ui.alert('Готово!',
      'Загружено товаров: ' + products.length + '\n' +
      'Загружено остатков: ' + stocks.length + '\n' +
      'Загружено цен: ' + prices.length,
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString(), ui.ButtonSet.OK);
    Logger.log('Ошибка: ' + e.toString());
  }
}

/**
 * Обновить только остатки в существующей таблице
 */
function updateStocks() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();

  // Проверяем, что есть столбцы Product ID
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var productIdCol = headers.indexOf('Product ID') + 1;
  var fboStockCol = headers.indexOf('FBO остаток') + 1;
  var fbsStockCol = headers.indexOf('FBS остаток') + 1;

  if (productIdCol === 0) {
    ui.alert('Ошибка', 'Столбец "Product ID" не найден на текущем листе', ui.ButtonSet.OK);
    return;
  }

  if (fboStockCol === 0 || fbsStockCol === 0) {
    ui.alert('Ошибка', 'Столбцы остатков не найдены на текущем листе', ui.ButtonSet.OK);
    return;
  }

  try {
    // Получаем все Product ID из таблицы
    var lastRow = sheet.getLastRow();
    var productIds = sheet.getRange(2, productIdCol, lastRow - 1, 1)
      .getValues()
      .map(function(row) { return row[0]; })
      .filter(function(id) { return id !== ''; });

    if (productIds.length === 0) {
      ui.alert('Внимание', 'Товары не найдены в таблице', ui.ButtonSet.OK);
      return;
    }

    // Получаем остатки
    Logger.log('Загрузка остатков для ' + productIds.length + ' товаров...');
    var stocks = getProductStocks(productIds);

    // Создаем карту остатков
    var stocksMap = {};
    stocks.forEach(function(item) {
      var fboStock = 0;
      var fbsStock = 0;

      item.stocks.forEach(function(stock) {
        if (stock.type === 'fbo') {
          fboStock += (stock.present || 0);
        } else if (stock.type === 'fbs') {
          fbsStock += (stock.present || 0);
        }
      });

      stocksMap[item.product_id] = {
        fbo: fboStock,
        fbs: fbsStock
      };
    });

    // Обновляем остатки в таблице
    for (var i = 2; i <= lastRow; i++) {
      var productId = sheet.getRange(i, productIdCol).getValue();

      if (stocksMap[productId]) {
        sheet.getRange(i, fboStockCol).setValue(stocksMap[productId].fbo);
        sheet.getRange(i, fbsStockCol).setValue(stocksMap[productId].fbs);
      }
    }

    ui.alert('Готово!', 'Остатки обновлены для ' + productIds.length + ' товаров', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString(), ui.ButtonSet.OK);
    Logger.log('Ошибка: ' + e.toString());
  }
}

/**
 * Обновить только цены в существующей таблице
 */
function updatePrices() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();

  // Проверяем, что есть столбцы Product ID и цен
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var productIdCol = headers.indexOf('Product ID') + 1;
  var priceCol = headers.indexOf('Цена') + 1;

  if (productIdCol === 0) {
    ui.alert('Ошибка', 'Столбец "Product ID" не найден на текущем листе', ui.ButtonSet.OK);
    return;
  }

  if (priceCol === 0) {
    ui.alert('Ошибка', 'Столбец "Цена" не найден на текущем листе', ui.ButtonSet.OK);
    return;
  }

  try {
    // Получаем все Product ID из таблицы
    var lastRow = sheet.getLastRow();
    var productIds = sheet.getRange(2, productIdCol, lastRow - 1, 1)
      .getValues()
      .map(function(row) { return row[0]; })
      .filter(function(id) { return id !== ''; });

    if (productIds.length === 0) {
      ui.alert('Внимание', 'Товары не найдены в таблице', ui.ButtonSet.OK);
      return;
    }

    // Получаем цены
    Logger.log('Загрузка цен для ' + productIds.length + ' товаров...');
    var prices = getProductPrices(productIds);

    // Создаем карту цен
    var pricesMap = {};
    prices.forEach(function(item) {
      pricesMap[item.product_id] = item.price;
    });

    // Находим индексы столбцов цен (если есть)
    var oldPriceCol = headers.indexOf('Старая цена') + 1;
    var premiumPriceCol = headers.indexOf('Цена с премиум') + 1;
    var recommendedPriceCol = headers.indexOf('Рекомендованная цена') + 1;
    var minPriceCol = headers.indexOf('Минимальная цена') + 1;
    var marketingPriceCol = headers.indexOf('Маркетинговая цена') + 1;

    // Обновляем цены в таблице
    for (var i = 2; i <= lastRow; i++) {
      var productId = sheet.getRange(i, productIdCol).getValue();

      if (pricesMap[productId]) {
        var priceData = pricesMap[productId];

        sheet.getRange(i, priceCol).setValue(priceData.price || '');

        if (oldPriceCol > 0) sheet.getRange(i, oldPriceCol).setValue(priceData.old_price || '');
        if (premiumPriceCol > 0) sheet.getRange(i, premiumPriceCol).setValue(priceData.premium_price || '');
        if (recommendedPriceCol > 0) sheet.getRange(i, recommendedPriceCol).setValue(priceData.recommended_price || '');
        if (minPriceCol > 0) sheet.getRange(i, minPriceCol).setValue(priceData.min_price || '');
        if (marketingPriceCol > 0) sheet.getRange(i, marketingPriceCol).setValue(priceData.marketing_price || '');
      }
    }

    ui.alert('Готово!', 'Цены обновлены для ' + productIds.length + ' товаров', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString(), ui.ButtonSet.OK);
    Logger.log('Ошибка: ' + e.toString());
  }
}

/**
 * Очистить текущий лист
 */
function clearSheet() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Подтверждение',
    'Вы действительно хотите очистить текущий лист?',
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.clear();
    ui.alert('Готово', 'Лист очищен', ui.ButtonSet.OK);
  }
}
