/**
 * Скрипт для Google Таблиц - Интеграция с Ozon Seller API
 * БЫСТРЫЙ СТАРТ С ПРЕДЗАПОЛНЕННЫМИ КЛЮЧАМИ И ОТЛАДКОЙ
 *
 * ⚠️ ВНИМАНИЕ: Этот файл содержит ваши API ключи в коде!
 * Используйте его только для быстрого тестирования.
 */

// ============================================================================
// НАСТРОЙКИ API - ПРЕДЗАПОЛНЕННЫЕ КЛЮЧИ
// ============================================================================

var OZON_CLIENT_ID = '60328';
var OZON_API_KEY = '6c397ca2-9650-4fe5-8b3b-0e7f80ad6223';

/**
 * Получить учетные данные API (используются жестко заданные значения)
 */
function getApiCredentials() {
  return {
    clientId: OZON_CLIENT_ID,
    apiKey: OZON_API_KEY
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
    .addItem('📦 Загрузить все товары', 'loadAllProducts')
    .addItem('📊 Загрузить товары с остатками и ценами', 'loadProductsWithStocksAndPrices')
    .addSeparator()
    .addItem('🔄 Обновить остатки', 'updateStocks')
    .addItem('💰 Обновить цены', 'updatePrices')
    .addSeparator()
    .addItem('🔍 Тест API (отладка)', 'testApi')
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

    Logger.log('API Request: ' + endpoint);
    Logger.log('Response Code: ' + responseCode);
    Logger.log('Response: ' + responseText.substring(0, 500)); // Первые 500 символов

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
      hasMore = lastId && response.result.items.length === 1000;

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
 */
function getProductStocks(productIds) {
  if (!productIds || productIds.length === 0) {
    Logger.log('getProductStocks: Пустой массив productIds');
    return [];
  }

  var allStocks = [];
  var batchSize = 100;

  for (var i = 0; i < productIds.length; i += batchSize) {
    var batch = productIds.slice(i, Math.min(i + batchSize, productIds.length));
    Logger.log('Загрузка остатков для батча: ' + i + ' - ' + (i + batch.length) + ' из ' + productIds.length);

    var cursor = '';
    var hasMore = true;
    var iterations = 0;

    while (hasMore && iterations < 10) { // Защита от бесконечного цикла
      iterations++;

      var payload = {
        filter: {
          product_id: batch,
          visibility: VISIBILITY.ALL
        },
        cursor: cursor,
        limit: 1000
      };

      Logger.log('Запрос остатков: cursor=' + cursor + ', товаров в батче: ' + batch.length);
      var response = makeOzonRequest('/v4/product/info/stocks', payload);

      if (response && response.result && response.result.items) {
        Logger.log('Получено остатков: ' + response.result.items.length);
        allStocks = allStocks.concat(response.result.items);

        cursor = response.result.cursor || '';
        hasMore = cursor !== '' && cursor !== null;

        if (hasMore) {
          Utilities.sleep(500);
        }
      } else {
        Logger.log('Нет данных в ответе остатков');
        hasMore = false;
      }
    }

    if (i + batchSize < productIds.length) {
      Utilities.sleep(1000);
    }
  }

  Logger.log('Всего загружено остатков: ' + allStocks.length);
  return allStocks;
}

/**
 * Получить цены товаров (V5 API)
 */
function getProductPrices(productIds) {
  if (!productIds || productIds.length === 0) {
    Logger.log('getProductPrices: Пустой массив productIds');
    return [];
  }

  var allPrices = [];
  var batchSize = 100;

  for (var i = 0; i < productIds.length; i += batchSize) {
    var batch = productIds.slice(i, Math.min(i + batchSize, productIds.length));
    Logger.log('Загрузка цен для батча: ' + i + ' - ' + (i + batch.length) + ' из ' + productIds.length);

    var cursor = '';
    var hasMore = true;
    var iterations = 0;

    while (hasMore && iterations < 10) { // Защита от бесконечного цикла
      iterations++;

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

      Logger.log('Запрос цен: cursor=' + cursor + ', товаров в батче: ' + stringIds.length);
      var response = makeOzonRequest('/v5/product/info/prices', payload);

      if (response && response.result && response.result.items) {
        Logger.log('Получено цен: ' + response.result.items.length);
        allPrices = allPrices.concat(response.result.items);

        cursor = response.result.cursor || '';
        hasMore = cursor !== '' && cursor !== null;

        if (hasMore) {
          Utilities.sleep(500);
        }
      } else {
        Logger.log('Нет данных в ответе цен');
        hasMore = false;
      }
    }

    if (i + batchSize < productIds.length) {
      Utilities.sleep(1000);
    }
  }

  Logger.log('Всего загружено цен: ' + allPrices.length);
  return allPrices;
}

// ============================================================================
// ТЕСТОВАЯ ФУНКЦИЯ ДЛЯ ОТЛАДКИ
// ============================================================================

/**
 * Тест API - для отладки
 */
function testApi() {
  var ui = SpreadsheetApp.getUi();

  try {
    // Тест 1: Получить первые 10 товаров
    Logger.log('=== ТЕСТ 1: Получение товаров ===');
    var products = getAllProducts({ visibility: VISIBILITY.ALL });
    Logger.log('Получено товаров: ' + products.length);

    if (products.length > 0) {
      Logger.log('Первый товар: ' + JSON.stringify(products[0]));

      // Тест 2: Получить остатки для первых 5 товаров
      Logger.log('=== ТЕСТ 2: Получение остатков ===');
      var testProductIds = products.slice(0, 5).map(function(p) { return p.product_id; });
      Logger.log('Тестовые Product IDs: ' + testProductIds.join(', '));

      var stocks = getProductStocks(testProductIds);
      Logger.log('Получено остатков: ' + stocks.length);
      if (stocks.length > 0) {
        Logger.log('Первый остаток: ' + JSON.stringify(stocks[0]));
      }

      // Тест 3: Получить цены для первых 5 товаров
      Logger.log('=== ТЕСТ 3: Получение цен ===');
      var prices = getProductPrices(testProductIds);
      Logger.log('Получено цен: ' + prices.length);
      if (prices.length > 0) {
        Logger.log('Первая цена: ' + JSON.stringify(prices[0]));
      }

      ui.alert('Тест завершен',
        'Товаров: ' + products.length + '\n' +
        'Остатков (тест): ' + stocks.length + '\n' +
        'Цен (тест): ' + prices.length + '\n\n' +
        'Подробности в логах (Ctrl+Enter или View → Logs)',
        ui.ButtonSet.OK);
    } else {
      ui.alert('Ошибка', 'Товары не найдены', ui.ButtonSet.OK);
    }

  } catch (e) {
    Logger.log('Ошибка теста: ' + e.toString());
    ui.alert('Ошибка теста', e.toString(), ui.ButtonSet.OK);
  }
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
  var sheet = ss.getSheetByName('Товары') || ss.insertSheet('Товары');
  sheet.clear();

  ui.alert('Загрузка данных', 'Начинаем загрузку товаров...', ui.ButtonSet.OK);

  try {
    var products = getAllProducts({ visibility: VISIBILITY.ALL });

    if (products.length === 0) {
      ui.alert('Внимание', 'Товары не найдены', ui.ButtonSet.OK);
      return;
    }

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
  var sheet = ss.getSheetByName('Товары (полные данные)') || ss.insertSheet('Товары (полные данные)');
  sheet.clear();

  ui.alert('Загрузка данных', 'Начинаем загрузку товаров с остатками и ценами...\nЭто может занять несколько минут.', ui.ButtonSet.OK);

  try {
    // 1. Получаем все товары
    Logger.log('=== Загрузка списка товаров ===');
    var products = getAllProducts({ visibility: VISIBILITY.ALL });

    if (products.length === 0) {
      ui.alert('Внимание', 'Товары не найдены', ui.ButtonSet.OK);
      return;
    }

    Logger.log('Загружено товаров: ' + products.length);

    // 2. Получаем ID товаров
    var productIds = products.map(function(p) { return p.product_id; });
    Logger.log('Product IDs для загрузки: ' + productIds.length);
    Logger.log('Первые 5 Product IDs: ' + productIds.slice(0, 5).join(', '));

    // 3. Получаем остатки
    Logger.log('=== Загрузка остатков ===');
    var stocks = getProductStocks(productIds);
    Logger.log('Загружено остатков для ' + stocks.length + ' товаров');

    // 4. Получаем цены
    Logger.log('=== Загрузка цен ===');
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

    Logger.log('StocksMap размер: ' + Object.keys(stocksMap).length);
    Logger.log('PricesMap размер: ' + Object.keys(pricesMap).length);

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
    for (var i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }

    if (data.length > 0) {
      sheet.getRange(2, 4, data.length, 2).setNumberFormat('#,##0');
      sheet.getRange(2, 6, data.length, 6).setNumberFormat('#,##0.00 ₽');
    }

    ui.alert('Готово!',
      'Загружено товаров: ' + products.length + '\n' +
      'Загружено остатков: ' + stocks.length + '\n' +
      'Загружено цен: ' + prices.length + '\n\n' +
      'Проверьте лист "Товары (полные данные)"',
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString() + '\n\nПроверьте логи (View → Logs)', ui.ButtonSet.OK);
    Logger.log('Ошибка: ' + e.toString());
  }
}

/**
 * Обновить только остатки в существующей таблице
 */
function updateStocks() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();

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
    var lastRow = sheet.getLastRow();
    var productIds = sheet.getRange(2, productIdCol, lastRow - 1, 1)
      .getValues()
      .map(function(row) { return row[0]; })
      .filter(function(id) { return id !== ''; });

    if (productIds.length === 0) {
      ui.alert('Внимание', 'Товары не найдены в таблице', ui.ButtonSet.OK);
      return;
    }

    Logger.log('Загрузка остатков для ' + productIds.length + ' товаров...');
    var stocks = getProductStocks(productIds);

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
    var lastRow = sheet.getLastRow();
    var productIds = sheet.getRange(2, productIdCol, lastRow - 1, 1)
      .getValues()
      .map(function(row) { return row[0]; })
      .filter(function(id) { return id !== ''; });

    if (productIds.length === 0) {
      ui.alert('Внимание', 'Товары не найдены в таблице', ui.ButtonSet.OK);
      return;
    }

    Logger.log('Загрузка цен для ' + productIds.length + ' товаров...');
    var prices = getProductPrices(productIds);

    var pricesMap = {};
    prices.forEach(function(item) {
      pricesMap[item.product_id] = item.price;
    });

    var oldPriceCol = headers.indexOf('Старая цена') + 1;
    var premiumPriceCol = headers.indexOf('Цена с премиум') + 1;
    var recommendedPriceCol = headers.indexOf('Рекомендованная цена') + 1;
    var minPriceCol = headers.indexOf('Минимальная цена') + 1;
    var marketingPriceCol = headers.indexOf('Маркетинговая цена') + 1;

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
