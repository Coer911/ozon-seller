/**
 * Скрипт для Google Таблиц - Интеграция с Ozon Seller API
 * ИСПРАВЛЕННАЯ ВЕРСИЯ - правильная обработка структуры ответов API
 */

// ============================================================================
// НАСТРОЙКИ API
// ============================================================================

var OZON_CLIENT_ID = '60328';
var OZON_API_KEY = '6c397ca2-9650-4fe5-8b3b-0e7f80ad6223';

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
// МЕНЮ
// ============================================================================

function onOpen() {
  setupMenu();
}

function setupMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Ozon Seller')
    .addItem('📦 Загрузить все товары', 'loadAllProducts')
    .addItem('📊 Загрузить товары с остатками и ценами', 'loadProductsWithStocksAndPrices')
    .addSeparator()
    .addItem('🔍 Тест одного товара (отладка)', 'testSingleProduct')
    .addItem('🔄 Обновить остатки', 'updateStocks')
    .addItem('💰 Обновить цены', 'updatePrices')
    .addSeparator()
    .addItem('🗑️ Очистить лист', 'clearSheet')
    .addToUi();
}

// ============================================================================
// ФУНКЦИИ API
// ============================================================================

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

    if (responseCode !== 200) {
      Logger.log('Ошибка API (' + endpoint + '): ' + responseCode + ' - ' + responseText);
      throw new Error('Ошибка API: ' + responseCode + ' - ' + responseText);
    }

    var data = JSON.parse(responseText);
    Logger.log('API ' + endpoint + ' - успешно, элементов: ' + (data.result && data.result.items ? data.result.items.length : 0));
    return data;
  } catch (e) {
    Logger.log('Ошибка запроса (' + endpoint + '): ' + e.toString());
    throw e;
  }
}

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

function getProductStocks(productIds) {
  if (!productIds || productIds.length === 0) {
    return [];
  }

  var allStocks = [];
  var batchSize = 100;

  for (var i = 0; i < productIds.length; i += batchSize) {
    var batch = productIds.slice(i, Math.min(i + batchSize, productIds.length));

    var cursor = '';
    var hasMore = true;
    var iterations = 0;

    while (hasMore && iterations < 10) {
      iterations++;

      var payload = {
        filter: {
          product_id: batch,
          visibility: VISIBILITY.ALL
        },
        cursor: cursor,
        limit: 1000
      };

      var response = makeOzonRequest('/v4/product/info/stocks', payload);

      if (response && response.result && response.result.items) {
        allStocks = allStocks.concat(response.result.items);
        cursor = response.result.cursor || '';
        hasMore = cursor !== '' && cursor !== null;

        if (hasMore) {
          Utilities.sleep(500);
        }
      } else {
        hasMore = false;
      }
    }

    if (i + batchSize < productIds.length) {
      Utilities.sleep(1000);
    }
  }

  return allStocks;
}

function getProductPrices(productIds) {
  if (!productIds || productIds.length === 0) {
    return [];
  }

  var allPrices = [];
  var batchSize = 100;

  for (var i = 0; i < productIds.length; i += batchSize) {
    var batch = productIds.slice(i, Math.min(i + batchSize, productIds.length));

    var cursor = '';
    var hasMore = true;
    var iterations = 0;

    while (hasMore && iterations < 10) {
      iterations++;

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

      if (response && response.result && response.result.items) {
        allPrices = allPrices.concat(response.result.items);
        cursor = response.result.cursor || '';
        hasMore = cursor !== '' && cursor !== null;

        if (hasMore) {
          Utilities.sleep(500);
        }
      } else {
        hasMore = false;
      }
    }

    if (i + batchSize < productIds.length) {
      Utilities.sleep(1000);
    }
  }

  return allPrices;
}

// ============================================================================
// ТЕСТ ОДНОГО ТОВАРА
// ============================================================================

function testSingleProduct() {
  var ui = SpreadsheetApp.getUi();

  try {
    Logger.log('=== ТЕСТ: Загрузка одного товара ===');

    // Получаем первый товар
    var products = getAllProducts({ visibility: VISIBILITY.ALL });
    if (products.length === 0) {
      ui.alert('Ошибка', 'Товары не найдены', ui.ButtonSet.OK);
      return;
    }

    var testProduct = products[0];
    var testProductId = testProduct.product_id;

    Logger.log('Тестовый товар: ' + JSON.stringify(testProduct));
    Logger.log('Product ID: ' + testProductId);
    Logger.log('Offer ID: ' + testProduct.offer_id);

    // Получаем остатки
    Logger.log('=== Запрос остатков ===');
    var stocksResponse = makeOzonRequest('/v4/product/info/stocks', {
      filter: {
        product_id: [testProductId],
        visibility: VISIBILITY.ALL
      },
      cursor: '',
      limit: 100
    });

    Logger.log('Ответ остатков (весь): ' + JSON.stringify(stocksResponse));

    // Получаем цены
    Logger.log('=== Запрос цен ===');
    var pricesResponse = makeOzonRequest('/v5/product/info/prices', {
      filter: {
        product_id: [String(testProductId)],
        visibility: VISIBILITY.ALL
      },
      cursor: '',
      limit: 100
    });

    Logger.log('Ответ цен (весь): ' + JSON.stringify(pricesResponse));

    var message = 'Тест выполнен!\n\n';
    message += 'Товар: ' + testProduct.offer_id + '\n';
    message += 'Product ID: ' + testProductId + '\n\n';

    if (stocksResponse.result && stocksResponse.result.items && stocksResponse.result.items.length > 0) {
      message += 'Остатки найдены: ДА\n';
      message += 'Структура остатков:\n' + JSON.stringify(stocksResponse.result.items[0], null, 2).substring(0, 200) + '...\n\n';
    } else {
      message += 'Остатки найдены: НЕТ\n\n';
    }

    if (pricesResponse.result && pricesResponse.result.items && pricesResponse.result.items.length > 0) {
      message += 'Цены найдены: ДА\n';
      message += 'Структура цен:\n' + JSON.stringify(pricesResponse.result.items[0], null, 2).substring(0, 200) + '...\n\n';
    } else {
      message += 'Цены найдены: НЕТ\n\n';
    }

    message += '\nПолные данные в логах (View → Logs)';

    ui.alert('Результат теста', message, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log('Ошибка теста: ' + e.toString());
    ui.alert('Ошибка', e.toString() + '\n\nСмотрите логи: View → Logs', ui.ButtonSet.OK);
  }
}

// ============================================================================
// ОСНОВНЫЕ ФУНКЦИИ
// ============================================================================

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

function loadProductsWithStocksAndPrices() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Товары (полные данные)') || ss.insertSheet('Товары (полные данные)');
  sheet.clear();

  ui.alert('Загрузка данных', 'Начинаем загрузку...\nЭто может занять несколько минут.', ui.ButtonSet.OK);

  try {
    Logger.log('=== Загрузка товаров ===');
    var products = getAllProducts({ visibility: VISIBILITY.ALL });

    if (products.length === 0) {
      ui.alert('Внимание', 'Товары не найдены', ui.ButtonSet.OK);
      return;
    }

    Logger.log('Загружено товаров: ' + products.length);

    var productIds = products.map(function(p) { return p.product_id; });

    Logger.log('=== Загрузка остатков ===');
    var stocks = getProductStocks(productIds);
    Logger.log('Загружено остатков: ' + stocks.length);

    Logger.log('=== Загрузка цен ===');
    var prices = getProductPrices(productIds);
    Logger.log('Загружено цен: ' + prices.length);

    // Создаем индексы
    var stocksMap = {};
    stocks.forEach(function(item) {
      stocksMap[item.product_id] = item.stocks;
    });

    var pricesMap = {};
    prices.forEach(function(item) {
      // ВАЖНО: сохраняем весь объект price, а не только item.price!
      pricesMap[item.product_id] = item.price || {};
    });

    Logger.log('StocksMap: ' + Object.keys(stocksMap).length + ' товаров');
    Logger.log('PricesMap: ' + Object.keys(pricesMap).length + ' товаров');

    // Заголовки
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
      'Маркетинговая цена'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Данные
    var data = products.map(function(product) {
      var productId = product.product_id;

      // Остатки
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

      // Цены - правильно извлекаем из объекта
      var priceData = pricesMap[productId] || {};
      var price = priceData.price || '';
      var oldPrice = priceData.old_price || '';
      var premiumPrice = priceData.premium_price || '';
      var recommendedPrice = priceData.recommended_price || '';
      var minPrice = priceData.min_price || '';
      var marketingPrice = priceData.marketing_price || '';

      // Статус
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
        marketingPrice
      ];
    });

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    }

    // Форматирование
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
      'Загружено цен: ' + prices.length,
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString() + '\n\nСмотрите логи: View → Logs', ui.ButtonSet.OK);
    Logger.log('Ошибка: ' + e.toString());
  }
}

function updateStocks() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var productIdCol = headers.indexOf('Product ID') + 1;
  var fboStockCol = headers.indexOf('FBO остаток') + 1;
  var fbsStockCol = headers.indexOf('FBS остаток') + 1;

  if (productIdCol === 0 || fboStockCol === 0 || fbsStockCol === 0) {
    ui.alert('Ошибка', 'Нужные столбцы не найдены на текущем листе', ui.ButtonSet.OK);
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

function updatePrices() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var productIdCol = headers.indexOf('Product ID') + 1;
  var priceCol = headers.indexOf('Цена') + 1;

  if (productIdCol === 0 || priceCol === 0) {
    ui.alert('Ошибка', 'Нужные столбцы не найдены', ui.ButtonSet.OK);
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

    var prices = getProductPrices(productIds);

    var pricesMap = {};
    prices.forEach(function(item) {
      pricesMap[item.product_id] = item.price || {};
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
