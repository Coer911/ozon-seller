# Примеры использования Ozon Seller Google Apps Script

Этот документ содержит практические примеры расширения базовой функциональности скрипта.

## Содержание

1. [Фильтрация товаров](#фильтрация-товаров)
2. [Кастомные отчеты](#кастомные-отчеты)
3. [Анализ данных](#анализ-данных)
4. [Автоматизация](#автоматизация)
5. [Работа с ценами](#работа-с-ценами)

## Фильтрация товаров

### Загрузить только товары в продаже

```javascript
function loadOnlyActiveProducts() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Активные товары') || ss.insertSheet('Активные товары');
  sheet.clear();

  try {
    // Получаем только товары в продаже
    var products = getAllProducts({ visibility: VISIBILITY.IN_SALE });

    // Заголовки
    var headers = ['Product ID', 'Offer ID', 'FBO', 'FBS'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

    // Данные
    var data = products.map(function(product) {
      return [
        product.product_id,
        product.offer_id,
        product.has_fbo_stocks ? 'Да' : 'Нет',
        product.has_fbs_stocks ? 'Да' : 'Нет'
      ];
    });

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    }

    ui.alert('Готово!', 'Загружено активных товаров: ' + products.length, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString(), ui.ButtonSet.OK);
  }
}
```

### Загрузить товары без остатков

```javascript
function loadProductsWithoutStock() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Без остатков') || ss.insertSheet('Без остатков');
  sheet.clear();

  try {
    var products = getAllProducts({ visibility: VISIBILITY.EMPTY_STOCK });

    var headers = ['Product ID', 'Offer ID'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

    var data = products.map(function(product) {
      return [product.product_id, product.offer_id];
    });

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    }

    ui.alert('Внимание!', 'Найдено товаров без остатков: ' + products.length, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString(), ui.ButtonSet.OK);
  }
}
```

### Загрузить конкретные товары по Offer ID

```javascript
function loadSpecificProducts() {
  var ui = SpreadsheetApp.getUi();

  // Укажите ваши Offer ID
  var offerIds = ['OFFER-001', 'OFFER-002', 'OFFER-003'];

  var payload = {
    filter: {
      offer_id: offerIds,
      visibility: VISIBILITY.ALL
    },
    last_id: '',
    limit: 1000
  };

  try {
    var response = makeOzonRequest('/v3/product/list', payload);

    if (response.result && response.result.items) {
      var products = response.result.items;

      // Получаем остатки и цены для этих товаров
      var productIds = products.map(function(p) { return p.product_id; });
      var stocks = getProductStocks(productIds);
      var prices = getProductPrices(productIds);

      // Создаем отчет
      // ... ваш код для вывода данных
      ui.alert('Готово!', 'Загружено товаров: ' + products.length, ui.ButtonSet.OK);
    }

  } catch (e) {
    ui.alert('Ошибка', e.toString(), ui.ButtonSet.OK);
  }
}
```

## Кастомные отчеты

### Отчет по маржинальности

```javascript
function createMarginReport() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Маржинальность') || ss.insertSheet('Маржинальность');
  sheet.clear();

  try {
    // Загружаем товары с ценами
    var products = getAllProducts({ visibility: VISIBILITY.IN_SALE });
    var productIds = products.map(function(p) { return p.product_id; });
    var prices = getProductPrices(productIds);

    // Создаем индекс цен
    var pricesMap = {};
    prices.forEach(function(item) {
      pricesMap[item.product_id] = item.price;
    });

    // Заголовки
    var headers = [
      'Product ID',
      'Offer ID',
      'Текущая цена',
      'Минимальная цена',
      'Рекомендованная цена',
      'Маржа от минимальной',
      'Статус цены'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

    // Данные
    var data = products.map(function(product) {
      var priceData = pricesMap[product.product_id] || {};
      var currentPrice = parseFloat(priceData.price) || 0;
      var minPrice = parseFloat(priceData.min_price) || 0;
      var recommendedPrice = parseFloat(priceData.recommended_price) || 0;

      var margin = minPrice > 0 ? ((currentPrice - minPrice) / minPrice * 100).toFixed(2) + '%' : 'N/A';
      var status = currentPrice < recommendedPrice ? 'Ниже рекомендованной' :
                   currentPrice > recommendedPrice ? 'Выше рекомендованной' :
                   'Равна рекомендованной';

      return [
        product.product_id,
        product.offer_id,
        currentPrice,
        minPrice,
        recommendedPrice,
        margin,
        status
      ];
    });

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);

      // Форматирование
      sheet.getRange(2, 3, data.length, 3).setNumberFormat('#,##0.00 ₽');
    }

    ui.alert('Готово!', 'Отчет по маржинальности создан', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString(), ui.ButtonSet.OK);
  }
}
```

### Отчет по товарам с низкими остатками

```javascript
function createLowStockReport() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Низкие остатки') || ss.insertSheet('Низкие остатки');
  sheet.clear();

  // Порог низких остатков
  var LOW_STOCK_THRESHOLD = 10;

  try {
    var products = getAllProducts({ visibility: VISIBILITY.IN_SALE });
    var productIds = products.map(function(p) { return p.product_id; });
    var stocks = getProductStocks(productIds);

    // Создаем индекс остатков
    var stocksMap = {};
    stocks.forEach(function(item) {
      var totalStock = 0;
      item.stocks.forEach(function(stock) {
        totalStock += (stock.present || 0);
      });
      stocksMap[item.product_id] = totalStock;
    });

    // Заголовки
    var headers = ['Product ID', 'Offer ID', 'Остаток', 'Статус'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

    // Фильтруем товары с низкими остатками
    var lowStockProducts = [];
    products.forEach(function(product) {
      var stock = stocksMap[product.product_id] || 0;

      if (stock > 0 && stock <= LOW_STOCK_THRESHOLD) {
        var status = stock <= 3 ? '🔴 Критический' :
                     stock <= 5 ? '🟠 Низкий' :
                     '🟡 Требует внимания';

        lowStockProducts.push([
          product.product_id,
          product.offer_id,
          stock,
          status
        ]);
      }
    });

    if (lowStockProducts.length > 0) {
      sheet.getRange(2, 1, lowStockProducts.length, headers.length).setValues(lowStockProducts);

      // Сортировка по остаткам (по возрастанию)
      var range = sheet.getRange(2, 1, lowStockProducts.length, headers.length);
      range.sort(3);
    }

    ui.alert('Внимание!',
      'Найдено товаров с низкими остатками: ' + lowStockProducts.length,
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString(), ui.ButtonSet.OK);
  }
}
```

### Сравнение цен: текущая vs рекомендованная

```javascript
function comparePrices() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Сравнение цен') || ss.insertSheet('Сравнение цен');
  sheet.clear();

  try {
    var products = getAllProducts({ visibility: VISIBILITY.IN_SALE });
    var productIds = products.map(function(p) { return p.product_id; });
    var prices = getProductPrices(productIds);

    var pricesMap = {};
    prices.forEach(function(item) {
      pricesMap[item.product_id] = item.price;
    });

    var headers = [
      'Product ID',
      'Offer ID',
      'Текущая цена',
      'Рекомендованная',
      'Разница (₽)',
      'Разница (%)',
      'Рекомендация'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

    var data = products.map(function(product) {
      var priceData = pricesMap[product.product_id] || {};
      var currentPrice = parseFloat(priceData.price) || 0;
      var recommendedPrice = parseFloat(priceData.recommended_price) || 0;

      var diff = currentPrice - recommendedPrice;
      var diffPercent = recommendedPrice > 0 ? (diff / recommendedPrice * 100).toFixed(2) : 0;

      var recommendation = '';
      if (diff > 0) {
        recommendation = '⬇️ Снизить цену';
      } else if (diff < 0) {
        recommendation = '⬆️ Можно повысить';
      } else {
        recommendation = '✅ Оптимальная';
      }

      return [
        product.product_id,
        product.offer_id,
        currentPrice,
        recommendedPrice,
        diff.toFixed(2),
        diffPercent + '%',
        recommendation
      ];
    });

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
      sheet.getRange(2, 3, data.length, 2).setNumberFormat('#,##0.00 ₽');
      sheet.getRange(2, 5, data.length, 1).setNumberFormat('#,##0.00 ₽');
    }

    ui.alert('Готово!', 'Отчет сравнения цен создан', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString(), ui.ButtonSet.OK);
  }
}
```

## Анализ данных

### Статистика по товарам

```javascript
function createProductStatistics() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Статистика') || ss.insertSheet('Статистика');
  sheet.clear();

  try {
    var products = getAllProducts({ visibility: VISIBILITY.ALL });
    var productIds = products.map(function(p) { return p.product_id; });
    var stocks = getProductStocks(productIds);

    // Подсчет статистики
    var stats = {
      total: products.length,
      archived: 0,
      inSale: 0,
      withDiscount: 0,
      fboStock: 0,
      fbsStock: 0,
      noStock: 0
    };

    products.forEach(function(product) {
      if (product.archived) stats.archived++;
      if (product.is_discounted) stats.withDiscount++;
      if (product.has_fbo_stocks || product.has_fbs_stocks) {
        stats.inSale++;
      }
    });

    stocks.forEach(function(item) {
      var hasFbo = false;
      var hasFbs = false;

      item.stocks.forEach(function(stock) {
        if (stock.type === 'fbo' && stock.present > 0) {
          hasFbo = true;
          stats.fboStock++;
        } else if (stock.type === 'fbs' && stock.present > 0) {
          hasFbs = true;
          stats.fbsStock++;
        }
      });

      if (!hasFbo && !hasFbs) {
        stats.noStock++;
      }
    });

    // Вывод статистики
    var data = [
      ['Метрика', 'Значение'],
      ['Всего товаров', stats.total],
      ['В архиве', stats.archived],
      ['В продаже', stats.inSale],
      ['Со скидкой', stats.withDiscount],
      ['С остатками FBO', stats.fboStock],
      ['С остатками FBS', stats.fbsStock],
      ['Без остатков', stats.noStock]
    ];

    sheet.getRange(1, 1, data.length, 2).setValues(data);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    sheet.getRange(2, 2, data.length - 1, 1).setNumberFormat('#,##0');

    // Автоподбор столбцов
    sheet.autoResizeColumn(1);
    sheet.autoResizeColumn(2);

    ui.alert('Готово!', 'Статистика создана', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString(), ui.ButtonSet.OK);
  }
}
```

### ABC-анализ по остаткам

```javascript
function createABCAnalysis() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ABC-анализ') || ss.insertSheet('ABC-анализ');
  sheet.clear();

  try {
    var products = getAllProducts({ visibility: VISIBILITY.IN_SALE });
    var productIds = products.map(function(p) { return p.product_id; });
    var stocks = getProductStocks(productIds);

    // Создаем массив с товарами и их остатками
    var productsWithStocks = [];

    products.forEach(function(product) {
      var stockData = stocks.find(function(s) { return s.product_id === product.product_id; });
      var totalStock = 0;

      if (stockData) {
        stockData.stocks.forEach(function(stock) {
          totalStock += (stock.present || 0);
        });
      }

      productsWithStocks.push({
        product_id: product.product_id,
        offer_id: product.offer_id,
        stock: totalStock
      });
    });

    // Сортируем по остаткам (по убыванию)
    productsWithStocks.sort(function(a, b) { return b.stock - a.stock; });

    // Вычисляем общий остаток
    var totalStock = productsWithStocks.reduce(function(sum, p) { return sum + p.stock; }, 0);

    // Классифицируем товары
    var cumulativeStock = 0;
    var data = productsWithStocks.map(function(product) {
      cumulativeStock += product.stock;
      var cumulativePercent = (cumulativeStock / totalStock * 100);

      var category = '';
      if (cumulativePercent <= 80) {
        category = 'A (Топ 80%)';
      } else if (cumulativePercent <= 95) {
        category = 'B (80-95%)';
      } else {
        category = 'C (95-100%)';
      }

      return [
        product.product_id,
        product.offer_id,
        product.stock,
        cumulativePercent.toFixed(2) + '%',
        category
      ];
    });

    // Заголовки
    var headers = ['Product ID', 'Offer ID', 'Остаток', 'Накопленный %', 'Категория ABC'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
      sheet.getRange(2, 3, data.length, 1).setNumberFormat('#,##0');
    }

    ui.alert('Готово!', 'ABC-анализ создан', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString(), ui.ButtonSet.OK);
  }
}
```

## Автоматизация

### Ежедневное обновление данных

```javascript
/**
 * Функция для ежедневного автоматического обновления
 * Настройте триггер: Triggers → Add Trigger → dailyUpdate → Time-driven → Day timer
 */
function dailyUpdate() {
  try {
    loadProductsWithStocksAndPrices();

    // Отправить email с уведомлением
    var email = Session.getActiveUser().getEmail();
    MailApp.sendEmail({
      to: email,
      subject: 'Ozon Seller - Данные обновлены',
      body: 'Данные о товарах, остатках и ценах успешно обновлены.\n\nВремя обновления: ' + new Date()
    });

  } catch (e) {
    // Отправить email с ошибкой
    var email = Session.getActiveUser().getEmail();
    MailApp.sendEmail({
      to: email,
      subject: 'Ozon Seller - Ошибка обновления',
      body: 'Произошла ошибка при обновлении данных:\n\n' + e.toString()
    });

    Logger.log('Ошибка ежедневного обновления: ' + e.toString());
  }
}
```

### Мониторинг критических остатков

```javascript
/**
 * Функция для мониторинга критически низких остатков
 * Настройте триггер: каждый час или каждые 6 часов
 */
function monitorCriticalStock() {
  var CRITICAL_THRESHOLD = 5;

  try {
    var products = getAllProducts({ visibility: VISIBILITY.IN_SALE });
    var productIds = products.map(function(p) { return p.product_id; });
    var stocks = getProductStocks(productIds);

    var criticalProducts = [];

    stocks.forEach(function(item) {
      var totalStock = 0;
      item.stocks.forEach(function(stock) {
        totalStock += (stock.present || 0);
      });

      if (totalStock > 0 && totalStock <= CRITICAL_THRESHOLD) {
        var product = products.find(function(p) { return p.product_id === item.product_id; });
        criticalProducts.push({
          offer_id: product ? product.offer_id : 'N/A',
          product_id: item.product_id,
          stock: totalStock
        });
      }
    });

    if (criticalProducts.length > 0) {
      var message = 'ВНИМАНИЕ! Обнаружены товары с критически низкими остатками:\n\n';

      criticalProducts.forEach(function(p) {
        message += 'Offer ID: ' + p.offer_id + ', Product ID: ' + p.product_id + ', Остаток: ' + p.stock + '\n';
      });

      var email = Session.getActiveUser().getEmail();
      MailApp.sendEmail({
        to: email,
        subject: '🔴 Ozon Seller - Критические остатки!',
        body: message
      });
    }

  } catch (e) {
    Logger.log('Ошибка мониторинга остатков: ' + e.toString());
  }
}
```

## Работа с ценами

### Массовое обновление цен (подготовка файла)

```javascript
/**
 * Создает шаблон для массового обновления цен
 * После заполнения используйте API Ozon для загрузки новых цен
 */
function createPriceUpdateTemplate() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Обновление цен') || ss.insertSheet('Обновление цен');
  sheet.clear();

  try {
    var products = getAllProducts({ visibility: VISIBILITY.IN_SALE });
    var productIds = products.map(function(p) { return p.product_id; });
    var prices = getProductPrices(productIds);

    var pricesMap = {};
    prices.forEach(function(item) {
      pricesMap[item.product_id] = item.price;
    });

    var headers = [
      'Product ID',
      'Offer ID',
      'Текущая цена',
      'Новая цена',
      'Старая цена (зачеркнутая)',
      'Комментарий'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

    var data = products.map(function(product) {
      var priceData = pricesMap[product.product_id] || {};
      var currentPrice = priceData.price || '';

      return [
        product.product_id,
        product.offer_id,
        currentPrice,
        '', // Новая цена - заполнить вручную
        '', // Старая цена - заполнить вручную
        ''  // Комментарий
      ];
    });

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);

      // Защитить первые 3 столбца от редактирования
      var protection = sheet.getRange(2, 1, data.length, 3).protect();
      protection.setDescription('Защита от случайного изменения');
      protection.setWarningOnly(true);

      // Выделить столбцы для заполнения
      sheet.getRange(2, 4, data.length, 2).setBackground('#fff3cd');
    }

    ui.alert('Готово!',
      'Шаблон создан. Заполните столбцы "Новая цена" и "Старая цена", затем используйте API Ozon для загрузки.',
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString(), ui.ButtonSet.OK);
  }
}
```

### Расчет оптимальной цены

```javascript
/**
 * Рассчитывает оптимальную цену на основе минимальной и рекомендованной
 */
function calculateOptimalPrice() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Оптимальные цены') || ss.insertSheet('Оптимальные цены');
  sheet.clear();

  // Коэффициент наценки (по умолчанию 20%)
  var MARKUP = 0.20;

  try {
    var products = getAllProducts({ visibility: VISIBILITY.IN_SALE });
    var productIds = products.map(function(p) { return p.product_id; });
    var prices = getProductPrices(productIds);

    var pricesMap = {};
    prices.forEach(function(item) {
      pricesMap[item.product_id] = item.price;
    });

    var headers = [
      'Product ID',
      'Offer ID',
      'Текущая цена',
      'Минимальная',
      'Рекомендованная',
      'Оптимальная (расчет)',
      'Выгода',
      'Рекомендация'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

    var data = products.map(function(product) {
      var priceData = pricesMap[product.product_id] || {};
      var currentPrice = parseFloat(priceData.price) || 0;
      var minPrice = parseFloat(priceData.min_price) || 0;
      var recommendedPrice = parseFloat(priceData.recommended_price) || 0;

      // Оптимальная цена = минимальная + 20% или рекомендованная (что меньше)
      var optimalPrice = Math.min(
        minPrice * (1 + MARKUP),
        recommendedPrice
      );

      var benefit = optimalPrice - currentPrice;
      var recommendation = benefit > 0 ? '⬆️ Повысить на ' + benefit.toFixed(2) + ' ₽' :
                          benefit < 0 ? '⬇️ Снизить на ' + Math.abs(benefit).toFixed(2) + ' ₽' :
                          '✅ Оптимальная';

      return [
        product.product_id,
        product.offer_id,
        currentPrice,
        minPrice,
        recommendedPrice,
        optimalPrice.toFixed(2),
        benefit.toFixed(2),
        recommendation
      ];
    });

    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
      sheet.getRange(2, 3, data.length, 4).setNumberFormat('#,##0.00 ₽');
    }

    ui.alert('Готово!', 'Расчет оптимальных цен завершен', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Ошибка', e.toString(), ui.ButtonSet.OK);
  }
}
```

## Добавление функций в меню

Чтобы добавить ваши кастомные функции в меню, обновите функцию `setupMenu()`:

```javascript
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
    .addSubMenu(ui.createMenu('📈 Отчеты')
      .addItem('Низкие остатки', 'createLowStockReport')
      .addItem('Маржинальность', 'createMarginReport')
      .addItem('Сравнение цен', 'comparePrices')
      .addItem('ABC-анализ', 'createABCAnalysis')
      .addItem('Статистика', 'createProductStatistics'))
    .addSeparator()
    .addSubMenu(ui.createMenu('💵 Работа с ценами')
      .addItem('Шаблон обновления цен', 'createPriceUpdateTemplate')
      .addItem('Расчет оптимальных цен', 'calculateOptimalPrice'))
    .addSeparator()
    .addItem('🗑️ Очистить лист', 'clearSheet')
    .addToUi();
}
```

---

Эти примеры можно адаптировать под ваши конкретные нужды. Не забудьте протестировать функции на небольшом наборе данных перед использованием на всем каталоге!
