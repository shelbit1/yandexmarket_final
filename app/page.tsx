'use client';

import { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

export default function Home() {
  const [token, setToken] = useState('ACMA:qYAokJ0Qp5OtWzlXTT4EqiDYvfoCWp4bgTSCAkXI:d2b30414');
  
  // Выбор завершенного месяца
  const currentDate = new Date();
  const [selectedMonth, setSelectedMonth] = useState(currentDate.getMonth()); // 0-11
  const [selectedYear, setSelectedYear] = useState(currentDate.getFullYear());
  
  // Вычисляем dateA и dateB на основе выбранного месяца
  const dateA = `${selectedYear}-${String(selectedMonth + 1).padStart(2, '0')}-01`;
  const lastDay = new Date(selectedYear, selectedMonth + 1, 0).getDate();
  const dateB = `${selectedYear}-${String(selectedMonth + 1).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
  
  const [campaignId, setCampaignId] = useState('');
  const [campaigns, setCampaigns] = useState<any[]>([]);
  const [isLoadingCampaigns, setIsLoadingCampaigns] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [showCostPriceModal, setShowCostPriceModal] = useState(false);
  const [products, setProducts] = useState<any[]>([]);
  const [costPrices, setCostPrices] = useState<Record<string, number>>({});
  const [fullReportProgress, setFullReportProgress] = useState<string>('');

  // Загружаем сохраненные данные себестоимости из localStorage при монтировании
  useEffect(() => {
    const savedCostPrices = localStorage.getItem('yandex_market_cost_prices');
    if (savedCostPrices) {
      try {
        const parsed = JSON.parse(savedCostPrices);
        setCostPrices(parsed);
        console.log('Загружены сохраненные данные себестоимости:', Object.keys(parsed).length, 'товаров');
      } catch (error) {
        console.error('Ошибка при загрузке данных себестоимости:', error);
      }
    }
  }, []);

  // Сохраняем данные себестоимости в localStorage при каждом изменении
  useEffect(() => {
    if (Object.keys(costPrices).length > 0) {
      localStorage.setItem('yandex_market_cost_prices', JSON.stringify(costPrices));
      console.log('Данные себестоимости сохранены:', Object.keys(costPrices).length, 'товаров');
    }
  }, [costPrices]);

  // Функция для загрузки списка магазинов
  const loadCampaigns = async () => {
    if (!token) {
      return;
    }

    setIsLoadingCampaigns(true);
    try {
      const response = await fetch('/api/campaigns', {
        method: 'GET',
        headers: {
          'x-api-key': token,
          'Content-Type': 'application/json',
        }
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`HTTP ${response.status}: ${errorText}`);
      }

      const data = await response.json();
      
      if (data.campaigns && data.campaigns.length > 0) {
        // Сохраняем все магазины
        const campaignsList = data.campaigns.map((campaign: any) => ({
          id: campaign.id,
          name: campaign.domain || `Магазин ${campaign.id}`,
          apiAvailability: campaign.apiAvailability,
          placementType: campaign.placementType,
          business: campaign.business
        }));
        
        setCampaigns(campaignsList);
        
        // Автоматически выбираем первый доступный магазин
        if (campaignsList.length > 0 && !campaignId) {
          setCampaignId(campaignsList[0].id.toString());
        }
      } else {
        setCampaigns([]);
        throw new Error('Не найдено ни одного магазина');
      }
    } catch (error: any) {
      console.error('Ошибка при получении кампаний:', error);
      setCampaigns([]);
      
      // Показываем сообщение об ошибке
      alert(`Ошибка загрузки магазинов: ${error.message}`);
    } finally {
      setIsLoadingCampaigns(false);
    }
  };

  // Функция для очистки названия листа от запрещенных символов Excel
  const sanitizeSheetName = (name: string): string => {
    // Excel не разрешает символы: : \ / ? * [ ]
    // Также ограничение - максимум 31 символ
    return name
      .replace(/[:\\/\?*\[\]]/g, '-') // Заменяем запрещенные символы на дефис
      .substring(0, 31); // Обрезаем до 31 символа
  };

  // Универсальный хелпер для извлечения URL файла из result
      const resolveReportFileUrl = (resultObj: any): string | null => {
        if (!resultObj) return null;
        const candidates: any[] = [];
        if (resultObj.file && typeof resultObj.file === 'object') {
          candidates.push(
            resultObj.file.url,
            resultObj.file.href,
            resultObj.file.downloadUrl,
            resultObj.file.fileUrl,
            resultObj.file.mdsUrl
          );
        }
    if (typeof resultObj.file === 'string') candidates.push(resultObj.file);
        candidates.push(
          resultObj.url,
          resultObj.fileUrl,
          resultObj.downloadUrl,
          resultObj.href,
          resultObj.mdsUrl
        );
        if (Array.isArray(resultObj.files)) {
          for (const f of resultObj.files) {
        if (!f) continue;
        if (typeof f === 'string') candidates.push(f);
        else if (typeof f === 'object') {
              candidates.push(f.url, f.href, f.downloadUrl, f.fileUrl, f.mdsUrl);
        }
      }
    }
    const isUrl = (v: any) => typeof v === 'string' && /^https?:\/\//.test(v);
    const found = candidates.find(isUrl);
    return found || null;
  };

  // Функция для кнопки "Себестоимость"
  const downloadCostPrice = async () => {
    if (!token || !campaignId) {
      alert('Заполните токен и выберите магазин');
      return;
    }

    setIsLoading(true);
    try {
      console.log('Загрузка списка товаров для себестоимости...');

      let allOffers: any[] = [];
      let pageToken = null;
      let attempts = 0;
      const maxAttempts = 100;

      // Получаем все товары с пагинацией
      do {
        const requestBody: any = {
          limit: 200
        };

        if (pageToken) {
          requestBody.page_token = pageToken;
        }

        const response = await fetch('/api/offers', {
          method: 'POST',
          headers: {
            'x-api-key': token,
            'x-campaign-id': campaignId,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(requestBody)
        });

        if (!response.ok) {
          console.error('Ошибка получения товаров:', response.status);
          alert('Ошибка при загрузке товаров. Проверьте доступы API-ключа.');
          setIsLoading(false);
          return;
        }

        const data = await response.json();
        
        if (data.result?.offers && data.result.offers.length > 0) {
          allOffers = allOffers.concat(data.result.offers);
        }

        pageToken = data.result?.paging?.nextPageToken || null;
        attempts++;

        if (pageToken && attempts < maxAttempts) {
          await new Promise(resolve => setTimeout(resolve, 100));
        }

      } while (pageToken && attempts < maxAttempts);

      console.log(`Загружено товаров: ${allOffers.length}`);
      
      // Сохраняем товары и открываем модальное окно
      setProducts(allOffers);
      setShowCostPriceModal(true);
      
    } catch (error) {
      console.error('Ошибка при загрузке товаров:', error);
      alert('Произошла ошибка при загрузке товаров');
    } finally {
      setIsLoading(false);
    }
  };

  // Функция для обновления себестоимости товара
  const updateCostPrice = (offerId: string, value: string) => {
    const numValue = parseFloat(value);
    if (!isNaN(numValue) && numValue >= 0) {
      setCostPrices(prev => ({
        ...prev,
        [offerId]: numValue
      }));
    } else if (value === '') {
      setCostPrices(prev => {
        const newPrices = { ...prev };
        delete newPrices[offerId];
        return newPrices;
      });
    }
  };

  // Функция для очистки всех сохраненных данных себестоимости
  const clearAllCostPrices = () => {
    if (confirm('Вы уверены, что хотите удалить все сохраненные данные себестоимости? Это действие нельзя отменить.')) {
      setCostPrices({});
      localStorage.removeItem('yandex_market_cost_prices');
      alert('Все данные себестоимости удалены');
    }
  };

  // Функция для импорта данных о себестоимости из Excel
  const importCostPricesFromExcel = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Читаем первый лист
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Конвертируем в JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
        
        // Находим строку с заголовками (ищем строку, содержащую "Артикул" и "Себестоимость")
        let headerRowIndex = -1;
        for (let i = 0; i < jsonData.length; i++) {
          const row = jsonData[i];
          if (row.some((cell: any) => 
            typeof cell === 'string' && 
            (cell.includes('Артикул') || cell.includes('SKU'))
          )) {
            headerRowIndex = i;
            break;
          }
        }

        if (headerRowIndex === -1) {
          alert('Не найдена строка с заголовками. Убедитесь, что файл содержит колонки "Артикул" и "Себестоимость".');
          return;
        }

        // Находим индексы нужных колонок
        const headers = jsonData[headerRowIndex];
        const skuIndex = headers.findIndex((h: any) => 
          typeof h === 'string' && (h.includes('Артикул') || h.includes('SKU'))
        );
        const costPriceIndex = headers.findIndex((h: any) => 
          typeof h === 'string' && h.includes('Себестоимость')
        );

        if (skuIndex === -1 || costPriceIndex === -1) {
          alert('Не найдены необходимые колонки. Файл должен содержать "Артикул (SKU)" и "Себестоимость".');
          return;
        }

        // Парсим данные
        const importedPrices: Record<string, number> = {};
        let importedCount = 0;

        for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
          const row = jsonData[i];
          if (!row || row.length === 0) continue;

          const sku = row[skuIndex];
          const costPrice = row[costPriceIndex];

          if (sku && costPrice !== undefined && costPrice !== null && costPrice !== '') {
            const parsedPrice = typeof costPrice === 'number' 
              ? costPrice 
              : parseFloat(String(costPrice).replace(/[^\d.-]/g, ''));
            
            if (!isNaN(parsedPrice) && parsedPrice >= 0) {
              importedPrices[String(sku).trim()] = parsedPrice;
              importedCount++;
            }
          }
        }

        if (importedCount === 0) {
          alert('Не найдено данных для импорта. Проверьте формат файла.');
          return;
        }

        // Обновляем состояние
        setCostPrices(prev => ({
          ...prev,
          ...importedPrices
        }));

        alert(`✅ Импорт завершен!\n\nЗагружено цен: ${importedCount}\nОбщее количество товаров с себестоимостью: ${Object.keys({ ...costPrices, ...importedPrices }).length}`);
        
      } catch (error) {
        console.error('Ошибка при импорте:', error);
        alert('Произошла ошибка при чтении файла. Убедитесь, что файл имеет правильный формат Excel.');
      }
    };

    reader.readAsArrayBuffer(file);
    
    // Сбрасываем значение input, чтобы можно было загрузить тот же файл повторно
    event.target.value = '';
  };

  // Функция для экспорта данных с себестоимостью в Excel
  const exportCostPricesToExcel = () => {
      const workbook = XLSX.utils.book_new();
      
    const headers = [
      'Артикул (SKU)',
      'Себестоимость',
      'Цена продажи',
      'Валюта',
      'Маржа (₽)',
      'Маржа (%)',
      'Статус'
    ];

    const rows = [
      [`Себестоимость товаров | Campaign ID: ${campaignId} | Создано: ${new Date().toLocaleString('ru-RU')}`],
      [`Всего товаров: ${products.length} | Заполнено: ${Object.keys(costPrices).length}`],
      [''],
      headers
    ];

    const statusNames: Record<string, string> = {
      'PUBLISHED': 'Готов к продаже',
      'CHECKING': 'На проверке',
      'DISABLED_BY_PARTNER': 'Скрыт вами',
      'REJECTED_BY_MARKET': 'Отклонен',
      'DISABLED_AUTOMATICALLY': 'Исправьте ошибки',
      'CREATING_CARD': 'Создается карточка',
      'NO_CARD': 'Нужна карточка',
      'NO_STOCKS': 'Нет на складе',
      'ARCHIVED': 'В архиве'
    };

    products.forEach((product: any) => {
      const costPrice = costPrices[product.offerId] || 0;
      const salePrice = product.campaignPrice?.value || product.basicPrice?.value || 0;
      const currency = product.campaignPrice?.currencyId || product.basicPrice?.currencyId || 'RUR';
      const margin = salePrice - costPrice;
      const marginPercent = costPrice > 0 ? ((margin / costPrice) * 100).toFixed(2) : '0';

      rows.push([
        product.offerId || '',
        costPrice || '',
        salePrice || '',
        currency,
        margin.toFixed(2),
        marginPercent + '%',
        statusNames[product.status] || product.status || ''
              ]);
            });

    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    
    worksheet['!cols'] = [
      { wch: 30 }, // Артикул
      { wch: 15 }, // Себестоимость
      { wch: 15 }, // Цена продажи
      { wch: 10 }, // Валюта
      { wch: 15 }, // Маржа ₽
      { wch: 12 }, // Маржа %
      { wch: 20 }  // Статус
    ];
    
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Себестоимость');

      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      link.href = url;
    link.download = `cost_prices_${campaignId}_${new Date().toISOString().split('T')[0]}.xlsx`;
      link.click();
      URL.revokeObjectURL(url);
    
    alert(`Файл с себестоимостью экспортирован! Товаров: ${products.length}, заполнено: ${Object.keys(costPrices).length}`);
  };

  // Функция для кнопки "Остатки" - загрузка информации об остатках товаров
  const downloadStocks = async () => {
    setIsLoading(true);
    try {
      if (!token || !campaignId) {
        alert('Заполните токен и выберите магазин');
        setIsLoading(false);
        return;
      }

      console.log('Загрузка информации об остатках...');

      let allStocks: any[] = [];
      let pageToken = null;
      let attempts = 0;
      const maxAttempts = 100;

      // Получаем все остатки с пагинацией
      do {
        const requestBody: any = {
          limit: 200,
          withTurnover: true // Получаем информацию об оборачиваемости
        };

        if (pageToken) {
          requestBody.page_token = pageToken;
        }

        console.log(`Запрос остатков (попытка ${attempts + 1})...`);

        const response = await fetch('/api/stocks', {
        method: 'POST',
        headers: {
          'x-api-key': token,
            'x-campaign-id': campaignId,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(requestBody)
        });

        if (!response.ok) {
          console.error('Ошибка получения остатков:', response.status);
          const errorText = await response.text();
          console.error('Детали ошибки:', errorText);
          alert('Ошибка при загрузке остатков. Проверьте доступы API-ключа.');
          setIsLoading(false);
        return;
      }

        const data = await response.json();
        
        if (data.result?.warehouses && data.result.warehouses.length > 0) {
          // Собираем все остатки со всех складов
          data.result.warehouses.forEach((warehouse: any) => {
            if (warehouse.offers && warehouse.offers.length > 0) {
              warehouse.offers.forEach((offer: any) => {
                allStocks.push({
                  ...offer,
                  warehouseId: warehouse.warehouseId
                });
              });
            }
          });
          console.log(`Получено остатков, всего: ${allStocks.length}`);
        }

        pageToken = data.result?.paging?.nextPageToken || null;
        attempts++;

        if (pageToken && attempts < maxAttempts) {
          await new Promise(resolve => setTimeout(resolve, 100));
        }

      } while (pageToken && attempts < maxAttempts);

      console.log(`Итого получено остатков: ${allStocks.length}`);

      // Создаем Excel файл
      const workbook = XLSX.utils.book_new();
      
      const headers = [
        'Артикул (SKU)',
        'Склад ID',
        'Доступно (AVAILABLE)',
        'Годный (FIT)',
        'Заморожено (FREEZE)',
        'Карантин (QUARANTINE)',
        'Брак (DEFECT)',
        'Просрочен (EXPIRED)',
        'Утилизация (UTILIZATION)',
        'Оборачиваемость',
        'Дней оборачиваемости',
        'Дата обновления'
      ];

      const rows = [
        [`Остатки товаров | Campaign ID: ${campaignId} | Создано: ${new Date().toLocaleString('ru-RU')}`],
        [`Всего записей: ${allStocks.length}`],
        [''],
        headers
      ];

      // Словарь для расшифровки оборачиваемости
      const turnoverNames: Record<string, string> = {
        'LOW': 'Низкая (≥120 дней)',
        'ALMOST_LOW': 'Почти низкая (100-119 дней)',
        'HIGH': 'Высокая (45-99 дней)',
        'VERY_HIGH': 'Очень высокая (<45 дней)',
        'NO_SALES': 'Нет продаж',
        'FREE_STORE': 'Бесплатное хранение'
      };

      allStocks.forEach((stock: any) => {
        // Извлекаем остатки по типам
        const stocksByType: Record<string, number> = {};
        if (stock.stocks && stock.stocks.length > 0) {
          stock.stocks.forEach((s: any) => {
            stocksByType[s.type] = s.count;
          });
        }

        rows.push([
          stock.offerId || '',
          stock.warehouseId || '',
          stocksByType['AVAILABLE'] || 0,
          stocksByType['FIT'] || 0,
          stocksByType['FREEZE'] || 0,
          stocksByType['QUARANTINE'] || 0,
          stocksByType['DEFECT'] || 0,
          stocksByType['EXPIRED'] || 0,
          stocksByType['UTILIZATION'] || 0,
          stock.turnoverSummary ? turnoverNames[stock.turnoverSummary.turnover] || stock.turnoverSummary.turnover : '',
          stock.turnoverSummary?.turnoverDays || '',
          stock.updatedAt || ''
        ]);
      });

      const worksheet = XLSX.utils.aoa_to_sheet(rows);
      
      worksheet['!cols'] = [
        { wch: 30 }, // Артикул
        { wch: 12 }, // Склад ID
        { wch: 18 }, // AVAILABLE
        { wch: 15 }, // FIT
        { wch: 18 }, // FREEZE
        { wch: 18 }, // QUARANTINE
        { wch: 15 }, // DEFECT
        { wch: 18 }, // EXPIRED
        { wch: 20 }, // UTILIZATION
        { wch: 30 }, // Оборачиваемость
        { wch: 20 }, // Дней
        { wch: 25 }  // Дата обновления
      ];
      
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Остатки');

      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
          const link = document.createElement('a');
          const url = URL.createObjectURL(blob);
          link.href = url;
      link.download = `stocks_${campaignId}_${new Date().toISOString().split('T')[0]}.xlsx`;
          link.click();
          URL.revokeObjectURL(url);
      
      alert(`Остатки товаров скачаны! Записей: ${allStocks.length}`);
      
    } catch (error) {
      console.error('Ошибка при загрузке остатков:', error);
      alert('Произошла ошибка при загрузке остатков: ' + (error as Error).message);
    } finally {
      setIsLoading(false);
    }
  };

  // Функция для кнопки "Остатки2" - загрузка остатков в формате для обновления
  const downloadStocks2 = async () => {
    setIsLoading(true);
    try {
      if (!token || !campaignId) {
        alert('Заполните токен и выберите магазин');
        setIsLoading(false);
        return;
      }

      console.log('Загрузка информации об остатках для обновления...');

      let allStocks: any[] = [];
      let pageToken = null;
      let attempts = 0;
      const maxAttempts = 100;

      // Получаем все остатки с пагинацией
      do {
        const requestBody: any = {
          limit: 200,
          withTurnover: false // Для обновления не нужна оборачиваемость
        };

        if (pageToken) {
          requestBody.page_token = pageToken;
        }

        const response = await fetch('/api/stocks', {
          method: 'POST',
          headers: {
            'x-api-key': token,
            'x-campaign-id': campaignId,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(requestBody)
        });

        if (!response.ok) {
          console.error('Ошибка получения остатков:', response.status);
          alert('Ошибка при загрузке остатков. Проверьте доступы API-ключа.');
          setIsLoading(false);
          return;
        }

        const data = await response.json();
        
        if (data.result?.warehouses && data.result.warehouses.length > 0) {
          data.result.warehouses.forEach((warehouse: any) => {
            if (warehouse.offers && warehouse.offers.length > 0) {
              warehouse.offers.forEach((offer: any) => {
                allStocks.push({
                  ...offer,
                  warehouseId: warehouse.warehouseId
                });
              });
            }
          });
        }

        pageToken = data.result?.paging?.nextPageToken || null;
        attempts++;

        if (pageToken && attempts < maxAttempts) {
          await new Promise(resolve => setTimeout(resolve, 100));
        }

      } while (pageToken && attempts < maxAttempts);

      console.log(`Итого получено остатков: ${allStocks.length}`);

      // Создаем Excel файл в формате для обновления
      const workbook = XLSX.utils.book_new();
      
      const headers = [
        'SKU (Артикул)',
        'Склад ID',
        'Количество (count)',
        'Дата обновления (updatedAt)',
        'Примечание'
      ];

      const rows = [
        [`Остатки для обновления | Campaign ID: ${campaignId} | Создано: ${new Date().toLocaleString('ru-RU')}`],
        ['⚠️ Этот файл предназначен для редактирования и последующей отправки остатков на Яндекс.Маркет'],
        ['📝 Заполните колонку "Количество" нужными значениями и используйте для обновления через API'],
        [''],
        headers
      ];

      // Группируем остатки по SKU и складу
      const stocksMap = new Map<string, any>();
      
      allStocks.forEach((stock: any) => {
        const key = `${stock.offerId}_${stock.warehouseId}`;
        
        if (!stocksMap.has(key)) {
          // Извлекаем доступное количество (AVAILABLE)
          let availableCount = 0;
          if (stock.stocks && stock.stocks.length > 0) {
            const availableStock = stock.stocks.find((s: any) => s.type === 'AVAILABLE');
            if (availableStock) {
              availableCount = availableStock.count;
            }
          }
          
          stocksMap.set(key, {
            sku: stock.offerId,
            warehouseId: stock.warehouseId,
            count: availableCount,
            updatedAt: stock.updatedAt || new Date().toISOString(),
            note: 'Текущее количество доступных товаров'
          });
        }
      });

      // Добавляем строки
      stocksMap.forEach((stockData) => {
        rows.push([
          stockData.sku,
          stockData.warehouseId,
          stockData.count,
          stockData.updatedAt,
          stockData.note
        ]);
      });

      // Добавляем инструкцию внизу
      rows.push(['']);
      rows.push(['ИНСТРУКЦИЯ ПО ИСПОЛЬЗОВАНИЮ:']);
      rows.push(['1. Отредактируйте значения в колонке "Количество (count)"']);
      rows.push(['2. Сохраните файл']);
      rows.push(['3. Используйте API метод PUT для отправки обновленных остатков']);
      rows.push(['4. Формат для API: { "skus": [{ "sku": "...", "items": [{ "count": 123, "updatedAt": "..." }] }] }']);

      const worksheet = XLSX.utils.aoa_to_sheet(rows);
      
      worksheet['!cols'] = [
        { wch: 30 }, // SKU
        { wch: 12 }, // Склад ID
        { wch: 20 }, // Количество
        { wch: 25 }, // Дата обновления
        { wch: 50 }  // Примечание
      ];
      
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Остатки для обновления');

      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      link.href = url;
      link.download = `stocks_for_update_${campaignId}_${new Date().toISOString().split('T')[0]}.xlsx`;
      link.click();
      URL.revokeObjectURL(url);
      
      alert(`Файл остатков для обновления скачан! Записей: ${stocksMap.size}\n\nОтредактируйте количества и используйте для обновления через API.`);
      
    } catch (error) {
      console.error('Ошибка при загрузке остатков:', error);
      alert('Произошла ошибка при загрузке остатков: ' + (error as Error).message);
    } finally {
      setIsLoading(false);
    }
  };

  // Функция для кнопки "Мои товары" - загрузка списка всех артикулов магазина
  const downloadMyProducts = async () => {
    setIsLoading(true);
    try {
      if (!token || !campaignId) {
        alert('Заполните токен и выберите магазин');
        setIsLoading(false);
        return;
      }

      console.log('Загрузка списка товаров магазина...');

      let allOffers: any[] = [];
      let pageToken = null;
      let attempts = 0;
      const maxAttempts = 100; // максимум 100 страниц

      // Получаем все товары с пагинацией
      do {
        const requestBody: any = {
          limit: 200 // максимальное количество за запрос
        };

        if (pageToken) {
          requestBody.page_token = pageToken;
        }

        console.log(`Запрос товаров (попытка ${attempts + 1})...`);

        const response = await fetch('/api/offers', {
        method: 'POST',
        headers: {
          'x-api-key': token,
            'x-campaign-id': campaignId,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(requestBody)
        });

        if (!response.ok) {
          console.error('Ошибка получения товаров:', response.status);
          const errorText = await response.text();
          console.error('Детали ошибки:', errorText);
          alert('Ошибка при загрузке товаров. Проверьте доступы API-ключа.');
          setIsLoading(false);
        return;
      }

        const data = await response.json();
        
        if (data.result?.offers && data.result.offers.length > 0) {
          allOffers = allOffers.concat(data.result.offers);
          console.log(`Получено ${data.result.offers.length} товаров, всего: ${allOffers.length}`);
        }

        // Проверяем есть ли следующая страница
        pageToken = data.result?.paging?.nextPageToken || null;
        attempts++;

        // Добавляем задержку между запросами
        if (pageToken && attempts < maxAttempts) {
          await new Promise(resolve => setTimeout(resolve, 100));
        }

      } while (pageToken && attempts < maxAttempts);

      console.log(`Итого получено товаров: ${allOffers.length}`);

      // Создаем Excel файл
      const workbook = XLSX.utils.book_new();
      
      // Подготавливаем данные для Excel
      const headers = [
        'Артикул (SKU)',
        'Статус',
        'Доступен для продажи',
        'Базовая цена',
        'Валюта',
        'Зачеркнутая цена',
        'Цена магазина',
        'Валюта магазина',
        'НДС',
        'Минимальное количество',
        'Шаг количества',
        'Дата обновления цены',
        'Ошибки',
        'Предупреждения',
        'Себестоимость'
      ];

      const rows = [
        [`Список товаров магазина | Campaign ID: ${campaignId} | Создано: ${new Date().toLocaleString('ru-RU')}`],
        [`Всего товаров: ${allOffers.length}`],
        [''],
        headers
      ];

      // Словарь статусов для понятного отображения
      const statusNames: Record<string, string> = {
        'PUBLISHED': 'Готов к продаже',
        'CHECKING': 'На проверке',
        'DISABLED_BY_PARTNER': 'Скрыт вами',
        'REJECTED_BY_MARKET': 'Отклонен',
        'DISABLED_AUTOMATICALLY': 'Исправьте ошибки',
        'CREATING_CARD': 'Создается карточка',
        'NO_CARD': 'Нужна карточка',
        'NO_STOCKS': 'Нет на складе',
        'ARCHIVED': 'В архиве'
      };

      allOffers.forEach((offer: any) => {
        const basicPrice = offer.basicPrice;
        const campaignPrice = offer.campaignPrice;
        const quantum = offer.quantum;
        
        // Собираем ошибки в одну строку
        const errors = offer.errors && offer.errors.length > 0
          ? offer.errors.map((e: any) => `${e.message}${e.comment ? ': ' + e.comment : ''}`).join('; ')
          : '';
        
        // Собираем предупреждения в одну строку
        const warnings = offer.warnings && offer.warnings.length > 0
          ? offer.warnings.map((w: any) => `${w.message}${w.comment ? ': ' + w.comment : ''}`).join('; ')
          : '';

        rows.push([
          offer.offerId || '',
          statusNames[offer.status] || offer.status || '',
          offer.available !== undefined ? (offer.available ? 'Да' : 'Нет') : '',
          basicPrice?.value || '',
          basicPrice?.currencyId || '',
          basicPrice?.discountBase || '',
          campaignPrice?.value || '',
          campaignPrice?.currencyId || '',
          campaignPrice?.vat || '',
          quantum?.minQuantity || '',
          quantum?.stepQuantity || '',
          basicPrice?.updatedAt || campaignPrice?.updatedAt || '',
          errors,
          warnings,
          costPrices[offer.offerId] || 0
        ]);
            });

      // Создаем лист
      const worksheet = XLSX.utils.aoa_to_sheet(rows);
      
      // Устанавливаем ширину колонок
      worksheet['!cols'] = [
        { wch: 30 }, // Артикул
        { wch: 20 }, // Статус
        { wch: 15 }, // Доступен
        { wch: 12 }, // Базовая цена
        { wch: 8 },  // Валюта
        { wch: 15 }, // Зачеркнутая
        { wch: 12 }, // Цена магазина
        { wch: 8 },  // Валюта магазина
        { wch: 8 },  // НДС
        { wch: 15 }, // Мин количество
        { wch: 15 }, // Шаг
        { wch: 20 }, // Дата обновления
        { wch: 50 }, // Ошибки
        { wch: 50 }, // Предупреждения
        { wch: 15 }  // Себестоимость
      ];
      
      // Добавляем лист в книгу
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Мои товары');

      // Генерируем и скачиваем файл
      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
          const link = document.createElement('a');
          const url = URL.createObjectURL(blob);
          link.href = url;
      link.download = `my_products_${campaignId}_${new Date().toISOString().split('T')[0]}.xlsx`;
          link.click();
          URL.revokeObjectURL(url);
      
      alert(`Список товаров скачан! Всего товаров: ${allOffers.length}`);
      
    } catch (error) {
      console.error('Ошибка при загрузке товаров:', error);
      alert('Произошла ошибка при загрузке товаров: ' + (error as Error).message);
    } finally {
      setIsLoading(false);
    }
  };

  // Функция для кнопки "Тест2" - скачивание отчета по стоимости услуг с данными
  const downloadTest2Excel = async () => {
    setIsLoading(true);
    try {
      if (!token || !dateA || !dateB) {
        alert('Заполните токен и выберите даты');
        setIsLoading(false);
        return;
      }

      const businessId = campaigns[0]?.business?.id;
      if (!businessId) {
        alert('Не найден businessId. Нажмите «Обновить список» и выберите магазин.');
        setIsLoading(false);
        return;
      }

      console.log(`Тестовая загрузка отчета по стоимости услуг за период ${dateA} - ${dateB}`);

      // 1) Генерация отчета в формате JSON
      const generateRes = await fetch('/api/services-cost', {
        method: 'POST',
        headers: {
          'x-api-key': token,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          businessId,
          dateFrom: dateA,
          dateTo: dateB,
          campaignIds: campaignId ? [Number(campaignId)] : undefined,
          format: 'JSON',
          language: 'RU'
        })
      });

      if (!generateRes.ok) {
        const txt = await generateRes.text();
        console.error('Ошибка генерации:', txt);
        alert('Ошибка запуска генерации тестового отчета по стоимости услуг');
        setIsLoading(false);
        return;
      }

      const genData = await generateRes.json();
      const reportId = genData?.result?.reportId;
      if (!reportId) {
        alert('ReportId не получен');
        setIsLoading(false);
        return;
      }

      console.log(`Отчет по стоимости услуг запущен, reportId: ${reportId}`);

      // 2) Ожидание готовности отчета
      let fileUrl = null;
      for (let i = 0; i < 30; i++) {
        await new Promise(res => setTimeout(res, 3000));
        const statusRes = await fetch(`/api/report-status?reportId=${reportId}`, {
          headers: { 'x-api-key': token }
        });
        if (!statusRes.ok) continue;
        const statusData = await statusRes.json();
        const status = statusData?.result?.status;
        
        console.log(`Статус отчета по стоимости услуг (попытка ${i + 1}): ${status}`);
        
        if (status === 'DONE') {
          fileUrl = resolveReportFileUrl(statusData?.result);
          break;
        } else if (status === 'FAILED') {
          alert('Генерация отчета завершилась с ошибкой');
          setIsLoading(false);
          return;
        }
      }

      if (!fileUrl) {
        alert('Не удалось получить ссылку на файл');
        setIsLoading(false);
        return;
      }

      // 3) Скачивание ZIP архива с JSON файлами
      const downloadRes = await fetch(`/api/download-report?url=${encodeURIComponent(fileUrl)}`, {
        headers: { 'x-api-key': token }
      });
      
      if (!downloadRes.ok) {
        alert('Не удалось скачать файл');
        setIsLoading(false);
        return;
      }

      // 4) Обработка ZIP архива и создание Excel
      const arrayBuffer = await downloadRes.arrayBuffer();
      
      // Используем библиотеку jszip для распаковки
      const JSZip = (await import('jszip')).default;
      const zip = await JSZip.loadAsync(arrayBuffer);
      
      // Создаем Excel книгу
      const workbook = XLSX.utils.book_new();
      
      // Определение всех 27 листов по стоимости услуг
      const sheetsToProcess = [
        { fileName: 'placement', ru: 'Размещение товаров' },
        { fileName: 'warehouse_processing', ru: 'Складская обработка' },
        { fileName: 'goods_acceptance', ru: 'Приемка поставки' },
        { fileName: 'loyalty_and_reviews', ru: 'Лояльность и отзывы' },
        { fileName: 'boost', ru: 'Буст продаж (продажи)' },
        { fileName: 'installment_plan', ru: 'Рассрочка' },
        { fileName: 'shelf', ru: 'Полки' },
        { fileName: 'cpm-boost', ru: 'Буст продаж (показы)' },
        { fileName: 'banners', ru: 'Баннеры' },
        { fileName: 'pushes', ru: 'Пуш-уведомления' },
        { fileName: 'popups', ru: 'Поп-ап уведомления' },
        { fileName: 'delivery', ru: 'Доставка покупателю' },
        { fileName: 'express_delivery', ru: 'Экспресс-доставка' },
        { fileName: 'delivery_from_abroad', ru: 'Доставка из-за рубежа' },
        { fileName: 'payment_accepting', ru: 'Приём платежа' },
        { fileName: 'payment_transfer', ru: 'Перевод платежа' },
        { fileName: 'paid_storage_before_31-05-22', ru: 'Хранение до 31.05.22' },
        { fileName: 'paid_storage_after_01-06-22', ru: 'Хранение с 01.06.22' },
        { fileName: 'delivery_via_transit_warehouse', ru: 'Транзитный склад' },
        { fileName: 'reception_of_surplus', ru: 'Приём излишков' },
        { fileName: 'product_marking', ru: 'Маркировка' },
        { fileName: 'export_from_warehouse', ru: 'Вывоз со склада' },
        { fileName: 'intake_logistics', ru: 'Забор заказов' },
        { fileName: 'order_processing', ru: 'Обработка в СЦ/ПВЗ' },
        { fileName: 'order_processing_on_warehouse', ru: 'Обработка на складе' },
        { fileName: 'storage_of_returns', ru: 'Хранение возвратов' },
        { fileName: 'expropriation', ru: 'Вознаграждение' },
        { fileName: 'utilization', ru: 'Утилизация' },
        { fileName: 'extended_service_access', ru: 'Расширенный доступ' },
        { fileName: 'personal_manager', ru: 'Персональный менеджер' }
      ];

      let sheetsCreated = 0;

      // Обрабатываем каждый JSON файл из архива
      for (const sheetInfo of sheetsToProcess) {
        const jsonFileName = `${sheetInfo.fileName}.json`;
        const jsonFile = zip.file(jsonFileName);
        
        let sheetData: any[][] = [];
        
        if (!jsonFile) {
          console.log(`Файл ${jsonFileName} не найден в архиве - создаем пустой лист`);
          // Создаем пустой лист с информационным сообщением
          sheetData = [['Данные отсутствуют']];
        } else {
          const jsonContent = await jsonFile.async('string');
          const jsonData = JSON.parse(jsonContent);
          
          if (jsonData.rows && jsonData.rows.length > 0) {
            // Получаем заголовки из первой строки данных
            const headers = Object.keys(jsonData.rows[0]);
            
            // Создаем массив для листа с заголовками
            sheetData = [headers];
            
            // Добавляем данные
            jsonData.rows.forEach((row: any) => {
              const rowData = headers.map(header => row[header] || '');
              sheetData.push(rowData);
            });
            
            console.log(`Лист ${sheetInfo.ru}: ${jsonData.rows.length} строк данных`);
          } else {
            console.log(`Файл ${jsonFileName} не содержит данных - создаем пустой лист`);
            // Создаем пустой лист с информационным сообщением
            sheetData = [['Нет данных за выбранный период']];
          }
        }
        
        // Создаем лист (даже если он пустой)
        const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
        
        // Добавляем лист в книгу (очищаем название от запрещенных символов и обрезаем до 31 символа)
        const sheetName = sanitizeSheetName(sheetInfo.ru);
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        
        sheetsCreated++;
        console.log(`Добавлен лист: ${sheetName}`);
      }

      // 5) Генерация и скачивание Excel файла
      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      link.href = url;
      link.download = `services_cost_report_${dateA}_${dateB}.xlsx`;
      link.click();
      URL.revokeObjectURL(url);
      
      alert(`Отчет по стоимости услуг скачан! Листов: ${sheetsCreated}`);
      
    } catch (error) {
      console.error('Ошибка при создании отчета по стоимости услуг:', error);
      alert('Произошла ошибка: ' + (error as Error).message);
    } finally {
      setIsLoading(false);
    }
  };

  // Функция для кнопки "Полный отчет" - собирает все данные в один Excel файл
  const downloadFullReport = async () => {
    if (!token || !campaignId) {
      alert('Заполните токен и выберите магазин');
      return;
    }

    setIsLoading(true);
    setFullReportProgress('Начинаем формирование полного отчета...');
    
    try {
      const workbook = XLSX.utils.book_new();
      let sheetsCreated = 0;

      // 1. Загружаем список товаров
      setFullReportProgress('📦 Загрузка списка товаров...');
      console.log('Загрузка списка товаров для полного отчета...');
      
      let allOffers: any[] = [];
      let pageToken = null;
      let attempts = 0;
      const maxAttempts = 100;

      do {
        const requestBody: any = { limit: 200 };
        if (pageToken) requestBody.page_token = pageToken;

        const response = await fetch('/api/offers', {
          method: 'POST',
          headers: {
            'x-api-key': token,
            'x-campaign-id': campaignId,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(requestBody)
        });

        if (!response.ok) {
          console.error('Ошибка получения товаров:', response.status);
          break;
        }

        const data = await response.json();
        if (data.result?.offers && data.result.offers.length > 0) {
          allOffers = allOffers.concat(data.result.offers);
        }

        pageToken = data.result?.paging?.nextPageToken || null;
        attempts++;

        if (pageToken && attempts < maxAttempts) {
          await new Promise(resolve => setTimeout(resolve, 100));
        }
      } while (pageToken && attempts < maxAttempts);

      console.log(`Загружено товаров: ${allOffers.length}`);

      // Создаем лист "Аналитика" с артикулами
      const analyticsRows = [
        [''], // A1 пустая
        ['Артикул (SKU)'], // A2
        ['Период'] // A3
      ];
      
      // Добавляем все артикулы начиная с A4
      allOffers.forEach((offer: any) => {
        analyticsRows.push([offer.offerId || '']);
      });
      
      const analyticsSheet = XLSX.utils.aoa_to_sheet(analyticsRows);
      analyticsSheet['!cols'] = [{ wch: 30 }]; // Ширина колонки A
      XLSX.utils.book_append_sheet(workbook, analyticsSheet, 'Аналитика');
      sheetsCreated++;

      // Создаем лист "Мои товары"
      const statusNames: Record<string, string> = {
        'PUBLISHED': 'Готов к продаже',
        'CHECKING': 'На проверке',
        'DISABLED_BY_PARTNER': 'Скрыт вами',
        'REJECTED_BY_MARKET': 'Отклонен',
        'DISABLED_AUTOMATICALLY': 'Исправьте ошибки',
        'CREATING_CARD': 'Создается карточка',
        'NO_CARD': 'Нужна карточка',
        'NO_STOCKS': 'Нет на складе',
        'ARCHIVED': 'В архиве'
      };

      const productHeaders = [
        'Артикул (SKU)', 'Статус', 'Доступен для продажи', 'Базовая цена', 'Валюта',
        'Зачеркнутая цена', 'Цена магазина', 'Валюта магазина', 'НДС',
        'Минимальное количество', 'Шаг количества', 'Дата обновления цены',
        'Ошибки', 'Предупреждения', 'Себестоимость'
      ];

      const productRows = [
        [`Список товаров магазина | Campaign ID: ${campaignId} | Создано: ${new Date().toLocaleString('ru-RU')}`],
        [`Всего товаров: ${allOffers.length}`],
        [''],
        productHeaders
      ];

      allOffers.forEach((offer: any) => {
        const basicPrice = offer.basicPrice;
        const campaignPrice = offer.campaignPrice;
        const quantum = offer.quantum;
        
        const errors = offer.errors && offer.errors.length > 0
          ? offer.errors.map((e: any) => `${e.message}${e.comment ? ': ' + e.comment : ''}`).join('; ')
          : '';
        
        const warnings = offer.warnings && offer.warnings.length > 0
          ? offer.warnings.map((w: any) => `${w.message}${w.comment ? ': ' + w.comment : ''}`).join('; ')
          : '';

        productRows.push([
          offer.offerId || '',
          statusNames[offer.status] || offer.status || '',
          offer.available !== undefined ? (offer.available ? 'Да' : 'Нет') : '',
          basicPrice?.value || '',
          basicPrice?.currencyId || '',
          basicPrice?.discountBase || '',
          campaignPrice?.value || '',
          campaignPrice?.currencyId || '',
          campaignPrice?.vat || '',
          quantum?.minQuantity || '',
          quantum?.stepQuantity || '',
          basicPrice?.updatedAt || campaignPrice?.updatedAt || '',
          errors,
          warnings,
          costPrices[offer.offerId] || 0
        ]);
      });

      const productWorksheet = XLSX.utils.aoa_to_sheet(productRows);
      productWorksheet['!cols'] = [
        { wch: 30 }, { wch: 20 }, { wch: 15 }, { wch: 12 }, { wch: 8 },
        { wch: 15 }, { wch: 12 }, { wch: 8 }, { wch: 8 }, { wch: 15 },
        { wch: 15 }, { wch: 20 }, { wch: 50 }, { wch: 50 }, { wch: 15 }
      ];
      XLSX.utils.book_append_sheet(workbook, productWorksheet, 'Мои товары');
      sheetsCreated++;

      // 2. Загружаем остатки
      setFullReportProgress('📊 Загрузка остатков на складах...');
      console.log('Загрузка остатков для полного отчета...');

      let allStocks: any[] = [];
      pageToken = null;
      attempts = 0;

      do {
        const requestBody: any = { limit: 200, withTurnover: true };
        if (pageToken) requestBody.page_token = pageToken;

        const response = await fetch('/api/stocks', {
          method: 'POST',
          headers: {
            'x-api-key': token,
            'x-campaign-id': campaignId,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(requestBody)
        });

        if (!response.ok) {
          console.error('Ошибка получения остатков:', response.status);
          break;
        }

        const data = await response.json();
        if (data.result?.warehouses && data.result.warehouses.length > 0) {
          data.result.warehouses.forEach((warehouse: any) => {
            if (warehouse.offers && warehouse.offers.length > 0) {
              warehouse.offers.forEach((offer: any) => {
                allStocks.push({ ...offer, warehouseId: warehouse.warehouseId });
              });
            }
          });
        }

        pageToken = data.result?.paging?.nextPageToken || null;
        attempts++;

        if (pageToken && attempts < maxAttempts) {
          await new Promise(resolve => setTimeout(resolve, 100));
        }
      } while (pageToken && attempts < maxAttempts);

      console.log(`Загружено остатков: ${allStocks.length}`);

      // Создаем лист "Остатки"
      const stockHeaders = [
        'Артикул (SKU)', 'Склад ID', 'Доступно (AVAILABLE)', 'Годный (FIT)',
        'Заморожено (FREEZE)', 'Карантин (QUARANTINE)', 'Брак (DEFECT)',
        'Просрочен (EXPIRED)', 'Утилизация (UTILIZATION)', 'Оборачиваемость',
        'Дней оборачиваемости', 'Дата обновления'
      ];

      const stockRows = [
        [`Остатки товаров | Campaign ID: ${campaignId} | Создано: ${new Date().toLocaleString('ru-RU')}`],
        [`Всего записей: ${allStocks.length}`],
        [''],
        stockHeaders
      ];

      const turnoverNames: Record<string, string> = {
        'LOW': 'Низкая (≥120 дней)',
        'ALMOST_LOW': 'Почти низкая (100-119 дней)',
        'HIGH': 'Высокая (45-99 дней)',
        'VERY_HIGH': 'Очень высокая (<45 дней)',
        'NO_SALES': 'Нет продаж',
        'FREE_STORE': 'Бесплатное хранение'
      };

      allStocks.forEach((stock: any) => {
        const stocksByType: Record<string, number> = {};
        if (stock.stocks && stock.stocks.length > 0) {
          stock.stocks.forEach((s: any) => {
            stocksByType[s.type] = s.count;
          });
        }

        stockRows.push([
          stock.offerId || '',
          stock.warehouseId || '',
          stocksByType['AVAILABLE'] || 0,
          stocksByType['FIT'] || 0,
          stocksByType['FREEZE'] || 0,
          stocksByType['QUARANTINE'] || 0,
          stocksByType['DEFECT'] || 0,
          stocksByType['EXPIRED'] || 0,
          stocksByType['UTILIZATION'] || 0,
          stock.turnoverSummary ? turnoverNames[stock.turnoverSummary.turnover] || stock.turnoverSummary.turnover : '',
          stock.turnoverSummary?.turnoverDays || '',
          stock.updatedAt || ''
        ]);
      });

      const stockWorksheet = XLSX.utils.aoa_to_sheet(stockRows);
      stockWorksheet['!cols'] = [
        { wch: 30 }, { wch: 12 }, { wch: 18 }, { wch: 15 }, { wch: 18 },
        { wch: 18 }, { wch: 15 }, { wch: 18 }, { wch: 20 }, { wch: 30 },
        { wch: 20 }, { wch: 25 }
      ];
      XLSX.utils.book_append_sheet(workbook, stockWorksheet, 'Остатки');
      sheetsCreated++;

      // 3. Загружаем отчет по реализации
      setFullReportProgress('📋 Генерация отчета по реализации...');
      console.log('Генерация отчета по реализации...');

      const dateParts = dateB.split('-');
      const year = parseInt(dateParts[0], 10);
      const month = parseInt(dateParts[1], 10);

      const realizationGenRes = await fetch('/api/goods-realization', {
        method: 'POST',
        headers: {
          'x-api-key': token,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          campaignId: Number(campaignId),
          month: month,
          year: year,
          format: 'JSON'
        })
      });

      if (realizationGenRes.ok) {
        const genData = await realizationGenRes.json();
        const reportId = genData?.result?.reportId;

        if (reportId) {
          let fileUrl = null;
          for (let i = 0; i < 30; i++) {
            await new Promise(res => setTimeout(res, 3000));
            const statusRes = await fetch(`/api/report-status?reportId=${reportId}`, {
              headers: { 'x-api-key': token }
            });
            if (statusRes.ok) {
              const statusData = await statusRes.json();
              const status = statusData?.result?.status;
              
              if (status === 'DONE') {
                fileUrl = resolveReportFileUrl(statusData?.result);
                break;
              } else if (status === 'FAILED') {
                break;
              }
            }
          }

          if (fileUrl) {
            const downloadRes = await fetch(`/api/download-report?url=${encodeURIComponent(fileUrl)}`, {
              headers: { 'x-api-key': token }
            });
            
            if (downloadRes.ok) {
              const arrayBuffer = await downloadRes.arrayBuffer();
              const JSZip = (await import('jszip')).default;
              const zip = await JSZip.loadAsync(arrayBuffer);
              
              const realizationSheets = [
                { fileName: 'transferred_to_delivery', ru: 'Переданные в доставку' },
                { fileName: 'delivered', ru: 'Доставленные' },
                { fileName: 'unredeemed', ru: 'Невыкупленные' },
                { fileName: 'returned', ru: 'Возвращенные' },
                { fileName: 'lost_items', ru: 'Утраченные в доставке' }
              ];

              for (const sheetInfo of realizationSheets) {
                const jsonFileName = `${sheetInfo.fileName}.json`;
                const jsonFile = zip.file(jsonFileName);
                
                let sheetData: any[][] = [];
                
                if (!jsonFile) {
                  console.log(`Файл ${jsonFileName} не найден - создаем пустой лист`);
                  sheetData = [['Данные отсутствуют']];
                } else {
                  const jsonContent = await jsonFile.async('string');
                  const jsonData = JSON.parse(jsonContent);
                  
                  if (jsonData.rows && jsonData.rows.length > 0) {
                    const headers = Object.keys(jsonData.rows[0]);
                    sheetData = [headers];
                    
                    jsonData.rows.forEach((row: any) => {
                      const rowData = headers.map(header => row[header] || '');
                      sheetData.push(rowData);
                    });
                    
                    console.log(`Лист ${sheetInfo.ru}: ${jsonData.rows.length} строк данных`);
                  } else {
                    console.log(`Файл ${jsonFileName} не содержит данных - создаем пустой лист`);
                    sheetData = [['Нет данных за выбранный период']];
                  }
                }
                
                const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
                const sheetName = sanitizeSheetName(sheetInfo.ru);
                XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
                sheetsCreated++;
              }
            }
          }
        }
      }

      // 4. Загружаем отчет по стоимости услуг
      setFullReportProgress('💸 Генерация отчета по стоимости услуг...');
      console.log('Генерация отчета по стоимости услуг...');

      const businessId = campaigns.find((c: any) => c.id.toString() === campaignId)?.business?.id;

      if (businessId) {
        const servicesGenRes = await fetch('/api/services-cost', {
          method: 'POST',
          headers: {
            'x-api-key': token,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            businessId,
            dateFrom: dateA,
            dateTo: dateB,
            campaignIds: [Number(campaignId)],
            format: 'JSON',
            language: 'RU'
          })
        });

        if (servicesGenRes.ok) {
          const genData = await servicesGenRes.json();
          const reportId = genData?.result?.reportId;

          if (reportId) {
            let fileUrl = null;
            for (let i = 0; i < 30; i++) {
              await new Promise(res => setTimeout(res, 3000));
              const statusRes = await fetch(`/api/report-status?reportId=${reportId}`, {
                headers: { 'x-api-key': token }
              });
              if (statusRes.ok) {
                const statusData = await statusRes.json();
                const status = statusData?.result?.status;
                
                if (status === 'DONE') {
                  fileUrl = resolveReportFileUrl(statusData?.result);
                  break;
                } else if (status === 'FAILED') {
                  break;
                }
              }
            }

            if (fileUrl) {
              const downloadRes = await fetch(`/api/download-report?url=${encodeURIComponent(fileUrl)}`, {
                headers: { 'x-api-key': token }
              });
              
              if (downloadRes.ok) {
                const arrayBuffer = await downloadRes.arrayBuffer();
                const JSZip = (await import('jszip')).default;
                const zip = await JSZip.loadAsync(arrayBuffer);
                
                const serviceSheets = [
                  { fileName: 'placement', ru: 'Размещение товаров' },
                  { fileName: 'warehouse_processing', ru: 'Складская обработка' },
                  { fileName: 'goods_acceptance', ru: 'Приемка поставки' },
                  { fileName: 'loyalty_and_reviews', ru: 'Лояльность и отзывы' },
                  { fileName: 'boost', ru: 'Буст продаж (продажи)' },
                  { fileName: 'installment_plan', ru: 'Рассрочка' },
                  { fileName: 'shelf', ru: 'Полки' },
                  { fileName: 'cpm-boost', ru: 'Буст продаж (показы)' },
                  { fileName: 'banners', ru: 'Баннеры' },
                  { fileName: 'pushes', ru: 'Пуш-уведомления' },
                  { fileName: 'popups', ru: 'Поп-ап уведомления' },
                  { fileName: 'delivery', ru: 'Доставка покупателю' },
                  { fileName: 'express_delivery', ru: 'Экспресс-доставка' },
                  { fileName: 'delivery_from_abroad', ru: 'Доставка из-за рубежа' },
                  { fileName: 'payment_accepting', ru: 'Приём платежа' },
                  { fileName: 'payment_transfer', ru: 'Перевод платежа' },
                  { fileName: 'paid_storage_before_31-05-22', ru: 'Хранение до 31.05.22' },
                  { fileName: 'paid_storage_after_01-06-22', ru: 'Хранение с 01.06.22' },
                  { fileName: 'delivery_via_transit_warehouse', ru: 'Транзитный склад' },
                  { fileName: 'reception_of_surplus', ru: 'Приём излишков' },
                  { fileName: 'product_marking', ru: 'Маркировка' },
                  { fileName: 'export_from_warehouse', ru: 'Вывоз со склада' },
                  { fileName: 'intake_logistics', ru: 'Забор заказов' },
                  { fileName: 'order_processing', ru: 'Обработка в СЦ/ПВЗ' },
                  { fileName: 'order_processing_on_warehouse', ru: 'Обработка на складе' },
                  { fileName: 'storage_of_returns', ru: 'Хранение возвратов' },
                  { fileName: 'expropriation', ru: 'Вознаграждение' },
                  { fileName: 'utilization', ru: 'Утилизация' },
                  { fileName: 'extended_service_access', ru: 'Расширенный доступ' },
                  { fileName: 'personal_manager', ru: 'Персональный менеджер' }
                ];

                for (const sheetInfo of serviceSheets) {
                  const jsonFileName = `${sheetInfo.fileName}.json`;
                  const jsonFile = zip.file(jsonFileName);
                  
                  let sheetData: any[][] = [];
                  
                  if (!jsonFile) {
                    console.log(`Файл ${jsonFileName} не найден - создаем пустой лист`);
                    sheetData = [['Данные отсутствуют']];
                  } else {
                    const jsonContent = await jsonFile.async('string');
                    const jsonData = JSON.parse(jsonContent);
                    
                    if (jsonData.rows && jsonData.rows.length > 0) {
                      const headers = Object.keys(jsonData.rows[0]);
                      sheetData = [headers];
                      
                      jsonData.rows.forEach((row: any) => {
                        const rowData = headers.map(header => row[header] || '');
                        sheetData.push(rowData);
                      });
                      
                      console.log(`Лист ${sheetInfo.ru}: ${jsonData.rows.length} строк данных`);
                    } else {
                      console.log(`Файл ${jsonFileName} не содержит данных - создаем пустой лист`);
                      sheetData = [['Нет данных за выбранный период']];
                    }
                  }
                  
                  const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
                  const sheetName = sanitizeSheetName(sheetInfo.ru);
                  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
                  sheetsCreated++;
                }
              }
            }
          }
        }
      }

      // Генерация и скачивание итогового файла
      setFullReportProgress('💾 Сохранение файла...');
      
      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      link.href = url;
      link.download = `full_report_${campaignId}_${new Date().toISOString().split('T')[0]}.xlsx`;
      link.click();
      URL.revokeObjectURL(url);
      
      setFullReportProgress('');
      alert(`✅ Полный отчет скачан!\n\nСоздано листов: ${sheetsCreated}\n• Мои товары\n• Остатки\n• Отчет по реализации (5 листов)\n• Стоимость услуг (до 30 листов)`);
      
    } catch (error) {
      console.error('Ошибка при создании полного отчета:', error);
      setFullReportProgress('');
      alert('Произошла ошибка при создании полного отчета: ' + (error as Error).message);
    } finally {
      setIsLoading(false);
    }
  };

  // Функция для кнопки "Тест" - скачивание отчета по реализации с данными
  const downloadTestExcel = async () => {
    setIsLoading(true);
    try {
      if (!token || !campaignId) {
        alert('Заполните токен и выберите магазин');
        setIsLoading(false);
        return;
      }

      // Получаем месяц и год из dateB
      const dateParts = dateB.split('-');
      const year = parseInt(dateParts[0], 10);
      const month = parseInt(dateParts[1], 10);

      console.log(`Тестовая загрузка отчета по реализации за ${month}/${year}`);

      // 1) Генерация отчета в формате JSON
      const generateRes = await fetch('/api/goods-realization', {
        method: 'POST',
        headers: {
          'x-api-key': token,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          campaignId: Number(campaignId),
          month: month,
          year: year,
          format: 'JSON' // Запрашиваем JSON формат
        })
      });

      if (!generateRes.ok) {
        const txt = await generateRes.text();
        console.error('Ошибка генерации:', txt);
        alert('Ошибка запуска генерации тестового отчета');
        setIsLoading(false);
        return;
      }

      const genData = await generateRes.json();
      const reportId = genData?.result?.reportId;
      if (!reportId) {
        alert('ReportId не получен');
        setIsLoading(false);
        return;
      }

      console.log(`Отчет запущен, reportId: ${reportId}`);

      // 2) Ожидание готовности отчета
      let fileUrl = null;
      for (let i = 0; i < 30; i++) {
        await new Promise(res => setTimeout(res, 3000));
        const statusRes = await fetch(`/api/report-status?reportId=${reportId}`, {
          headers: { 'x-api-key': token }
        });
        if (!statusRes.ok) continue;
        const statusData = await statusRes.json();
        const status = statusData?.result?.status;
        
        console.log(`Статус (попытка ${i + 1}): ${status}`);
        
        if (status === 'DONE') {
          fileUrl = resolveReportFileUrl(statusData?.result);
          break;
        } else if (status === 'FAILED') {
          alert('Генерация отчета завершилась с ошибкой');
          setIsLoading(false);
          return;
        }
      }

      if (!fileUrl) {
        alert('Не удалось получить ссылку на файл');
        setIsLoading(false);
        return;
      }

      // 3) Скачивание ZIP архива с JSON файлами
      const downloadRes = await fetch(`/api/download-report?url=${encodeURIComponent(fileUrl)}`, {
        headers: { 'x-api-key': token }
      });
      
      if (!downloadRes.ok) {
        alert('Не удалось скачать файл');
        setIsLoading(false);
        return;
      }

      // 4) Обработка ZIP архива и создание Excel
      const arrayBuffer = await downloadRes.arrayBuffer();
      
      // Используем библиотеку jszip для распаковки
      const JSZip = (await import('jszip')).default;
      const zip = await JSZip.loadAsync(arrayBuffer);
      
      // Создаем Excel книгу
      const workbook = XLSX.utils.book_new();
      
      // Определение листов которые нужно обработать
      const sheetsToProcess = [
        { fileName: 'transferred_to_delivery', ru: 'Переданные в доставку' },
        { fileName: 'delivered', ru: 'Доставленные' },
        { fileName: 'unredeemed', ru: 'Невыкупленные' },
        { fileName: 'returned', ru: 'Возвращенные' },
        { fileName: 'lost_items', ru: 'Утраченные в доставке' }
      ];

      // Обрабатываем каждый JSON файл из архива
      for (const sheetInfo of sheetsToProcess) {
        const jsonFileName = `${sheetInfo.fileName}.json`;
        const jsonFile = zip.file(jsonFileName);
        
        let sheetData: any[][] = [];
        
        if (!jsonFile) {
          console.log(`Файл ${jsonFileName} не найден в архиве - создаем пустой лист`);
          // Создаем пустой лист с информационным сообщением
          sheetData = [['Данные отсутствуют']];
        } else {
          const jsonContent = await jsonFile.async('string');
          const jsonData = JSON.parse(jsonContent);
          
          if (jsonData.rows && jsonData.rows.length > 0) {
            // Получаем заголовки из первой строки данных
            const headers = Object.keys(jsonData.rows[0]);
            
            // Создаем массив для листа с заголовками
            sheetData = [headers];
            
            // Добавляем данные
            jsonData.rows.forEach((row: any) => {
              const rowData = headers.map(header => row[header] || '');
              sheetData.push(rowData);
            });
            
            console.log(`Лист ${sheetInfo.ru}: ${jsonData.rows.length} строк данных`);
          } else {
            console.log(`Файл ${jsonFileName} не содержит данных - создаем пустой лист`);
            // Создаем пустой лист с информационным сообщением
            sheetData = [['Нет данных за выбранный период']];
          }
        }
        
        // Создаем лист (даже если он пустой)
        const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
        
        // Добавляем лист в книгу (очищаем название от запрещенных символов и обрезаем до 31 символа)
        const sheetName = sanitizeSheetName(sheetInfo.ru);
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        
        console.log(`Добавлен лист: ${sheetName}`);
      }

      // 5) Генерация и скачивание Excel файла
      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      link.href = url;
      link.download = `realization_report_${year}_${month.toString().padStart(2, '0')}.xlsx`;
      link.click();
      URL.revokeObjectURL(url);
      
      alert(`Тестовый отчет скачан! Листов: ${workbook.SheetNames.length}`);
      
    } catch (error) {
      console.error('Ошибка при создании тестового файла:', error);
      alert('Произошла ошибка при создании тестового файла: ' + (error as Error).message);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 dark:from-gray-900 dark:to-gray-800 p-8">
      <div className="max-w-2xl mx-auto">
        <div className="bg-white dark:bg-gray-800 rounded-xl shadow-lg p-8">
          <h1 className="text-3xl font-bold text-gray-900 dark:text-white mb-8 text-center">
            Яндекс.Маркет Отчеты
          </h1>
          
          <div className="space-y-6">
            {/* Поле токена */}
            <div>
              <label htmlFor="token" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">
                API-Key Токен
              </label>
              <input
                type="text"
                id="token"
                value={token}
                onChange={(e) => setToken(e.target.value)}
                className="w-full px-4 py-3 border border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent dark:bg-gray-700 dark:text-white transition-colors"
                placeholder="ACMA:ваш_токен:код"
              />
              <p className="mt-1 text-xs text-gray-500 dark:text-gray-400">
                💡 Формат: ACMA:токен:код (создается в кабинете Яндекс.Маркета → Модули и API)
              </p>
            </div>

            {/* Выбор магазина */}
            <div>
              <div className="flex items-center justify-between mb-2">
                <label htmlFor="campaignSelect" className="block text-sm font-medium text-gray-700 dark:text-gray-300">
                  Выберите магазин
                </label>
                <button
                  type="button"
                  onClick={loadCampaigns}
                  disabled={isLoadingCampaigns || !token}
                  className="text-xs bg-blue-500 hover:bg-blue-600 disabled:bg-gray-400 text-white px-3 py-1 rounded transition-colors"
                >
                  {isLoadingCampaigns ? 'Загрузка...' : 'Обновить список'}
                </button>
        </div>
              
              {campaigns.length > 0 ? (
                <select
                  id="campaignSelect"
                  value={campaignId}
                  onChange={(e) => setCampaignId(e.target.value)}
                  className="w-full px-4 py-3 border border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent dark:bg-gray-700 dark:text-white transition-colors"
                >
                  <option value="">Выберите магазин</option>
                  {campaigns.map((campaign: any) => (
                    <option key={campaign.id} value={campaign.id}>
                      {campaign.name} (ID: {campaign.id}) - {campaign.placementType}
                      {campaign.apiAvailability === 'DISABLED_BY_INACTIVITY' ? ' - НЕАКТИВЕН' : ''}
                    </option>
                  ))}
                </select>
              ) : (
                <div className="w-full px-4 py-3 border border-gray-300 dark:border-gray-600 rounded-lg bg-gray-50 dark:bg-gray-700 text-gray-500 dark:text-gray-400">
                  {isLoadingCampaigns ? 'Загружаем список магазинов...' : 'Нажмите "Обновить список" для загрузки магазинов'}
                </div>
              )}
              
              <p className="mt-1 text-xs text-gray-500 dark:text-gray-400">
                💡 Магазины загружаются автоматически из вашего кабинета Яндекс.Маркета
              </p>
            </div>

            {/* Выбор завершенного месяца */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label htmlFor="month" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">
                  Выберите месяц
                </label>
                <select
                  id="month"
                  value={selectedMonth}
                  onChange={(e) => setSelectedMonth(Number(e.target.value))}
                  className="w-full px-4 py-3 border border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent dark:bg-gray-700 dark:text-white transition-colors"
                >
                  <option value={0}>Январь</option>
                  <option value={1}>Февраль</option>
                  <option value={2}>Март</option>
                  <option value={3}>Апрель</option>
                  <option value={4}>Май</option>
                  <option value={5}>Июнь</option>
                  <option value={6}>Июль</option>
                  <option value={7}>Август</option>
                  <option value={8}>Сентябрь</option>
                  <option value={9}>Октябрь</option>
                  <option value={10}>Ноябрь</option>
                  <option value={11}>Декабрь</option>
                </select>
              </div>
              
              <div>
                <label htmlFor="year" className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">
                  Выберите год
                </label>
                <select
                  id="year"
                  value={selectedYear}
                  onChange={(e) => setSelectedYear(Number(e.target.value))}
                  className="w-full px-4 py-3 border border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent dark:bg-gray-700 dark:text-white transition-colors"
                >
                  {Array.from({ length: 5 }, (_, i) => currentDate.getFullYear() - i).map(year => (
                    <option key={year} value={year}>{year}</option>
                  ))}
                </select>
              </div>
            </div>

            {/* Показываем выбранный период */}
            <div className="mt-2 p-3 bg-gray-50 dark:bg-gray-700 rounded-lg">
              <p className="text-sm text-gray-600 dark:text-gray-400">
                📅 Выбранный период: <span className="font-medium text-gray-900 dark:text-white">{dateA} — {dateB}</span>
              </p>
            </div>

            {/* Информация о токене */}
            <div className="mt-4 p-4 bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-800 rounded-lg">
              <h4 className="text-sm font-semibold text-blue-800 dark:text-blue-200 mb-1">
                🔑 Требования к токену
              </h4>
              <div className="text-xs text-blue-700 dark:text-blue-300 space-y-1">
                <p>• Токен должен иметь формат: ACMA:токен:код</p>
                <p>• Требуемые доступы: "Отчеты" и "Информация о кампаниях"</p>
                <p>• Токен создается в кабинете Яндекс.Маркета → Модули и API</p>
                <p>• Запросы проходят через серверный прокси (без CORS ограничений)</p>
              </div>
            </div>

            {/* Кнопка "Себестоимость" */}
            <div className="pt-6">
              <button
                onClick={downloadCostPrice}
                disabled={isLoading}
                className="w-full bg-cyan-500 hover:bg-cyan-600 disabled:bg-cyan-400 disabled:cursor-not-allowed text-white font-medium py-3 px-4 rounded-lg transition-colors duration-200 flex items-center justify-center gap-2 shadow-sm hover:shadow-md"
              >
                {isLoading ? (
                  <>
                    <svg className="w-5 h-5 animate-spin" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                    </svg>
                    Загрузка...
                  </>
                ) : (
                  <>
                    <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                    💰 Себестоимость
                  </>
                )}
              </button>

              <p className="mt-2 text-xs text-gray-500 dark:text-gray-400 text-center">
                💡 Откроет окно для ввода себестоимости товаров и расчета маржи. Поддерживает импорт/экспорт Excel
              </p>
            </div>

            {/* Кнопка "Налог" */}
            <div className="pt-6">
              <button
                onClick={() => {}}
                disabled={isLoading}
                className="w-full bg-orange-500 hover:bg-orange-600 disabled:bg-orange-400 disabled:cursor-not-allowed text-white font-medium py-3 px-4 rounded-lg transition-colors duration-200 flex items-center justify-center gap-2 shadow-sm hover:shadow-md"
              >
                {isLoading ? (
                  <>
                    <svg className="w-5 h-5 animate-spin" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                    </svg>
                    Загрузка...
                  </>
                ) : (
                  <>
                    <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 14l6-6m-5.5.5h.01m4.99 5h.01M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16l3.5-2 3.5 2 3.5-2 3.5 2zM10 8.5a.5.5 0 11-1 0 .5.5 0 011 0zm5 5a.5.5 0 11-1 0 .5.5 0 011 0z" />
                    </svg>
                    🧾 Налог
                  </>
                )}
              </button>

              <p className="mt-2 text-xs text-gray-500 dark:text-gray-400 text-center">
                💡 Функционал в разработке
              </p>
            </div>

            {/* Кнопка "Полный отчет" */}
            <div className="pt-6 border-t-2 border-gray-200 dark:border-gray-700 mt-6">
              <button
                onClick={downloadFullReport}
                disabled={isLoading}
                className="w-full bg-gradient-to-r from-purple-600 to-pink-600 hover:from-purple-700 hover:to-pink-700 disabled:from-purple-400 disabled:to-pink-400 disabled:cursor-not-allowed text-white font-bold py-4 px-6 rounded-lg transition-all duration-200 flex items-center justify-center gap-3 shadow-lg hover:shadow-xl transform hover:scale-[1.02]"
              >
                {isLoading ? (
                  <>
                    <svg className="w-6 h-6 animate-spin" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                    </svg>
                    {fullReportProgress || 'Формирование полного отчета...'}
                  </>
                ) : (
                  <>
                    <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                    </svg>
                    ⭐ ПОЛНЫЙ ОТЧЕТ (ВСЁ ВМЕСТЕ)
                  </>
                )}
              </button>

              <div className="mt-3 p-3 bg-gradient-to-r from-purple-50 to-pink-50 dark:from-purple-900/20 dark:to-pink-900/20 border border-purple-200 dark:border-purple-800 rounded-lg">
                <h4 className="text-sm font-semibold text-purple-800 dark:text-purple-200 mb-1">
                  📊 Что включает полный отчет:
                </h4>
                <div className="text-xs text-purple-700 dark:text-purple-300 space-y-1 ml-4">
                  <p>✓ Мои товары (с ценами, статусами и себестоимостью)</p>
                  <p>✓ Остатки на складах (с оборачиваемостью)</p>
                  <p>✓ Отчет по реализации (5 листов)</p>
                  <p>✓ Стоимость услуг (до 30 листов)</p>
                  <p className="font-semibold mt-2">⚡ Всё в одном файле Excel!</p>
                </div>
              </div>

              <p className="mt-2 text-xs text-gray-500 dark:text-gray-400 text-center">
                ⏱️ Формирование может занять 2-5 минут (зависит от объема данных)
              </p>
            </div>

            </div>

          {/* Информация о текущих настройках */}
          <div className="mt-8 p-4 bg-gray-50 dark:bg-gray-700 rounded-lg">
            <h3 className="text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">
              Текущие настройки:
            </h3>
            <div className="text-sm text-gray-600 dark:text-gray-400 space-y-1">
              <p><span className="font-medium">Токен:</span> {token ? `${token.substring(0, 20)}...` : 'Не указан'}</p>
              <p><span className="font-medium">Загружено магазинов:</span> {campaigns.length > 0 ? campaigns.length : 'Не загружены'}</p>
              <p><span className="font-medium">Выбранный магазин:</span> {
                campaignId && campaigns.length > 0 
                  ? campaigns.find((c: any) => c.id.toString() === campaignId)?.name || `ID: ${campaignId}`
                  : 'Не выбран'
              }</p>
              <p><span className="font-medium">Период отчета:</span> {dateA && dateB ? `${dateA} — ${dateB}` : 'Не выбран'}</p>
            </div>
          </div>
        </div>
            </div>

      {/* Модальное окно для себестоимости */}
      {showCostPriceModal && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
          <div className="bg-white dark:bg-gray-800 rounded-xl shadow-2xl max-w-6xl w-full max-h-[90vh] flex flex-col">
            {/* Заголовок модального окна */}
            <div className="p-6 border-b border-gray-200 dark:border-gray-700 flex justify-between items-center">
              <h2 className="text-2xl font-bold text-gray-900 dark:text-white">
                💰 Управление себестоимостью товаров
              </h2>
              <button
                onClick={() => setShowCostPriceModal(false)}
                className="text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-200 text-2xl"
              >
                ×
              </button>
            </div>

            {/* Информационная панель */}
            <div className="p-4 bg-blue-50 dark:bg-blue-900/20 border-b border-blue-200 dark:border-blue-800">
              <div className="flex justify-between items-center text-sm">
                <span className="text-blue-800 dark:text-blue-200">
                  📊 Всего товаров: <strong>{products.length}</strong> | 
                  ✅ Заполнено: <strong>{Object.keys(costPrices).length}</strong> | 
                  ⏳ Осталось: <strong>{products.length - Object.keys(costPrices).length}</strong>
                  <span className="ml-2 text-xs">💾 Автосохранение включено</span>
                </span>
                <div className="flex gap-2">
              <button
                    onClick={clearAllCostPrices}
                    className="bg-red-500 hover:bg-red-600 text-white px-4 py-2 rounded-lg text-sm font-medium transition-colors"
                  >
                    🗑️ Очистить всё
              </button>
              <label className="bg-blue-500 hover:bg-blue-600 text-white px-4 py-2 rounded-lg text-sm font-medium transition-colors cursor-pointer">
                📤 Импортировать из Excel
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={importCostPricesFromExcel}
                  className="hidden"
                />
              </label>
              <button
                    onClick={exportCostPricesToExcel}
                    className="bg-green-500 hover:bg-green-600 text-white px-4 py-2 rounded-lg text-sm font-medium transition-colors"
                  >
                    📥 Экспортировать в Excel
              </button>
                </div>
              </div>
            </div>

            {/* Таблица с товарами */}
            <div className="flex-1 overflow-auto p-6">
              <table className="w-full border-collapse">
                <thead className="sticky top-0 bg-gray-100 dark:bg-gray-700 z-10">
                  <tr>
                    <th className="p-3 text-left text-sm font-semibold text-gray-700 dark:text-gray-300 border-b-2 border-gray-300 dark:border-gray-600">
                      Артикул (SKU)
                    </th>
                    <th className="p-3 text-left text-sm font-semibold text-gray-700 dark:text-gray-300 border-b-2 border-gray-300 dark:border-gray-600">
                      Цена продажи
                    </th>
                    <th className="p-3 text-left text-sm font-semibold text-gray-700 dark:text-gray-300 border-b-2 border-gray-300 dark:border-gray-600">
                      Себестоимость
                    </th>
                    <th className="p-3 text-left text-sm font-semibold text-gray-700 dark:text-gray-300 border-b-2 border-gray-300 dark:border-gray-600">
                      Маржа (₽)
                    </th>
                    <th className="p-3 text-left text-sm font-semibold text-gray-700 dark:text-gray-300 border-b-2 border-gray-300 dark:border-gray-600">
                      Маржа (%)
                    </th>
                    <th className="p-3 text-left text-sm font-semibold text-gray-700 dark:text-gray-300 border-b-2 border-gray-300 dark:border-gray-600">
                      Статус
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {products.map((product: any, index: number) => {
                    const salePrice = product.campaignPrice?.value || product.basicPrice?.value || 0;
                    const costPrice = costPrices[product.offerId] || 0;
                    const margin = salePrice - costPrice;
                    const marginPercent = costPrice > 0 ? ((margin / costPrice) * 100).toFixed(2) : '0';
                    const currency = product.campaignPrice?.currencyId || product.basicPrice?.currencyId || 'RUR';

                    const statusNames: Record<string, string> = {
                      'PUBLISHED': 'Готов к продаже',
                      'CHECKING': 'На проверке',
                      'DISABLED_BY_PARTNER': 'Скрыт',
                      'REJECTED_BY_MARKET': 'Отклонен',
                      'DISABLED_AUTOMATICALLY': 'Ошибки',
                      'CREATING_CARD': 'Создается',
                      'NO_CARD': 'Нужна карточка',
                      'NO_STOCKS': 'Нет на складе',
                      'ARCHIVED': 'Архив'
                    };

                    return (
                      <tr 
                        key={product.offerId} 
                        className={`${index % 2 === 0 ? 'bg-white dark:bg-gray-800' : 'bg-gray-50 dark:bg-gray-750'} hover:bg-blue-50 dark:hover:bg-blue-900/20 transition-colors`}
                      >
                        <td className="p-3 text-sm text-gray-900 dark:text-gray-100 border-b border-gray-200 dark:border-gray-700">
                          {product.offerId}
                        </td>
                        <td className="p-3 text-sm text-gray-900 dark:text-gray-100 border-b border-gray-200 dark:border-gray-700">
                          {salePrice} {currency}
                        </td>
                        <td className="p-3 border-b border-gray-200 dark:border-gray-700">
                          <input
                            type="number"
                            min="0"
                            step="0.01"
                            value={costPrices[product.offerId] || ''}
                            onChange={(e) => updateCostPrice(product.offerId, e.target.value)}
                            placeholder="Введите"
                            className="w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:border-transparent dark:bg-gray-700 dark:text-white text-sm"
                          />
                        </td>
                        <td className={`p-3 text-sm font-medium border-b border-gray-200 dark:border-gray-700 ${
                          margin > 0 ? 'text-green-600 dark:text-green-400' : 
                          margin < 0 ? 'text-red-600 dark:text-red-400' : 
                          'text-gray-500 dark:text-gray-400'
                        }`}>
                          {costPrice > 0 ? `${margin.toFixed(2)} ${currency}` : '-'}
                        </td>
                        <td className={`p-3 text-sm font-medium border-b border-gray-200 dark:border-gray-700 ${
                          parseFloat(marginPercent) > 0 ? 'text-green-600 dark:text-green-400' : 
                          parseFloat(marginPercent) < 0 ? 'text-red-600 dark:text-red-400' : 
                          'text-gray-500 dark:text-gray-400'
                        }`}>
                          {costPrice > 0 ? `${marginPercent}%` : '-'}
                        </td>
                        <td className="p-3 text-xs text-gray-600 dark:text-gray-400 border-b border-gray-200 dark:border-gray-700">
                          {statusNames[product.status] || product.status}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
          </div>

            {/* Футер модального окна */}
            <div className="p-6 border-t border-gray-200 dark:border-gray-700 flex justify-between items-center">
              <div className="text-sm text-gray-600 dark:text-gray-400">
                💡 Введите себестоимость вручную или импортируйте из Excel. Данные сохраняются автоматически в браузере. Вы можете экспортировать файл, отредактировать его и импортировать обратно.
            </div>
              <button
                onClick={() => setShowCostPriceModal(false)}
                className="bg-gray-500 hover:bg-gray-600 text-white px-6 py-2 rounded-lg font-medium transition-colors"
              >
                Закрыть
              </button>
          </div>
        </div>
      </div>
      )}
    </div>
  );
}