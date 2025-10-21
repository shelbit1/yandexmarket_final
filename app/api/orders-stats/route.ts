import { NextRequest, NextResponse } from 'next/server';

export async function POST(request: NextRequest) {
  try {
    const apiKey = request.headers.get('x-api-key');
    const campaignId = request.headers.get('x-campaign-id');
    
    if (!apiKey) {
      return NextResponse.json({ error: 'API ключ не указан' }, { status: 400 });
    }

    if (!campaignId) {
      return NextResponse.json({ error: 'Campaign ID не указан' }, { status: 400 });
    }

    const body = await request.json();
    console.log('Получен запрос на детальную информацию по заказам:', body);

    // Формируем URL для запроса к Яндекс.Маркету
    const yandexUrl = `https://api.partner.market.yandex.ru/v2/campaigns/${campaignId}/stats/orders`;
    
    // Добавляем параметры пагинации если они есть
    const url = new URL(yandexUrl);
    if (body.limit) {
      url.searchParams.append('limit', body.limit.toString());
    }
    if (body.page_token) {
      url.searchParams.append('page_token', body.page_token);
    }

    // Подготавливаем тело запроса
    const requestBody: any = {};
    
    // Фильтр по датам создания заказов
    if (body.dateFrom) {
      requestBody.dateFrom = body.dateFrom;
    }
    if (body.dateTo) {
      requestBody.dateTo = body.dateTo;
    }

    // Фильтр по датам обновления заказов
    if (body.updateFrom) {
      requestBody.updateFrom = body.updateFrom;
    }
    if (body.updateTo) {
      requestBody.updateTo = body.updateTo;
    }

    // Фильтр по статусам заказов
    if (body.statuses && Array.isArray(body.statuses) && body.statuses.length > 0) {
      requestBody.statuses = body.statuses;
    }

    // Фильтр по конкретным заказам
    if (body.orders && Array.isArray(body.orders) && body.orders.length > 0) {
      requestBody.orders = body.orders;
    }

    // Фильтр по товарам с кодами маркировки
    if (typeof body.hasCis === 'boolean') {
      requestBody.hasCis = body.hasCis;
    }

    console.log('Отправляем запрос к Яндекс.Маркету:', {
      url: url.toString(),
      body: requestBody
    });

    const response = await fetch(url.toString(), {
      method: 'POST',
      headers: {
        'Api-Key': apiKey,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    console.log('Статус ответа от Яндекс.Маркета:', response.status);

    if (!response.ok) {
      const errorText = await response.text();
      console.error('Ошибка от Яндекс.Маркета:', errorText);
      
      let errorMessage = `HTTP ${response.status}`;
      try {
        const errorData = JSON.parse(errorText);
        if (errorData.errors && errorData.errors.length > 0) {
          errorMessage = errorData.errors.map((e: any) => `${e.code}: ${e.message || 'Неизвестная ошибка'}`).join('; ');
        }
      } catch (e) {
        errorMessage += `: ${errorText}`;
      }
      
      return NextResponse.json({ 
        error: errorMessage,
        status: response.status,
        details: errorText
      }, { status: response.status });
    }

    const data = await response.json();
    console.log('Успешно получены данные от Яндекс.Маркета, количество заказов:', data.result?.orders?.length || 0);

    return NextResponse.json(data);

  } catch (error: any) {
    console.error('Ошибка при обработке запроса:', error);
    return NextResponse.json(
      { error: 'Внутренняя ошибка сервера', details: error.message },
      { status: 500 }
    );
  }
}
