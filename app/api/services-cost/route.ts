import { NextRequest, NextResponse } from 'next/server';

export async function POST(request: NextRequest) {
  try {
    const token = request.headers.get('x-api-key');
    if (!token) {
      return NextResponse.json(
        { error: 'API-Key токен не предоставлен' },
        { status: 401 }
      );
    }

    const body = await request.json();
    const { businessId, dateFrom, dateTo, campaignIds, format = 'FILE', language = 'RU' } = body;

    if (!businessId) {
      return NextResponse.json(
        { error: 'Необходимо указать businessId' },
        { status: 400 }
      );
    }

    if (!dateFrom || !dateTo) {
      return NextResponse.json(
        { error: 'Необходимо указать dateFrom и dateTo' },
        { status: 400 }
      );
    }

    const url = `https://api.partner.market.yandex.ru/v2/reports/united-marketplace-services/generate?format=${format}&language=${language}`;

    const requestBody: any = {
      businessId: Number(businessId),
      dateFrom,
      dateTo,
    };

    // Добавляем campaignIds если указан
    if (campaignIds && Array.isArray(campaignIds) && campaignIds.length > 0) {
      requestBody.campaignIds = campaignIds.map((id: any) => Number(id));
    }

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Api-Key': token,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    const data = await response.json();

    if (!response.ok) {
      console.error('Ошибка при генерации отчета по стоимости услуг:', data);
      return NextResponse.json(
        { error: 'Ошибка при генерации отчета', details: data },
        { status: response.status }
      );
    }

    return NextResponse.json(data);
  } catch (error: any) {
    console.error('Ошибка в API services-cost:', error);
    return NextResponse.json(
      { error: 'Внутренняя ошибка сервера', message: error.message },
      { status: 500 }
    );
  }
}



