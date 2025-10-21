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
    const { campaignId, month, year, format = 'FILE' } = body;

    if (!campaignId || !month || !year) {
      return NextResponse.json(
        { error: 'Необходимо указать campaignId, month и year' },
        { status: 400 }
      );
    }

    // Проверяем валидность месяца
    if (month < 1 || month > 12) {
      return NextResponse.json(
        { error: 'Месяц должен быть в диапазоне 1-12' },
        { status: 400 }
      );
    }

    const url = `https://api.partner.market.yandex.ru/v2/reports/goods-realization/generate?format=${format}`;

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Api-Key': token,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        campaignId: Number(campaignId),
        month: Number(month),
        year: Number(year),
      }),
    });

    const data = await response.json();

    if (!response.ok) {
      console.error('Ошибка при генерации отчета по реализации:', data);
      return NextResponse.json(
        { error: 'Ошибка при генерации отчета', details: data },
        { status: response.status }
      );
    }

    return NextResponse.json(data);
  } catch (error: any) {
    console.error('Ошибка в API goods-realization:', error);
    return NextResponse.json(
      { error: 'Внутренняя ошибка сервера', message: error.message },
      { status: 500 }
    );
  }
}

