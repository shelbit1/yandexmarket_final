import { NextRequest, NextResponse } from 'next/server';

export async function POST(request: NextRequest) {
  const apiKey = request.headers.get('x-api-key');
  const body = await request.json();

  if (!apiKey) {
    return NextResponse.json(
      { error: 'API ключ не предоставлен' },
      { status: 400 }
    );
  }

  try {
    const url = 'https://api.partner.market.yandex.ru/v2/reports/boost-consolidated/generate?format=FILE';

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Api-Key': apiKey,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return NextResponse.json(
        { error: `Boost Report API Error: ${response.status} - ${errorText}` },
        { status: response.status }
      );
    }

    const data = await response.json();
    return NextResponse.json(data);
  } catch (error: any) {
    console.error('Ошибка при генерации буста продаж:', error);
    return NextResponse.json(
      { error: 'Ошибка при запросе к API', details: error.message },
      { status: 500 }
    );
  }
}


