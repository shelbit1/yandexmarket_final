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
    const url = 'https://api.partner.market.yandex.ru/v2/reports/shows-sales/generate?format=FILE';
    
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
        { error: `Sales Analytics API Error: ${response.status} - ${errorText}` },
        { status: response.status }
      );
    }

    const data = await response.json();
    return NextResponse.json(data);

  } catch (error) {
    console.error('Ошибка при запросе аналитики продаж:', error);
    return NextResponse.json(
      { error: 'Ошибка при запросе к API' },
      { status: 500 }
    );
  }
}
