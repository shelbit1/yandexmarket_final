import { NextResponse } from 'next/server';

export async function POST(request: Request) {
  const apiKey = request.headers.get('x-api-key');
  const body = await request.json();

  if (!apiKey) {
    return NextResponse.json({ error: 'API Key is missing' }, { status: 400 });
  }

  try {
    const yandexResponse = await fetch('https://api.partner.market.yandex.ru/v2/reports/united-orders/generate?format=FILE&language=RU', {
      method: 'POST',
      headers: {
        'Api-Key': apiKey,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
    });

    if (!yandexResponse.ok) {
      const errorText = await yandexResponse.text();
      return NextResponse.json({ error: `Yandex API Error: ${yandexResponse.status} - ${errorText}` }, { status: yandexResponse.status });
    }

    const data = await yandexResponse.json();
    return NextResponse.json(data);
  } catch (error: any) {
    console.error('Proxy error for orders report:', error);
    return NextResponse.json({ error: 'Internal Server Error', details: error.message }, { status: 500 });
  }
}
