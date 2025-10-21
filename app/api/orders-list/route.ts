import { NextResponse } from 'next/server';

export async function GET(request: Request) {
  const apiKey = request.headers.get('x-api-key');
  const campaignId = request.headers.get('x-campaign-id');
  
  if (!apiKey || !campaignId) {
    return NextResponse.json({ error: 'API Key and Campaign ID are required' }, { status: 400 });
  }

  try {
    // Получаем параметры запроса из URL
    const url = new URL(request.url);
    const searchParams = url.searchParams;
    
    // Строим URL для Yandex API с переданными параметрами
    const yandexUrl = new URL(`https://api.partner.market.yandex.ru/v2/campaigns/${campaignId}/orders`);
    
    // Передаем все query параметры в Yandex API
    searchParams.forEach((value, key) => {
      yandexUrl.searchParams.append(key, value);
    });

    const yandexResponse = await fetch(yandexUrl.toString(), {
      method: 'GET',
      headers: {
        'Api-Key': apiKey,
        'Content-Type': 'application/json',
      },
    });

    if (!yandexResponse.ok) {
      const errorText = await yandexResponse.text();
      return NextResponse.json({ 
        error: `Yandex API Error: ${yandexResponse.status} - ${errorText}` 
      }, { status: yandexResponse.status });
    }

    const data = await yandexResponse.json();
    return NextResponse.json(data);
  } catch (error: any) {
    console.error('Proxy error for orders list:', error);
    return NextResponse.json({ 
      error: 'Internal Server Error', 
      details: error.message 
    }, { status: 500 });
  }
}

