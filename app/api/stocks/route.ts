import { NextRequest, NextResponse } from 'next/server';

export async function POST(request: NextRequest) {
  try {
    const token = request.headers.get('x-api-key');
    const campaignId = request.headers.get('x-campaign-id');
    
    if (!token) {
      return NextResponse.json(
        { error: 'API key is required' },
        { status: 401 }
      );
    }

    if (!campaignId) {
      return NextResponse.json(
        { error: 'Campaign ID is required' },
        { status: 400 }
      );
    }

    const body = await request.json();
    
    // Формируем URL с query параметрами
    const url = new URL(`https://api.partner.market.yandex.ru/v2/campaigns/${campaignId}/offers/stocks`);
    
    if (body.limit) {
      url.searchParams.append('limit', body.limit.toString());
    }
    
    if (body.page_token) {
      url.searchParams.append('page_token', body.page_token);
    }

    // Формируем тело запроса только с переданными параметрами
    const requestBody: any = {};
    if (body.archived !== undefined) requestBody.archived = body.archived;
    if (body.hasStocks !== undefined) requestBody.hasStocks = body.hasStocks;
    if (body.offerIds) requestBody.offerIds = body.offerIds;
    if (body.stocksWarehouseId) requestBody.stocksWarehouseId = body.stocksWarehouseId;
    if (body.withTurnover !== undefined) requestBody.withTurnover = body.withTurnover;

    const response = await fetch(url.toString(), {
      method: 'POST',
      headers: {
        'Api-Key': token,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    });

    const data = await response.json();

    if (!response.ok) {
      return NextResponse.json(
        { error: 'Yandex API error', details: data },
        { status: response.status }
      );
    }

    return NextResponse.json(data);
  } catch (error) {
    console.error('Error in stocks route:', error);
    return NextResponse.json(
      { error: 'Internal server error' },
      { status: 500 }
    );
  }
}


