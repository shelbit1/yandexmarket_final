import { NextRequest, NextResponse } from 'next/server';

export async function PUT(request: NextRequest) {
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
    
    const url = `https://api.partner.market.yandex.ru/v2/campaigns/${campaignId}/offers/stocks`;

    const response = await fetch(url, {
      method: 'PUT',
      headers: {
        'Api-Key': token,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
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
    console.error('Error in stocks-update route:', error);
    return NextResponse.json(
      { error: 'Internal server error' },
      { status: 500 }
    );
  }
}


