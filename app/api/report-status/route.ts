import { NextRequest, NextResponse } from 'next/server';

export async function GET(request: NextRequest) {
  const apiKey = request.headers.get('x-api-key');
  const reportId = request.nextUrl.searchParams.get('reportId');
  
  if (!apiKey || !reportId) {
    return NextResponse.json(
      { error: 'API ключ или reportId не предоставлен' }, 
      { status: 400 }
    );
  }

  try {
    const response = await fetch(`https://api.partner.market.yandex.ru/v2/reports/info/${reportId}`, {
      method: 'GET',
      headers: {
        'Api-Key': apiKey,
        'Content-Type': 'application/json',
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      return NextResponse.json(
        { error: `API Error: ${response.status} - ${errorText}` },
        { status: response.status }
      );
    }

    const data = await response.json();
    return NextResponse.json(data);

  } catch (error) {
    console.error('Ошибка при запросе статуса отчета:', error);
    return NextResponse.json(
      { error: 'Ошибка при запросе к API' },
      { status: 500 }
    );
  }
}
