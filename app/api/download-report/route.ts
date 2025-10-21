import { NextRequest, NextResponse } from 'next/server';

export async function GET(request: NextRequest) {
  const apiKey = request.headers.get('x-api-key');
  const downloadUrl = request.nextUrl.searchParams.get('url');
  
  if (!apiKey || !downloadUrl) {
    return NextResponse.json(
      { error: 'API ключ или URL не предоставлен' }, 
      { status: 400 }
    );
  }

  try {
    const response = await fetch(downloadUrl, {
      method: 'GET',
      headers: {
        'Api-Key': apiKey,
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      return NextResponse.json(
        { error: `Download Error: ${response.status} - ${errorText}` },
        { status: response.status }
      );
    }

    // Получаем Excel файл как ArrayBuffer
    const arrayBuffer = await response.arrayBuffer();
    
    // Возвращаем файл с правильными заголовками
    return new NextResponse(arrayBuffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="yandex_market_realization_report.xlsx"',
      },
    });

  } catch (error) {
    console.error('Ошибка при скачивании отчета:', error);
    return NextResponse.json(
      { error: 'Ошибка при скачивании файла' },
      { status: 500 }
    );
  }
}
