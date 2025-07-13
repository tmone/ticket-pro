import { NextRequest, NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';
import QRCode from 'qrcode';

export async function POST(request: NextRequest) {
  try {
    const { qrData } = await request.json();

    if (!qrData) {
      return NextResponse.json(
        { error: 'QR data is required' },
        { status: 400 }
      );
    }

    // Read the ticket.svg template
    const svgPath = path.join(process.cwd(), 'public', 'ticket.svg');
    let svgContent = fs.readFileSync(svgPath, 'utf-8');

    // Generate QR code SVG locally using qrcode package
    const qrSvgContent = await QRCode.toString(qrData, {
      type: 'svg',
      width: 368,
      margin: 0,
      color: {
        dark: '#000000',
        light: '#FFFFFF'
      }
    });
    
    // Extract the QR code paths/elements from the SVG (remove svg wrapper)
    const qrCodeElements = qrSvgContent
      .replace(/<\?xml[^>]*\?>/, '')
      .replace(/<svg[^>]*>/, '')
      .replace(/<\/svg>/, '')
      .trim();

    // Create a group element for the QR code centered in the specified area
    // Area is 368x368 at position (1200, 284), QR code is 29x29 units, scale to fit 368x368
    const scale = 368 / 29; // Scale factor to make QR code fill the entire area
    const qrCodeGroup = `
      <g transform="translate(1200, 284)">
        <g transform="scale(${scale})">
          ${qrCodeElements}
        </g>
      </g>
    `;

    // Replace the placeholder rect with the actual QR code
    svgContent = svgContent.replace(
      /<rect x="1200" y="284" width="368" height="368" stroke="white" stroke-width="0" id="qr-code"\/>/,
      qrCodeGroup
    );

    return new NextResponse(svgContent, {
      headers: {
        'Content-Type': 'image/svg+xml',
        'Content-Disposition': `attachment; filename="ticket-${Date.now()}.svg"`,
      },
    });

  } catch (error) {
    console.error('Generate ticket error:', error);
    return NextResponse.json(
      { error: 'Failed to generate ticket' },
      { status: 500 }
    );
  }
}