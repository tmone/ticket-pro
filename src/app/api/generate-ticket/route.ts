import { NextRequest, NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';
import QRCode from 'qrcode';

export async function POST(request: NextRequest) {
  try {
    const { qrData, rowData } = await request.json();

    if (!qrData) {
      return NextResponse.json(
        { error: 'QR data is required' },
        { status: 400 }
      );
    }

    // Check if this is a VIP ticket
    let isVip = false;
    if (rowData) {
      const vipColumnKey = Object.keys(rowData).find(key => 
        key.toLowerCase() === 'vip'
      );
      
      if (vipColumnKey) {
        const vipValue = rowData[vipColumnKey];
        isVip = vipValue && (
          vipValue.toString() === '1' || 
          vipValue.toString().toUpperCase() === 'X' || 
          vipValue.toString().toLowerCase() === 'yes'
        );
      }
    }

    // Read the appropriate template
    const svgPath = path.join(process.cwd(), 'public', isVip ? 'ticket_vip.svg' : 'ticket.svg');
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

    // Handle QR code placement based on template type
    if (isVip) {
      // VIP template uses different coordinates
      const vipScale = 226.72 / 29; // Scale factor for VIP QR code
      const vipQrCodeGroup = `
        <g transform="translate(1657.38, 787.36)">
          <g transform="scale(${vipScale})">
            ${qrCodeElements}
          </g>
        </g>
      `;
      svgContent = svgContent.replace(
        /<rect[^>]*id="QR-CODE"[^>]*\/>/,
        vipQrCodeGroup
      );
      
      // Replace NAME and TITLE for VIP template
      if (rowData) {
        // Look for NAME column
        const nameColumnKey = Object.keys(rowData).find(key => 
          key.toLowerCase() === 'name' || key.toLowerCase() === 'tên' || key.toLowerCase() === 'họ tên'
        );
        if (nameColumnKey) {
          const name = rowData[nameColumnKey] || '';
          svgContent = svgContent.replace(
            /(<text[^>]*id="NAME"[^>]*?)([^>]*>)/,
            `$1 text-anchor="middle"$2`
          );
          svgContent = svgContent.replace(
            /(<text[^>]*id="NAME"[^>]*>[\s\S]*?<tspan[^>]*x=")[^"]*("[\s\S]*?>)[^<]*([\s\S]*?<\/tspan>[\s\S]*?<\/text>)/,
            `$1320$2${name}$3`
          );
        }
        
        // Look for TITLE column
        const titleColumnKey = Object.keys(rowData).find(key => 
          key.toLowerCase() === 'title' || key.toLowerCase() === 'chức vụ' || key.toLowerCase() === 'position'
        );
        if (titleColumnKey) {
          const title = rowData[titleColumnKey] || '';
          svgContent = svgContent.replace(
            /(<text[^>]*id="TITLE"[^>]*?)([^>]*>)/,
            `$1 text-anchor="middle"$2`
          );
          svgContent = svgContent.replace(
            /(<text[^>]*id="TITLE"[^>]*>[\s\S]*?<tspan[^>]*x=")[^"]*("[\s\S]*?>)[^<]*([\s\S]*?<\/tspan>[\s\S]*?<\/text>)/,
            `$1320$2${title}$3`
          );
        }
      }
    } else {
      // Regular template
      const scale = 368 / 29; // Scale factor to make QR code fill the entire area
      const qrCodeGroup = `
        <g transform="translate(1200, 284)">
          <g transform="scale(${scale})">
            ${qrCodeElements}
          </g>
        </g>
      `;
      svgContent = svgContent.replace(
        /<rect x="1200" y="284" width="368" height="368" stroke="white" stroke-width="0" id="qr-code"\/>/,
        qrCodeGroup
      );
    }

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