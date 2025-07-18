import { NextRequest, NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';
import QRCode from 'qrcode';
import sharp from 'sharp';
import JSZip from 'jszip';

// Extend global interface
declare global {
  var ticketProgress: Record<string, { current: number; total: number; completed: boolean }> | undefined;
}

export async function POST(request: NextRequest) {
  try {
    const { tickets, sessionId } = await request.json();

    if (!tickets || !Array.isArray(tickets) || tickets.length === 0) {
      return NextResponse.json(
        { error: 'Tickets array is required' },
        { status: 400 }
      );
    }

    // Simple in-memory progress store for this request
    let progressState = { current: 0, total: tickets.length, completed: false };
    global.ticketProgress = global.ticketProgress || {};
    global.ticketProgress[sessionId] = progressState;

    // Update progress function
    const updateProgress = (current: number, total: number, completed = false) => {
      if (sessionId && global.ticketProgress) {
        global.ticketProgress[sessionId] = { current, total, completed };
      }
    };

    // Read the ticket.svg template
    const svgPath = path.join(process.cwd(), 'public', 'ticket.svg');
    const svgTemplate = fs.readFileSync(svgPath, 'utf-8');

    // Create a new JSZip instance
    const zip = new JSZip();

    // Initialize progress
    updateProgress(0, tickets.length);

    // Process tickets in batches for better performance
    const batchSize = 5;
    let processedCount = 0;
    
    for (let i = 0; i < tickets.length; i += batchSize) {
      const batch = tickets.slice(i, i + batchSize);
      
      await Promise.all(batch.map(async (ticket) => {
        try {
          // Generate QR code SVG locally
          const qrSvgContent = await QRCode.toString(ticket.qrData, {
            type: 'svg',
            width: 368,
            margin: 0,
            color: {
              dark: '#000000',
              light: '#FFFFFF'
            }
          });
          
          // Extract QR code elements
          const qrCodeElements = qrSvgContent
            .replace(/<\?xml[^>]*\?>/, '')
            .replace(/<svg[^>]*>/, '')
            .replace(/<\/svg>/, '')
            .trim();

          // Create QR code group with proper scaling
          const scale = 368 / 29; // Scale factor to make QR code fill the entire area
          const qrCodeGroup = `
            <g transform="translate(1200, 284)">
              <g transform="scale(${scale})">
                ${qrCodeElements}
              </g>
            </g>
          `;

          // Replace placeholder with QR code in SVG
          let ticketSvg = svgTemplate.replace(
            /<rect x="1200" y="284" width="368" height="368" stroke="white" stroke-width="0" id="qr-code"\/>/,
            qrCodeGroup
          );

          // Convert SVG to JPG using Sharp with optimized settings
          const jpgBuffer = await sharp(Buffer.from(ticketSvg))
            .jpeg({ 
              quality: 75,
              density: 72,
              progressive: true
            })
            .toBuffer();

          // Add to ZIP with sequential naming
          const filename = `${ticket.rowNumber}.jpg`;
          zip.file(filename, jpgBuffer);
          console.log(`Added ${filename} to ZIP (${jpgBuffer.length} bytes)`);

          // Update progress
          processedCount++;
          updateProgress(processedCount, tickets.length);

        } catch (ticketError) {
          console.error(`Error processing ticket ${ticket.rowNumber}:`, ticketError);
          // Continue with other tickets even if one fails
          processedCount++;
          updateProgress(processedCount, tickets.length);
        }
      }));
    }

    // Generate ZIP buffer
    console.log('Generating ZIP buffer...');
    const zipBuffer = await zip.generateAsync({ 
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });

    console.log(`Generated ZIP file with ${zipBuffer.length} bytes`);

    // Mark as completed
    updateProgress(tickets.length, tickets.length, true);

    return new NextResponse(zipBuffer, {
      headers: {
        'Content-Type': 'application/zip',
        'Content-Disposition': `attachment; filename="tickets-${Date.now()}.zip"`,
        'Content-Length': zipBuffer.length.toString(),
      },
    });

  } catch (error) {
    console.error('Generate tickets ZIP error:', error);
    // Clean up progress on error
    if (sessionId && global.ticketProgress?.[sessionId]) {
      delete global.ticketProgress[sessionId];
    }
    return NextResponse.json(
      { error: 'Failed to generate tickets ZIP', details: error instanceof Error ? error.message : String(error) },
      { status: 500 }
    );
  }
}