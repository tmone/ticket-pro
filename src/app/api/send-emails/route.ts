import { NextRequest, NextResponse } from 'next/server';
const nodemailer = require('nodemailer');
import fs from 'fs';
import path from 'path';
import QRCode from 'qrcode';
import sharp from 'sharp';

interface EmailData {
  email: string;
  name?: string;
  qrData: string;
  rowNumber: string;
  originalRowIndex?: number; // Index in the original data for updating
}

export async function POST(request: NextRequest) {
  try {
    const { 
      emails, 
      subject, 
      message, 
      senderEmail, 
      senderName, 
      attachTicket,
      spreadsheetId,
      sheetName,
      emailColumnIndex,
      emailSentColumnName
    }: {
      emails: EmailData[];
      subject: string;
      message: string;
      senderEmail: string;
      senderName?: string;
      attachTicket: boolean;
      spreadsheetId?: string;
      sheetName?: string;
      emailColumnIndex?: number;
      emailSentColumnName?: string;
    } = await request.json();

    if (!emails || !Array.isArray(emails) || emails.length === 0) {
      return NextResponse.json(
        { error: 'Emails array is required' },
        { status: 400 }
      );
    }

    if (!subject || !message || !senderEmail) {
      return NextResponse.json(
        { error: 'Subject, message, and sender email are required' },
        { status: 400 }
      );
    }

    // Create nodemailer transporter with custom SMTP server
    const transporter = nodemailer.createTransport({
      host: process.env.EMAIL_SERVER,
      port: 587, // or 465 for SSL
      secure: false, // true for 465, false for other ports
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_APP_PASSWORD,
      },
      tls: {
        rejectUnauthorized: false // Allow self-signed certificates
      }
    });

    const svgTemplate = attachTicket ? fs.readFileSync(path.join(process.cwd(), 'public', 'ticket.svg'), 'utf-8') : null;
    
    // Load email template
    const emailTemplatePath = path.join(process.cwd(), 'public', 'templates', 'email.eml');
    let emailTemplate = '';
    try {
      emailTemplate = fs.readFileSync(emailTemplatePath, 'utf-8');
    } catch (error) {
      // Fallback to provided message if template doesn't exist
      emailTemplate = message;
    }
    
    // Use environment sender info as defaults
    const defaultSenderName = process.env.SENDER_NAME || senderName || 'Event Team';
    const defaultSenderEmail = process.env.SENDER_EMAIL || process.env.EMAIL_USER || senderEmail;
    
    // Parse BCC list from environment
    const bccList = process.env.SENDER_BCC 
      ? process.env.SENDER_BCC.replace(/"/g, '').split(';').filter(email => email.trim())
      : [];
    
    let successCount = 0;
    let failureCount = 0;
    const errors: string[] = [];

    // Send emails one by one to avoid rate limiting
    for (const emailData of emails) {
      try {
        let attachments = [];

        if (attachTicket && svgTemplate && emailData.qrData) {
          // Generate QR code and ticket
          const qrSvgContent = await QRCode.toString(emailData.qrData, {
            type: 'svg',
            width: 368,
            margin: 0,
            color: {
              dark: '#000000',
              light: '#FFFFFF'
            }
          });
          
          const qrCodeElements = qrSvgContent
            .replace(/<\?xml[^>]*\?>/, '')
            .replace(/<svg[^>]*>/, '')
            .replace(/<\/svg>/, '')
            .trim();

          const scale = 368 / 29;
          const qrCodeGroup = `
            <g transform="translate(1200, 284)">
              <g transform="scale(${scale})">
                ${qrCodeElements}
              </g>
            </g>
          `;

          const ticketSvg = svgTemplate.replace(
            /<rect x="1200" y="284" width="368" height="368" stroke="white" stroke-width="0" id="qr-code"\/>/,
            qrCodeGroup
          );

          // Convert to JPG
          const jpgBuffer = await sharp(Buffer.from(ticketSvg))
            .jpeg({ 
              quality: 75,
              density: 72,
              progressive: true
            })
            .toBuffer();

          attachments.push({
            filename: `ticket-${emailData.rowNumber}.jpg`,
            content: jpgBuffer,
            contentType: 'image/jpeg'
          });
        }

        // Use template if available, otherwise use provided message
        const messageToUse = emailTemplate || message;
        
        // Replace all placeholders in message
        const personalizedMessage = messageToUse
          .replace(/\{name\}/g, emailData.name || 'Quý khách')
          .replace(/\{senderName\}/g, defaultSenderName)
          .replace(/\{senderEmail\}/g, defaultSenderEmail)
          .replace(/\{ticketCode\}/g, emailData.qrData)
          .replace(/\{eventName\}/g, subject.replace('Vé tham dự sự kiện - ', '') || 'Sự kiện')
          .replace(/\{eventDate\}/g, new Date().toLocaleDateString('vi-VN'))
          .replace(/\{eventLocation\}/g, 'Theo thông tin trong vé')
          .replace(/\{contactPhone\}/g, process.env.CONTACT_PHONE || '0123456789');

        const mailOptions = {
          from: defaultSenderName ? `"${defaultSenderName}" <${defaultSenderEmail}>` : defaultSenderEmail,
          to: emailData.email,
          bcc: bccList.length > 0 ? bccList : undefined,
          subject: subject,
          text: personalizedMessage,
          html: personalizedMessage.replace(/\n/g, '<br>'),
          attachments: attachments
        };

        await transporter.sendMail(mailOptions);
        successCount++;
        console.log(`Email sent successfully to ${emailData.email}`);

      } catch (emailError) {
        failureCount++;
        const errorMsg = `Failed to send to ${emailData.email}: ${emailError instanceof Error ? emailError.message : String(emailError)}`;
        errors.push(errorMsg);
        console.error(errorMsg);
      }
    }

    return NextResponse.json({
      success: true,
      message: `Sent ${successCount} emails successfully${failureCount > 0 ? `, ${failureCount} failed` : ''}`,
      successCount,
      failureCount,
      errors: errors.length > 0 ? errors.slice(0, 5) : undefined // Limit error details
    });

  } catch (error) {
    console.error('Send emails error:', error);
    return NextResponse.json(
      { 
        error: 'Failed to send emails', 
        details: error instanceof Error ? error.message : String(error) 
      },
      { status: 500 }
    );
  }
}