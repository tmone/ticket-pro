import { NextRequest, NextResponse } from 'next/server';
const nodemailer = require('nodemailer');
import fs from 'fs';
import path from 'path';
import QRCode from 'qrcode';
import sharp from 'sharp';
import { google } from 'googleapis';

// Function to update Google Sheets email status
async function updateGoogleSheetsEmailStatus(
  spreadsheetId: string,
  sheetName: string,
  rowIndices: number[],
  emailSentColumnName: string
) {
  console.log('=== Starting Google Sheets Email Status Update ===');
  console.log('Parameters:', {
    spreadsheetId,
    sheetName,
    rowIndices,
    emailSentColumnName
  });
  
  // Create service account auth
  console.log('Service account config for email status update:', {
    hasEmail: !!process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    hasPrivateKey: !!process.env.GOOGLE_PRIVATE_KEY,
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
  });
  
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const sheets = google.sheets({ version: 'v4', auth });

  // Get sheet metadata to find the column index
  const metadataResponse = await sheets.spreadsheets.get({
    spreadsheetId,
  });

  const sheet = metadataResponse.data.sheets?.find(
    s => s.properties?.title === sheetName
  );

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }

  // Get current headers to find the email sent column
  const headersResponse = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!1:1`,
  });

  const headers = headersResponse.data.values?.[0] as string[];
  console.log('Current headers:', headers);
  
  let emailSentColumnIndex = headers?.findIndex(h => 
    h?.toLowerCase().includes('email sent') || 
    h?.toLowerCase().includes('email_sent') ||
    h === emailSentColumnName
  );
  
  console.log('Email sent column search:', {
    emailSentColumnName,
    foundIndex: emailSentColumnIndex,
    headers: headers?.map((h, i) => `${i}: ${h}`)
  });

  // If column doesn't exist, add it
  if (emailSentColumnIndex === -1) {
    headers?.push(emailSentColumnName);
    emailSentColumnIndex = headers.length - 1;
    
    // Update headers
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!1:1`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [headers],
      },
    });
  }

  // Convert column index to letter (A, B, C, ..., Z, AA, AB, etc.)
  const getColumnLetter = (index: number): string => {
    let letter = '';
    while (index >= 0) {
      letter = String.fromCharCode(65 + (index % 26)) + letter;
      index = Math.floor(index / 26) - 1;
    }
    return letter;
  };
  const columnLetter = getColumnLetter(emailSentColumnIndex);

  // Update each successful row
  const timestamp = new Date().toLocaleString('vi-VN');
  const updates = rowIndices.map(rowIndex => ({
    range: `${sheetName}!${columnLetter}${rowIndex + 2}`, // +2 because row indices are 0-based and we skip header
    values: [[timestamp]],
  }));
  
  console.log('Preparing updates:', {
    columnLetter,
    rowIndices,
    updates: updates.slice(0, 3), // Show first 3 for debugging
    timestamp
  });

  // Batch update all rows
  try {
    const updateResponse = await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'RAW',
        data: updates,
      },
    });
    console.log('Google Sheets update response:', {
      updatedCells: updateResponse.data.totalUpdatedCells,
      updatedRows: updateResponse.data.totalUpdatedRows,
      updatedColumns: updateResponse.data.totalUpdatedColumns,
    });
  } catch (updateError: any) {
    console.error('Google Sheets update error:', {
      message: updateError.message,
      code: updateError.code,
      errors: updateError.errors,
      status: updateError.status,
    });
    throw updateError;
  }
}

// Function to update Google Sheets email error status
async function updateGoogleSheetsEmailErrors(
  spreadsheetId: string,
  sheetName: string,
  failedRows: { rowIndex: number; error: string }[],
  emailErrorColumnName: string
) {
  // Create service account auth
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const sheets = google.sheets({ version: 'v4', auth });

  // Get current headers to find the email error column
  const headersResponse = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!1:1`,
  });

  const headers = headersResponse.data.values?.[0] as string[];
  let emailErrorColumnIndex = headers?.findIndex(h => 
    h?.toLowerCase().includes('email error') || 
    h?.toLowerCase().includes('email_error') ||
    h === emailErrorColumnName
  );

  // If column doesn't exist, add it
  if (emailErrorColumnIndex === -1) {
    headers?.push(emailErrorColumnName);
    emailErrorColumnIndex = headers.length - 1;
    
    // Update headers
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!1:1`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [headers],
      },
    });
  }

  // Convert column index to letter (A, B, C, ..., Z, AA, AB, etc.)
  const getColumnLetter = (index: number): string => {
    let letter = '';
    while (index >= 0) {
      letter = String.fromCharCode(65 + (index % 26)) + letter;
      index = Math.floor(index / 26) - 1;
    }
    return letter;
  };
  const columnLetter = getColumnLetter(emailErrorColumnIndex);

  // Update each failed row with error message
  const timestamp = new Date().toLocaleString('vi-VN');
  const updates = failedRows.map(failedRow => ({
    range: `${sheetName}!${columnLetter}${failedRow.rowIndex + 2}`, // +2 because row indices are 0-based and we skip header
    values: [[`${timestamp}: ${failedRow.error}`]],
  }));

  // Batch update all rows
  try {
    const updateResponse = await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'RAW',
        data: updates,
      },
    });
    console.log('Google Sheets update response:', {
      updatedCells: updateResponse.data.totalUpdatedCells,
      updatedRows: updateResponse.data.totalUpdatedRows,
      updatedColumns: updateResponse.data.totalUpdatedColumns,
    });
  } catch (updateError: any) {
    console.error('Google Sheets update error:', {
      message: updateError.message,
      code: updateError.code,
      errors: updateError.errors,
      status: updateError.status,
    });
    throw updateError;
  }
}

interface EmailData {
  email: string;
  name?: string;
  qrData: string;
  rowNumber: string;
  originalRowIndex?: number; // Index in the original data for updating
  rowData?: Record<string, any>; // Full row data for template placeholders
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
      appendTicketInline,
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
      appendTicketInline?: boolean;
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

    const svgTemplate = (attachTicket || appendTicketInline) ? fs.readFileSync(path.join(process.cwd(), 'public', 'ticket.svg'), 'utf-8') : null;
    
    // Load email template
    const emailTemplatePath = path.join(process.cwd(), 'public', 'templates', 'email.eml');
    let emailTemplate = '';
    try {
      const templateContent = fs.readFileSync(emailTemplatePath, 'utf-8');
      
      // Parse to remove subject line from message body
      const lines = templateContent.split('\n');
      let messageStart = 0;
      
      for (let i = 0; i < lines.length; i++) {
        if (lines[i].startsWith('Subject: ')) {
          messageStart = i + 1;
          break;
        }
      }
      
      // Skip empty lines after subject
      while (messageStart < lines.length && lines[messageStart].trim() === '') {
        messageStart++;
      }
      
      // Get message content only (no subject)
      emailTemplate = lines.slice(messageStart).join('\n').trim();
      
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
    const successfulEmailRows: number[] = []; // Track rows that need to be updated
    const failedEmailRows: { rowIndex: number; error: string }[] = []; // Track failed rows with error messages

    // Function to validate email address
    const validateEmail = (email: string): boolean => {
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      return emailRegex.test(email);
    };

    // Helper function to delay between emails
    const delay = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));
    
    // Send emails one by one to avoid rate limiting
    for (let i = 0; i < emails.length; i++) {
      const emailData = emails[i];
      
      // Add delay between emails (except for the first one)
      if (i > 0) {
        // Progressive delay: 1-2 seconds for first 10, 2-3 seconds for next 10, etc.
        const baseDelay = Math.min(Math.floor(i / 10) + 1, 5) * 1000; // Max 5 seconds
        const randomDelay = Math.random() * 1000; // 0-1 second random
        const totalDelay = baseDelay + randomDelay;
        
        console.log(`Delaying ${totalDelay}ms before sending email ${i + 1}/${emails.length}`);
        await delay(totalDelay);
      }
      
      try {
        // Validate email address first
        if (!validateEmail(emailData.email)) {
          throw new Error('Invalid email format');
        }
        let attachments = [];
        let jpgBuffer: Buffer | null = null;
        
        console.log('Ticket generation check:', {
          attachTicket,
          appendTicketInline,
          hasTemplate: !!svgTemplate,
          hasQrData: !!emailData.qrData,
          shouldGenerate: (attachTicket || appendTicketInline) && svgTemplate && emailData.qrData
        });

        if ((attachTicket || appendTicketInline) && svgTemplate && emailData.qrData) {
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
          jpgBuffer = await sharp(Buffer.from(ticketSvg))
            .jpeg({ 
              quality: 75,
              density: 72,
              progressive: true
            })
            .toBuffer();

          // Create attachment
          const attachment: any = {
            filename: `ticket-${emailData.rowNumber}.jpg`,
            content: jpgBuffer,
            contentType: 'image/jpeg'
          };
          
          // If appendTicketInline, add CID for inline display
          if (appendTicketInline) {
            attachment.cid = `ticket-${emailData.rowNumber}`;
            attachment.contentDisposition = 'inline';
          }
          
          attachments.push(attachment);
          console.log('Created attachment:', {
            filename: attachment.filename,
            hasContent: !!attachment.content,
            contentLength: attachment.content?.length,
            cid: attachment.cid,
            contentDisposition: attachment.contentDisposition
          });
        }

        // Use template if available, otherwise use provided message
        const messageToUse = emailTemplate || message;
        
        // Function to replace placeholders with row data
        const replacePlaceholders = (text: string): string => {
          let result = text;
          
          // Debug logging
          console.log('Replacing placeholders for email:', emailData.email);
          console.log('Row data:', emailData.rowData);
          
          // Replace ALL column-based placeholders from rowData
          if (emailData.rowData) {
            Object.keys(emailData.rowData).forEach(columnName => {
              // Create regex to match {columnName} in any case
              const placeholder = new RegExp(`\\{${columnName}\\}`, 'gi');
              const value = emailData.rowData![columnName];
              console.log(`Replacing {${columnName}} with "${value}"`);
              result = result.replace(placeholder, value?.toString() || '');
            });
          }
          
          // System placeholders (only replace if not already replaced by column data)
          result = result
            .replace(/\{senderName\}/gi, defaultSenderName)
            .replace(/\{senderEmail\}/gi, defaultSenderEmail)
            .replace(/\{ticketCode\}/gi, emailData.qrData)
            .replace(/\{eventDate\}/gi, new Date().toLocaleDateString('vi-VN'))
            .replace(/\{contactPhone\}/gi, process.env.CONTACT_PHONE || '0123456789');
            
          return result;
        };
        
        // Replace placeholders in both message and subject
        const personalizedMessage = replacePlaceholders(messageToUse);
        const personalizedSubject = replacePlaceholders(subject);
        
        // Prepare HTML content
        // Check if message is already HTML (from rich text editor)
        let htmlContent = personalizedMessage;
        if (!personalizedMessage.includes('<') || !personalizedMessage.includes('>')) {
          // Plain text - convert newlines to <br>
          htmlContent = personalizedMessage.replace(/\n/g, '<br>');
        }
        
        // If appendTicketInline is true, add the image as inline at the bottom
        if (appendTicketInline && jpgBuffer) {
          // Add the image inline at the bottom with 100% width
          htmlContent += `<br><br><img src="cid:ticket-${emailData.rowNumber}" style="width: 100%; max-width: 800px; height: auto; display: block; margin: 0 auto;" alt="Event Ticket" />`;
        }

        const mailOptions = {
          from: defaultSenderName ? `"${defaultSenderName}" <${defaultSenderEmail}>` : defaultSenderEmail,
          to: emailData.email,
          bcc: bccList.length > 0 ? bccList : undefined,
          subject: personalizedSubject,
          text: personalizedMessage,
          html: htmlContent,
          attachments: attachments
        };
        
        console.log('Mail options:', {
          to: emailData.email,
          hasAttachments: attachments.length > 0,
          attachmentCount: attachments.length,
          htmlContainsImage: htmlContent.includes('<img'),
          appendTicketInline
        });

        await transporter.sendMail(mailOptions);
        successCount++;
        console.log(`Email sent successfully to ${emailData.email} (${i + 1}/${emails.length})`);
        
        // Track successful email for Google Sheets update
        if (emailData.originalRowIndex !== undefined) {
          console.log(`Tracking row ${emailData.originalRowIndex} for update`);
          successfulEmailRows.push(emailData.originalRowIndex);
        } else {
          console.log(`Warning: No originalRowIndex for ${emailData.email}`);
        }

      } catch (emailError) {
        failureCount++;
        const errorMsg = emailError instanceof Error ? emailError.message : String(emailError);
        const fullErrorMsg = `Failed to send to ${emailData.email}: ${errorMsg}`;
        errors.push(fullErrorMsg);
        console.error(fullErrorMsg);
        
        // Track failed email for Google Sheets update
        if (emailData.originalRowIndex !== undefined) {
          failedEmailRows.push({
            rowIndex: emailData.originalRowIndex,
            error: errorMsg
          });
        }
      }
    }

    // Update Google Sheets to mark emails as sent
    if ((successfulEmailRows.length > 0 || failedEmailRows.length > 0) && spreadsheetId && sheetName) {
      try {
        console.log('Attempting to update Google Sheets:', {
          spreadsheetId,
          sheetName,
          successfulRows: successfulEmailRows,
          failedRows: failedEmailRows.map(f => f.rowIndex),
          emailSentColumnName: emailSentColumnName || 'Email Sent'
        });
        
        if (successfulEmailRows.length > 0) {
          await updateGoogleSheetsEmailStatus(
            spreadsheetId, 
            sheetName, 
            successfulEmailRows, 
            emailSentColumnName || 'Email Sent'
          );
          console.log(`Updated ${successfulEmailRows.length} successful email rows in Google Sheets`);
        }
        
        if (failedEmailRows.length > 0) {
          await updateGoogleSheetsEmailErrors(
            spreadsheetId, 
            sheetName, 
            failedEmailRows,
            'Email Error'
          );
          console.log(`Updated ${failedEmailRows.length} failed email rows in Google Sheets`);
        }
      } catch (updateError) {
        console.error('Failed to update Google Sheets:', updateError);
        // Don't fail the entire operation if sheet update fails
      }
    }

    return NextResponse.json({
      success: true,
      message: `Sent ${successCount} emails successfully${failureCount > 0 ? `, ${failureCount} failed` : ''}`,
      successCount,
      failureCount,
      updatedSuccessRows: successfulEmailRows.length,
      updatedErrorRows: failedEmailRows.length,
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