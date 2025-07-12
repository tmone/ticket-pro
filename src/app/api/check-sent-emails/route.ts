import { NextRequest, NextResponse } from 'next/server';
import { google } from 'googleapis';

export async function POST(request: NextRequest) {
  try {
    const { spreadsheetId, sheetName, emailAddresses, emailColumn } = await request.json();

    if (!spreadsheetId || !sheetName || !emailAddresses || !Array.isArray(emailAddresses)) {
      return NextResponse.json(
        { error: 'Missing required parameters' },
        { status: 400 }
      );
    }

    // Create service account auth
    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      },
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });

    const sheets = google.sheets({ version: 'v4', auth });

    // Get all data from the sheet
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:ZZ`, // Get all columns
    });

    const rows = response.data.values;
    if (!rows || rows.length < 2) {
      // No data or only headers
      return NextResponse.json({ previouslySentEmails: [] });
    }

    const headers = rows[0] as string[];
    const dataRows = rows.slice(1);
    
    console.log('=== Check Sent Emails Debug ===');
    console.log('Headers:', headers);
    console.log('Email addresses to check:', emailAddresses);

    // Find email column - use provided column name or auto-detect
    let emailColumnIndex = -1;
    if (emailColumn) {
      emailColumnIndex = headers.findIndex(h => h === emailColumn);
      console.log(`Using provided email column: ${emailColumn} at index ${emailColumnIndex}`);
    } else {
      emailColumnIndex = headers.findIndex(h => {
        const normalized = h?.toLowerCase().replace(/[^a-z]/g, ''); // Remove non-letters
        return (normalized?.includes('email') || normalized?.includes('mail')) && 
          !h?.toLowerCase().includes('sent') && 
          !h?.toLowerCase().includes('error');
      });
      console.log(`Auto-detected email column at index ${emailColumnIndex}`);
    }

    const emailSentColumnIndex = headers.findIndex(h => 
      h?.toLowerCase().includes('email sent') || 
      h?.toLowerCase().includes('email_sent') ||
      h === 'Email Sent'
    );
    
    console.log('Email column index:', emailColumnIndex, emailColumnIndex >= 0 ? `(${headers[emailColumnIndex]})` : '');
    console.log('Email sent column index:', emailSentColumnIndex, emailSentColumnIndex >= 0 ? `(${headers[emailSentColumnIndex]})` : '');
    
    // Debug: Show sample data from first few rows
    if (dataRows.length > 0) {
      console.log('Sample data from first row:');
      dataRows.slice(0, 2).forEach((row, idx) => {
        console.log(`Row ${idx + 2}:`, {
          email: emailColumnIndex >= 0 ? row[emailColumnIndex] : 'N/A',
          emailSent: emailSentColumnIndex >= 0 ? row[emailSentColumnIndex] : 'N/A',
          fullRow: row
        });
      });
    }

    if (emailColumnIndex === -1) {
      return NextResponse.json({ 
        error: 'Email column not found',
        previouslySentEmails: [] 
      });
    }

    // If no email sent column exists, no emails have been sent
    if (emailSentColumnIndex === -1) {
      return NextResponse.json({ previouslySentEmails: [] });
    }

    // Check which of the provided email addresses have been sent
    const previouslySentEmails: string[] = [];

    for (const emailAddress of emailAddresses) {
      console.log(`Checking email: ${emailAddress}`);
      for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        const rowEmail = row[emailColumnIndex]?.toString().toLowerCase().trim();
        const emailSentValue = row[emailSentColumnIndex]?.toString().trim();
        
        if (i < 3) { // Log first 3 rows for debugging
          console.log(`Row ${i + 2}: email="${rowEmail}", sent="${emailSentValue}"`);
        }
        
        if (rowEmail === emailAddress.toLowerCase().trim() && emailSentValue) {
          // This email has been sent before (has a value in Email Sent column)
          console.log(`Found previously sent email: ${emailAddress} (sent value: ${emailSentValue})`);
          previouslySentEmails.push(emailAddress);
          break; // Found this email, no need to check other rows
        }
      }
    }

    return NextResponse.json({
      previouslySentEmails: [...new Set(previouslySentEmails)] // Remove duplicates
    });

  } catch (error) {
    console.error('Check sent emails error:', error);
    return NextResponse.json(
      { 
        error: 'Failed to check sent emails',
        details: error instanceof Error ? error.message : String(error),
        previouslySentEmails: [] // Return empty array on error to allow sending
      },
      { status: 500 }
    );
  }
}