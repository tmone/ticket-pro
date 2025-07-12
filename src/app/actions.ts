'use server';

import { getIronSession } from 'iron-session';
import { cookies } from 'next/headers';
import { google } from 'googleapis';
import type { SessionData } from '@/lib/session';
import { sessionOptions } from '@/lib/session';

// Helper function to convert array of arrays to array of objects
const toJsonObject = (data: any[][]): Record<string, any>[] => {
  if (!data || data.length < 1) return [];

  const headers = data[0].map(h => String(h).trim());
  const jsonData = [];

  for (let i = 1; i < data.length; i++) {
    const rowData = data[i];
    const rowObject: Record<string, any> = {};
    for (let j = 0; j < headers.length; j++) {
      rowObject[headers[j]] = rowData[j] ?? '';
    }
    jsonData.push(rowObject);
  }

  return jsonData;
};

export async function fetchGoogleSheetData(sheetUrl: string): Promise<{ data?: Record<string, any>[]; error?: string }> {
  const session = await getIronSession<SessionData>(cookies(), sessionOptions);

  if (!session.isLoggedIn || !session.tokens) {
    return { error: 'Authentication required. Please log in.' };
  }

  if (!sheetUrl) {
    return { error: 'Google Sheet URL is required.' };
  }

  try {
    const oauth2Client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET,
      process.env.GOOGLE_REDIRECT_URI
    );
    oauth2Client.setCredentials(session.tokens);

    const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

    const match = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/.exec(sheetUrl);
    if (!match || !match[1]) {
      throw new Error('Invalid Google Sheet URL. Could not find a valid Sheet ID.');
    }
    const spreadsheetId = match[1];
    
    // To keep it simple, we'll fetch the first sheet's data.
    // A more advanced version could let the user select a sheet.
    const sheetMetadata = await sheets.spreadsheets.get({ spreadsheetId });
    const firstSheetTitle = sheetMetadata.data.sheets?.[0]?.properties?.title;

    if (!firstSheetTitle) {
      throw new Error("Could not find any sheets in the spreadsheet.");
    }
    
    // Fetches all data from the first sheet.
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: firstSheetTitle,
    });

    const values = response.data.values;
    if (!values || values.length === 0) {
      return { data: [] }; // Return empty data instead of error for empty sheets
    }

    const jsonData = toJsonObject(values);
    return { data: jsonData };
  } catch (error: any) {
    console.error('Error fetching Google Sheet:', error);
    // Provide a more user-friendly error message for common issues
    if (error.code === 403) {
      return { error: 'Permission Denied (403). Make sure you have access to this Google Sheet and have granted the necessary permissions. You may need to log out and log back in.' };
    }
    if (error.code === 404) {
      return { error: `Not Found (404). The Google Sheet could not be found. Make sure the URL is correct.` };
    }
    if (error.code === 401) {
        return { error: `Authentication failed (401). Your session may have expired. Please log out and sign in again.` };
    }
    return { error: error.message || 'An unknown error occurred while fetching the sheet.' };
  }
}

export async function getSession(): Promise<SessionData> {
  const session = await getIronSession<SessionData>(cookies(), sessionOptions);
  // Return a plain object to avoid non-serializable data issues in Client Components
  return {
    isLoggedIn: !!session.isLoggedIn,
    name: session.name,
    email: session.email,
    picture: session.picture,
    // Do not return session.tokens here unless needed on the client,
    // as it might also contain non-serializable data.
  };
}

export async function logout() {
  const session = await getIronSession<SessionData>(cookies(), sessionOptions);
  session.destroy();
}
