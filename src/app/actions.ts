'use server';

import * as XLSX from 'xlsx';

// Helper function to convert CSV to JSON
const csvToJson = (csv: string): Record<string, any>[] => {
  const lines = csv.split('\n');
  if (lines.length < 1) return [];

  // Remove trailing carriage return from headers if present
  const headers = lines[0].split(',').map(h => h.trim().replace(/"/g, '').replace(/\r$/, ''));
  const jsonData = [];

  for (let i = 1; i < lines.length; i++) {
    const line = lines[i];
    if (!line.trim()) continue;

    // This regex handles commas inside quoted strings.
    const values = line.match(/(".*?"|[^",]+)(?=\s*,|\s*$)/g) || [];
    
    const row: Record<string, any> = {};
    for (let j = 0; j < headers.length; j++) {
      if (headers[j]) { // Ensure header exists
          const value = (values[j] || '').trim().replace(/"/g, '').replace(/\r$/, '');
          row[headers[j]] = value;
      }
    }
    if (Object.keys(row).length > 0) {
      jsonData.push(row);
    }
  }

  return jsonData;
};

export async function fetchGoogleSheetData(sheetUrl: string): Promise<{ data?: Record<string, any>[]; error?: string }> {
  if (!sheetUrl) {
    return { error: 'Google Sheet URL is required.' };
  }

  try {
    const url = new URL(sheetUrl);

    // More robustly find the sheet ID, which is usually between /d/ and the next /
    const match = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/.exec(url.pathname);
    if (!match || !match[1]) {
      throw new Error('Invalid Google Sheet URL. Could not find a valid Sheet ID.');
    }
    const sheetId = match[1];

    // Default to the first sheet (gid=0) if not specified in the hash
    const gid = url.hash.startsWith('#gid=') ? url.hash.substring(5) : '0';

    if (!/^\d+$/.test(gid)) {
        throw new Error('Invalid GID found in URL. GID must be a number.');
    }

    const csvUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`;

    const response = await fetch(csvUrl, {
      headers: {
        'Accept': 'text/csv',
      },
      // Use no-cache to get the latest data
      cache: 'no-store',
    });

    if (!response.ok) {
      // Provide more specific feedback for common errors
      if (response.status === 400) {
          return { error: `Bad Request (400). Please check if the Sheet ID '${sheetId}' and GID '${gid}' are correct and the sheet exists.` };
      }
       if (response.status === 404) {
          return { error: `Not Found (404). The Google Sheet could not be found. Make sure the URL is correct and the sheet is public.` };
      }
      throw new Error(`Failed to fetch sheet. Status: ${response.status}. Make sure the sheet is public ("Anyone with the link can view").`);
    }

    const csvData = await response.text();
    if (!csvData) {
        return { error: 'The Google Sheet appears to be empty or could not be read.' };
    }

    const jsonData = csvToJson(csvData);

    return { data: jsonData };
  } catch (error: any) {
    console.error('Error fetching Google Sheet:', error);
    return { error: error.message || 'An unknown error occurred while fetching the sheet.' };
  }
}
