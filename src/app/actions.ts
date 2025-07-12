'use server';

import * as XLSX from 'xlsx';

// Helper function to convert CSV to JSON
const csvToJson = (csv: string): Record<string, any>[] => {
  const lines = csv.split('\n');
  if (lines.length < 1) return [];

  const headers = lines[0].split(',').map(h => h.trim().replace(/"/g, ''));
  const jsonData = [];

  for (let i = 1; i < lines.length; i++) {
    const line = lines[i];
    if (!line.trim()) continue;

    // This is a simple regex to handle commas inside quoted strings.
    // It's not perfect but works for many standard CSVs.
    const values = line.match(/(".*?"|[^",]+)(?=\s*,|\s*$)/g) || [];
    
    const row: Record<string, any> = {};
    for (let j = 0; j < headers.length; j++) {
      const value = (values[j] || '').trim().replace(/"/g, '');
      row[headers[j]] = value;
    }
    jsonData.push(row);
  }

  return jsonData;
};

export async function fetchGoogleSheetData(sheetUrl: string): Promise<{ data?: Record<string, any>[]; error?: string }> {
  if (!sheetUrl) {
    return { error: 'Google Sheet URL is required.' };
  }

  try {
    const url = new URL(sheetUrl);
    const pathParts = url.pathname.split('/');
    const sheetId = pathParts[3];

    if (!sheetId) {
      throw new Error('Invalid Google Sheet URL. Could not find Sheet ID.');
    }
    
    // Default to the first sheet (gid=0) if not specified
    const gid = url.hash.startsWith('#gid=') ? url.hash.substring(5) : '0';

    const csvUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`;

    const response = await fetch(csvUrl, {
      headers: {
        'Accept': 'text/csv',
      },
      // Use no-cache to get the latest data
      cache: 'no-store',
    });

    if (!response.ok) {
      throw new Error(`Failed to fetch sheet. Status: ${response.status}. Make sure the sheet is public.`);
    }

    const csvData = await response.text();
    const jsonData = csvToJson(csvData);

    return { data: jsonData };
  } catch (error: any) {
    console.error('Error fetching Google Sheet:', error);
    return { error: error.message || 'An unknown error occurred while fetching the sheet.' };
  }
}
