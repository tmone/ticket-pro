/**
 * Client-side hook for Google Sheets integration via API routes
 */

import { useState, useCallback, useRef, useEffect } from 'react';

export interface GoogleSheetsState {
  isConnected: boolean;
  isLoading: boolean;
  error: string | null;
  spreadsheetTitle: string | null;
  sheetNames: string[];
  spreadsheetId: string | null;
  sheetName: string | null;
}

export interface AttendeeData {
  [key: string]: string | number | Date | null;
  __rowNum__: number;
  checkedInTime: Date | null;
}

export interface UseGoogleSheetsReturn {
  state: GoogleSheetsState;
  connectToSheets: (spreadsheetId: string, sheetName?: string) => Promise<void>;
  loadAttendeeData: (sheetName?: string, spreadsheetId?: string) => Promise<{ headers: string[]; rows: AttendeeData[] } | null>;
  saveCheckIn: (rowNumber: number, checkInTime: Date) => Promise<void>;
  disconnect: () => void;
}

export function useGoogleSheetsApi(): UseGoogleSheetsReturn {
  const [state, setState] = useState<GoogleSheetsState>({
    isConnected: false,
    isLoading: false,
    error: null,
    spreadsheetTitle: null,
    sheetNames: [],
    spreadsheetId: null,
    sheetName: null,
  });

  // Use ref to always have the latest state
  const stateRef = useRef(state);
  useEffect(() => {
    stateRef.current = state;
  }, [state]);

  const connectToSheets = useCallback(async (spreadsheetId: string, sheetName: string = 'Sheet1') => {
    setState(prev => ({ ...prev, isLoading: true, error: null }));

    try {
      const response = await fetch('/api/google-sheets/connect', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ spreadsheetId, sheetName }),
      });

      const data = await response.json();
      
      if (!response.ok) {
        throw new Error(data.error || 'Failed to connect to Google Sheets');
      }

      setState({
        isConnected: true,
        isLoading: false,
        error: null,
        spreadsheetTitle: data.spreadsheetTitle,
        sheetNames: data.sheetNames || [],
        spreadsheetId,
        sheetName,
      });

      console.log('✅ Connected to Google Sheets:', data.spreadsheetTitle);

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      setState(prev => ({
        ...prev,
        isConnected: false,
        isLoading: false,
        error: errorMessage,
      }));
      console.error('❌ Failed to connect to Google Sheets:', errorMessage);
      throw error;
    }
  }, []);

  const loadAttendeeData = useCallback(async (overrideSheetName?: string, overrideSpreadsheetId?: string): Promise<{ headers: string[]; rows: AttendeeData[] } | null> => {
    const spreadsheetId = overrideSpreadsheetId || state.spreadsheetId;
    
    if (!spreadsheetId) {
      throw new Error('Not connected to Google Sheets');
    }

    setState(prev => ({ ...prev, isLoading: true, error: null }));

    try {
      const sheetName = overrideSheetName || state.sheetName;
      console.log('Loading data with:', { spreadsheetId, sheetName });
      
      const response = await fetch('/api/google-sheets/read', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          spreadsheetId: spreadsheetId, 
          sheetName: sheetName 
        }),
      });

      const data = await response.json();
      
      if (!response.ok) {
        throw new Error(data.error || 'Failed to load data');
      }

      setState(prev => ({ ...prev, isLoading: false }));
      return data;
      
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Failed to load data';
      setState(prev => ({ ...prev, isLoading: false, error: errorMessage }));
      console.error('❌ Failed to load attendee data:', errorMessage);
      throw error;
    }
  }, [state.spreadsheetId, state.sheetName]);

  const saveCheckIn = useCallback(async (rowNumber: number, checkInTime: Date): Promise<void> => {
    const currentState = stateRef.current;
    console.log('saveCheckIn called with:', { rowNumber, checkInTime, currentState });
    
    if (!currentState.spreadsheetId) {
      throw new Error('Not connected to Google Sheets');
    }

    try {
      console.log('Sending check-in request:', {
        spreadsheetId: currentState.spreadsheetId,
        sheetName: currentState.sheetName,
        rowNumber,
        checkInTime: checkInTime.toISOString()
      });
      
      const response = await fetch('/api/google-sheets/checkin', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          spreadsheetId: currentState.spreadsheetId,
          sheetName: currentState.sheetName,
          rowNumber,
          checkInTime: checkInTime.toISOString(),
        }),
      });

      const data = await response.json();
      
      if (!response.ok) {
        throw new Error(data.error || 'Failed to save check-in');
      }

      console.log('✅ Check-in saved to Google Sheets');
      
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Failed to save check-in';
      console.error('❌ Failed to save check-in:', errorMessage);
      throw error;
    }
  }, []);

  const disconnect = useCallback(() => {
    setState({
      isConnected: false,
      isLoading: false,
      error: null,
      spreadsheetTitle: null,
      sheetNames: [],
      spreadsheetId: null,
      sheetName: null,
    });
    console.log('📤 Disconnected from Google Sheets');
  }, []);

  return {
    state,
    connectToSheets,
    loadAttendeeData,
    saveCheckIn,
    disconnect,
  };
}