/**
 * Google Sheets connection component
 */

import * as React from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Loader2, ExternalLink, CheckCircle, AlertCircle, Sheet, ChevronDown, ChevronUp } from "lucide-react";
import { UseGoogleSheetsReturn } from "@/hooks/use-google-sheets-api";
import { useToast } from "@/hooks/use-toast";
import { Collapsible, CollapsibleContent, CollapsibleTrigger } from "@/components/ui/collapsible";

interface GoogleSheetsConnectorProps {
  onDataLoaded: (data: { headers: string[]; rows: any[] }) => void;
  onConnectionChange: (isConnected: boolean) => void;
  googleSheetsApi: UseGoogleSheetsReturn;
}

export function GoogleSheetsConnector({ onDataLoaded, onConnectionChange, googleSheetsApi }: GoogleSheetsConnectorProps) {
  const { toast } = useToast();
  
  // Initialize from localStorage or env
  const getInitialUrl = () => {
    if (typeof window !== 'undefined') {
      const savedUrl = localStorage.getItem('googleSheetsUrl');
      if (savedUrl) return savedUrl;
    }
    return process.env.NEXT_PUBLIC_DEFAULT_GOOGLE_SHEETS_URL || '';
  };
  
  const getInitialSheetName = () => {
    if (typeof window !== 'undefined') {
      const savedSheetName = localStorage.getItem('googleSheetsSheetName');
      if (savedSheetName) return savedSheetName;
    }
    return 'Sheet1';
  };
  
  const [spreadsheetUrl, setSpreadsheetUrl] = React.useState(getInitialUrl());
  const [sheetName, setSheetName] = React.useState(getInitialSheetName());
  const [isConnecting, setIsConnecting] = React.useState(false);
  const [hasAutoConnected, setHasAutoConnected] = React.useState(false);
  const [isCollapsed, setIsCollapsed] = React.useState(false);
  
  const { state, connectToSheets, loadAttendeeData, disconnect } = googleSheetsApi;

  // Extract spreadsheet ID from Google Sheets URL
  const extractSpreadsheetId = (url: string): string | null => {
    try {
      // Google Sheets URL format: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit...
      const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      if (match && match[1]) {
        // Google Sheets IDs are typically 44 characters long
        if (match[1].length < 40) {
          console.warn(`Spreadsheet ID seems too short: ${match[1]} (${match[1].length} characters)`);
        }
        return match[1];
      }
      return null;
    } catch (error) {
      return null;
    }
  };

  const handleConnect = async () => {
    if (!spreadsheetUrl.trim()) {
      return;
    }

    const spreadsheetId = extractSpreadsheetId(spreadsheetUrl);
    if (!spreadsheetId) {
      alert('Invalid Google Sheets URL. Please use a valid Google Sheets link.');
      return;
    }

    setIsConnecting(true);
    try {
      const targetSheetName = sheetName || 'Sheet1';
      await connectToSheets(spreadsheetId, targetSheetName);
      
      // Save to localStorage after successful connection
      localStorage.setItem('googleSheetsUrl', spreadsheetUrl);
      localStorage.setItem('googleSheetsSheetName', targetSheetName);
      
      // Auto-load data after successful connection with the correct sheet
      const data = await loadAttendeeData(targetSheetName, spreadsheetId);
      if (data) {
        onDataLoaded(data);
        onConnectionChange(true);
      }
    } catch (error) {
      console.error('Connection failed:', error);
    } finally {
      setIsConnecting(false);
    }
  };

  const handleDisconnect = () => {
    disconnect();
    onConnectionChange(false);
    setSpreadsheetUrl('');
    setSheetName('Sheet1');
    
    // Clear localStorage
    localStorage.removeItem('googleSheetsUrl');
    localStorage.removeItem('googleSheetsSheetName');
  };

  const handleRefreshData = async () => {
    if (!state.isConnected) return;
    
    try {
      const data = await loadAttendeeData();
      if (data) {
        onDataLoaded(data);
      }
    } catch (error) {
      console.error('Failed to refresh data:', error);
    }
  };

  // Auto-connect on mount if URL is provided
  React.useEffect(() => {
    // Small delay to ensure component is fully mounted
    const timer = setTimeout(() => {
      if (spreadsheetUrl && !hasAutoConnected && !state.isConnected) {
        console.log('Auto-connecting with URL:', spreadsheetUrl, 'Sheet:', sheetName);
        setHasAutoConnected(true);
        handleConnect();
      }
    }, 100);
    
    return () => clearTimeout(timer);
  }, [spreadsheetUrl]); // Depend on spreadsheetUrl to re-run when it's set
  
  // Ensure connection state is synced and auto-collapse when connected
  React.useEffect(() => {
    if (state.isConnected) {
      onConnectionChange(true);
      setIsCollapsed(true); // Auto-collapse when connected
    }
  }, [state.isConnected, onConnectionChange]);

  // Auto-select first sheet if current sheet doesn't exist
  React.useEffect(() => {
    if (state.sheetNames.length > 0 && sheetName === 'Sheet1' && !state.sheetNames.includes('Sheet1')) {
      const firstSheet = state.sheetNames[0];
      setSheetName(firstSheet);
      
      // Reload data with correct sheet
      if (state.isConnected && state.spreadsheetId) {
        connectToSheets(state.spreadsheetId, firstSheet).then(() => {
          loadAttendeeData(firstSheet, state.spreadsheetId).then(data => {
            if (data) {
              onDataLoaded(data);
              onConnectionChange(true);
              // Save the correct sheet name
              localStorage.setItem('googleSheetsSheetName', firstSheet);
            }
          });
        });
      }
    }
  }, [state.sheetNames]);

  return (
    <Card>
      <Collapsible open={!isCollapsed} onOpenChange={(open) => setIsCollapsed(!open)}>
        <CardHeader>
          <div className="flex items-center justify-between">
            <div className="space-y-1">
              <CardTitle className="flex items-center gap-2">
                <Sheet className="h-5 w-5 text-green-600" />
                Google Sheets Integration
                {state.isConnected && (
                  <Badge variant="outline" className="ml-2 text-green-600 border-green-600">
                    Connected
                  </Badge>
                )}
              </CardTitle>
              {!state.isConnected ? (
                <CardDescription>
                  Connect directly to your Google Sheets to load attendee data and save check-ins in real-time.
                </CardDescription>
              ) : (
                isCollapsed && (
                  <CardDescription className="flex items-center gap-2">
                    <span>{state.spreadsheetTitle}</span>
                    {sheetName && (
                      <>
                        <span className="text-muted-foreground">•</span>
                        <span>Sheet: {sheetName}</span>
                      </>
                    )}
                  </CardDescription>
                )
              )}
            </div>
            {state.isConnected && (
              <CollapsibleTrigger asChild>
                <Button variant="ghost" size="sm" className="ml-auto">
                  {isCollapsed ? <ChevronDown className="h-4 w-4" /> : <ChevronUp className="h-4 w-4" />}
                </Button>
              </CollapsibleTrigger>
            )}
          </div>
        </CardHeader>
        <CollapsibleContent>
          <CardContent className="space-y-4">
        {!state.isConnected ? (
          <>
            <div className="space-y-2">
              <Label htmlFor="sheets-url">Google Sheets URL</Label>
              <Input
                id="sheets-url"
                placeholder="https://docs.google.com/spreadsheets/d/your-sheet-id/edit"
                value={spreadsheetUrl}
                onChange={(e) => setSpreadsheetUrl(e.target.value)}
                disabled={isConnecting}
              />
              <p className="text-xs text-muted-foreground">
                Paste the full URL of your Google Sheets document
              </p>
            </div>

            <div className="space-y-2">
              <Label htmlFor="sheet-name">Sheet Name (Optional)</Label>
              <Input
                id="sheet-name"
                placeholder="Sheet1"
                value={sheetName}
                onChange={(e) => setSheetName(e.target.value)}
                disabled={isConnecting}
              />
              <p className="text-xs text-muted-foreground">
                Leave empty to use the first sheet
              </p>
            </div>

            <Button 
              onClick={handleConnect} 
              disabled={!spreadsheetUrl.trim() || isConnecting || state.isLoading}
              className="w-full"
            >
              {isConnecting || state.isLoading ? (
                <>
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  Connecting...
                </>
              ) : (
                <>
                  <ExternalLink className="mr-2 h-4 w-4" />
                  Connect to Google Sheets
                </>
              )}
            </Button>

            {state.error && (
              <Alert variant="destructive">
                <AlertCircle className="h-4 w-4" />
                <AlertTitle>Connection Failed</AlertTitle>
                <AlertDescription className="space-y-2">
                  {state.error.includes('ENABLE_API_REQUIRED:') ? (
                    <div className="space-y-3">
                      <p className="font-medium">Google Sheets API is not enabled in your project</p>
                      <div className="border rounded p-3 bg-yellow-50">
                        <p className="text-sm font-medium mb-2">Quick Fix:</p>
                        <ol className="text-sm list-decimal list-inside space-y-1 mb-3">
                          <li>Click the button below to enable the API</li>
                          <li>Click "ENABLE" on the Google page</li>
                          <li>Wait 1-2 minutes for it to activate</li>
                          <li>Try connecting again</li>
                        </ol>
                        <Button
                          size="sm"
                          onClick={() => window.open(state.error.split('ENABLE_API_REQUIRED:')[1], '_blank')}
                          className="w-full"
                        >
                          <ExternalLink className="mr-2 h-4 w-4" />
                          Enable Google Sheets API
                        </Button>
                      </div>
                    </div>
                  ) : state.error.includes('INVALID_DOCUMENT:') ? (
                    <div className="space-y-3">
                      <p className="font-medium text-red-600">Invalid Document Type</p>
                      <div className="border rounded p-3 bg-red-50">
                        <p className="text-sm mb-2">{state.error.split('INVALID_DOCUMENT:')[1]}</p>
                        <p className="text-sm font-medium mb-2">Valid Google Sheets URL format:</p>
                        <code className="text-xs bg-gray-100 p-2 rounded block">
                          https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
                        </code>
                        <div className="mt-3 text-xs text-gray-600">
                          <p className="font-medium mb-1">Common issues:</p>
                          <ul className="list-disc list-inside space-y-1">
                            <li>Using Google Docs URL instead of Google Sheets</li>
                            <li>Using Excel Online or other spreadsheet URLs</li>
                            <li>Incorrect URL format</li>
                          </ul>
                        </div>
                      </div>
                    </div>
                  ) : state.error.includes('SHEET_NOT_FOUND:') ? (
                    <div className="space-y-3">
                      <p className="font-medium text-orange-600">Spreadsheet Not Found</p>
                      <div className="border rounded p-3 bg-orange-50">
                        <p className="text-sm mb-2">{state.error.split('SHEET_NOT_FOUND:')[1]}</p>
                        <p className="text-sm font-medium mb-2">Please check:</p>
                        <ul className="text-sm list-disc list-inside space-y-1">
                          <li>The URL is correct and complete</li>
                          <li>You have access to the spreadsheet</li>
                          <li>The spreadsheet hasn't been deleted</li>
                          <li>If using Service Account, the sheet is shared with the service account email</li>
                        </ul>
                      </div>
                    </div>
                  ) : (
                    <>
                      <p>{state.error}</p>
                      {(state.error.includes('403') || state.error.includes('access_denied') || state.error.includes('OAuth configuration error')) ? (
                    <div className="space-y-2">
                      <p className="text-xs font-medium">Your app is in testing mode. Choose one option:</p>
                      
                      <div className="space-y-3">
                        <div className="border rounded p-2">
                          <p className="font-medium text-xs mb-1">Option 1: Add Test User (Quick Fix)</p>
                          <ol className="text-xs list-decimal list-inside space-y-1 mb-2">
                            <li>Go to <a href="https://console.cloud.google.com/apis/credentials/consent" target="_blank" className="underline text-blue-600">OAuth consent screen</a></li>
                            <li>Click "ADD USERS" under Test users</li>
                            <li>Add your Google email</li>
                            <li>Save and try again</li>
                          </ol>
                          <Button
                            size="sm"
                            variant="outline"
                            onClick={() => window.open('https://console.cloud.google.com/apis/credentials/consent', '_blank')}
                            className="w-full text-xs"
                          >
                            <ExternalLink className="mr-2 h-3 w-3" />
                            Open OAuth Consent Screen
                          </Button>
                        </div>
                        
                        <div className="border rounded p-2">
                          <p className="font-medium text-xs mb-1">Option 2: Use Service Account (Recommended)</p>
                          <p className="text-xs mb-2">No OAuth needed, works immediately</p>
                          <Button
                            size="sm"
                            variant="outline"
                            onClick={() => window.open('/service-account-guide.html', '_blank')}
                            className="w-full text-xs"
                          >
                            View Service Account Setup Guide
                          </Button>
                        </div>
                      </div>
                    </div>
                  ) : null}
                    </>
                  )}
                </AlertDescription>
              </Alert>
            )}

            <Alert>
              <AlertCircle className="h-4 w-4" />
              <AlertTitle>Setup Required</AlertTitle>
              <AlertDescription className="space-y-2">
                <p>To use Google Sheets integration:</p>
                <div className="space-y-3">
                  <div>
                    <p className="font-medium mb-1">Option 1: Quick OAuth Setup</p>
                    <Button
                      size="sm"
                      variant="outline"
                      onClick={() => window.location.href = '/api/auth/google'}
                      className="w-full"
                    >
                      <ExternalLink className="mr-2 h-4 w-4" />
                      Authorize with Google
                    </Button>
                    <p className="text-xs text-muted-foreground mt-1">
                      This will get you a refresh token automatically
                    </p>
                  </div>
                  
                  <div className="text-xs text-muted-foreground">
                    <p className="font-medium mb-1">Option 2: Manual Setup</p>
                    <ol className="list-decimal list-inside space-y-0.5">
                      <li>Add credentials to <code className="bg-muted px-1 rounded">.env.local</code></li>
                      <li>Use Service Account for production</li>
                    </ol>
                  </div>
                </div>
              </AlertDescription>
            </Alert>
          </>
        ) : (
          <>
            <div className="space-y-3">
              <div className="flex items-center gap-2">
                <CheckCircle className="h-4 w-4 text-green-600" />
                <span className="text-sm font-medium">Connected Successfully</span>
              </div>
              
              {state.spreadsheetTitle && (
                <div className="text-sm">
                  <strong>Document:</strong> {state.spreadsheetTitle}
                </div>
              )}

              {state.sheetNames.length > 0 && (
                <div className="space-y-2">
                  <Label htmlFor="sheet-select">Select Sheet</Label>
                  <Select 
                    value={sheetName} 
                    onValueChange={async (value) => {
                      console.log('Switching to sheet:', value);
                      setSheetName(value);
                      setIsConnecting(true);
                      
                      // Reconnect with new sheet
                      if (state.spreadsheetId) {
                        try {
                          // First connect to the new sheet
                          console.log('Connecting to sheet:', value);
                          await connectToSheets(state.spreadsheetId, value);
                          
                          // Add small delay to ensure connection is established
                          await new Promise(resolve => setTimeout(resolve, 500));
                          
                          // Then load the data with explicit sheet name
                          console.log('Loading data from sheet:', value);
                          const data = await loadAttendeeData(value, state.spreadsheetId);
                          console.log('Data loaded:', data);
                          
                          if (data) {
                            onDataLoaded(data);
                            // Save new sheet name to localStorage
                            localStorage.setItem('googleSheetsSheetName', value);
                            toast({
                              title: "Sheet Changed",
                              description: `Loaded ${data.rows.length} rows from sheet: ${value}`,
                            });
                          } else {
                            throw new Error('No data returned from loadAttendeeData');
                          }
                        } catch (error) {
                          console.error('Failed to switch sheet:', error);
                          toast({
                            title: "Error",
                            description: `Failed to load data from sheet "${value}"`,
                            variant: "destructive",
                          });
                        } finally {
                          setIsConnecting(false);
                        }
                      }
                    }}
                    disabled={isConnecting || state.isLoading}
                  >
                    <SelectTrigger id="sheet-select">
                      {isConnecting ? (
                        <div className="flex items-center gap-2">
                          <Loader2 className="h-4 w-4 animate-spin" />
                          <span>Loading...</span>
                        </div>
                      ) : (
                        <SelectValue placeholder="Select a sheet" />
                      )}
                    </SelectTrigger>
                    <SelectContent>
                      {state.sheetNames.map((name) => (
                        <SelectItem key={name} value={name}>
                          {name}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              )}
            </div>

            <div className="flex gap-2">
              <Button 
                onClick={handleRefreshData} 
                disabled={state.isLoading}
                variant="outline"
                size="sm"
              >
                {state.isLoading ? (
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                ) : (
                  "Refresh Data"
                )}
              </Button>
              
              <Button 
                onClick={handleDisconnect}
                variant="outline"
                size="sm"
              >
                Disconnect
              </Button>
            </div>

            {state.error && (
              <Alert variant="destructive">
                <AlertCircle className="h-4 w-4" />
                <AlertTitle>Error</AlertTitle>
                <AlertDescription>
                  {state.error}
                </AlertDescription>
              </Alert>
            )}
          </>
        )}
          </CardContent>
        </CollapsibleContent>
      </Collapsible>
    </Card>
  );
}