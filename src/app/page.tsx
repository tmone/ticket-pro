
"use client";

import * as React from "react";
import { z } from "zod";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { format } from "date-fns";
import * as XLSX from "xlsx";
import { useRouter } from "next/navigation";
import type { WorkBook, WorkSheet, CellObject, ExcelDataType } from "xlsx";
import jsqr from "jsqr";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import { useToast } from "@/hooks/use-toast";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from "@/components/ui/alert-dialog";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Alert, AlertTitle, AlertDescription as AlertDescriptionUI } from "@/components/ui/alert";
import { Form, FormControl, FormField, FormItem, FormLabel, FormMessage } from "@/components/ui/form";
import { Checkbox } from "@/components/ui/checkbox";
import { Label } from "@/components/ui/label";
import {
  TicketCheck,
  Upload,
  QrCode,
  Download,
  UserCheck,
  AlertTriangle,
  Camera,
  FileSpreadsheet,
  CheckCircle,
  LogOut
} from "lucide-react";
import { cn } from "@/lib/utils";
import { UniversalExcelHandler } from "@/lib/excel-handler";
import { GoogleSheetsConnector } from "@/components/google-sheets-connector";
import { useGoogleSheetsApi } from "@/hooks/use-google-sheets-api";

// Note: Excel utility functions moved to src/lib/excel-utils.ts for better organization

const checkInSchema = z.object({
  uniqueCode: z.string().min(1, { message: "Code is required." }),
});

type DialogState = 'success' | 'duplicate' | 'not_found';

export default function DashboardPage() {
  const { toast } = useToast();
  const router = useRouter();
  const videoRef = React.useRef<HTMLVideoElement>(null);
  const canvasRef = React.useRef<HTMLCanvasElement>(null);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const inputRef = React.useRef<HTMLInputElement | null>(null);
  const animationFrameIdRef = React.useRef<number>();
  const scanSourceRef = React.useRef<'camera' | 'form' | null>(null);
  const rowRefs = React.useRef<(HTMLTableRowElement | null)[]>([]);

  const [headers, setHeaders] = React.useState<string[]>([]);
  const [rows, setRows] = React.useState<Record<string, any>[]>([]);
  const [scannedRow, setScannedRow] = React.useState<Record<string, any> | null | undefined>(null);
  const [isAlertOpen, setIsAlertOpen] = React.useState(false);
  const [dialogState, setDialogState] = React.useState<DialogState>('not_found');
  
  const [isScanning, setIsScanning] = React.useState(false);
  const [scanError, setScanError] = React.useState<string | null>(null);
  const [isContinuous, setIsContinuous] = React.useState(false);

  const [workbook, setWorkbook] = React.useState<WorkBook | null>(null);
  const [originalFileData, setOriginalFileData] = React.useState<Uint8Array | null>(null);
  const [excelHandler, setExcelHandler] = React.useState<UniversalExcelHandler | null>(null);
  const [sheetNames, setSheetNames] = React.useState<string[]>([]);
  const [activeSheetName, setActiveSheetName] = React.useState<string | null>(null);
  const [isSheetSelectorOpen, setIsSheetSelectorOpen] = React.useState(false);
  const [highlightedRowIndex, setHighlightedRowIndex] = React.useState<number | null>(null);
  const [isClient, setIsClient] = React.useState(false);
  const [dataSource, setDataSource] = React.useState<'excel' | 'google-sheets'>('excel');
  const [isGoogleSheetsConnected, setIsGoogleSheetsConnected] = React.useState(false);
  
  // Google Sheets integration - single instance
  const googleSheetsApi = useGoogleSheetsApi();
  
  React.useEffect(() => {
    setIsClient(true);
  }, []);
  
  // Safe date formatter to prevent hydration mismatch
  const formatDateSafe = (date: Date | string | null | undefined) => {
    if (!isClient || !date) return 'N/A';
    try {
      const dateObj = typeof date === 'string' ? new Date(date) : date;
      return format(dateObj, 'PPpp');
    } catch (error) {
      console.warn('Error formatting date:', error);
      return 'N/A';
    }
  };
  
  const checkInForm = useForm<z.infer<typeof checkInSchema>>({
    resolver: zodResolver(checkInSchema),
    defaultValues: {
      uniqueCode: "",
    },
  });

  const stopScan = React.useCallback(() => {
    if (animationFrameIdRef.current) {
        cancelAnimationFrame(animationFrameIdRef.current);
        animationFrameIdRef.current = undefined;
    }
    setIsScanning(false);
    if (videoRef.current && videoRef.current.srcObject) {
      const stream = videoRef.current.srcObject as MediaStream;
      stream.getTracks().forEach((track) => track.stop());
      videoRef.current.srcObject = null;
    }
  }, []);

  const handleCheckIn = React.useCallback((data: z.infer<typeof checkInSchema>) => {
    const { uniqueCode } = data;
    if (!uniqueCode) return;
    
    if (scanSourceRef.current === 'camera' && isScanning) {
      stopScan();
    }
    
    let codeToSearch = uniqueCode.trim();
    try {
        const url = new URL(codeToSearch);
        const params = url.searchParams;
        const codeParam = params.get("code") || params.get("id");
        if (codeParam) {
            codeToSearch = codeParam.trim();
        } else {
            const firstParam = params.values().next().value;
            if (firstParam) codeToSearch = firstParam.trim();
        }
    } catch(e) { /* Not a valid URL, use codeToSearch as is */ }

    let foundRowIndex = -1;
    
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        for (const header of headers) {
            const cellValue = row[header];
            if (cellValue === undefined || cellValue === null) continue;

            let cellCode = String(cellValue).trim();
            try {
                const url = new URL(cellCode);
                const params = url.searchParams;
                const codeParam = params.get("code") || params.get("id");
                if (codeParam) {
                    cellCode = codeParam.trim();
                } else {
                   const firstParam = params.values().next().value;
                   if (firstParam) cellCode = firstParam.trim();
                }
            } catch (e) { /* not a url */ }

            if (codeToSearch.toLowerCase() === cellCode.toLowerCase()) {
                foundRowIndex = i;
                break;
            }
        }
        if (foundRowIndex !== -1) break;
    }
    
    setHighlightedRowIndex(foundRowIndex !== -1 ? foundRowIndex : null);
    if(foundRowIndex !== -1) {
        const rowElement = rowRefs.current[foundRowIndex];
        if (rowElement) {
            rowElement.scrollIntoView({
                behavior: 'smooth',
                block: 'center'
            });
        }
    }

    if (foundRowIndex !== -1) {
      const foundRowData = rows[foundRowIndex];
      if (foundRowData.checkedInTime) {
        setScannedRow(foundRowData);
        setDialogState('duplicate');
        setIsAlertOpen(true);
      } else {
        const checkInTime = new Date();
        const updatedRow = { ...foundRowData, checkedInTime: checkInTime, __rowNum__: foundRowData.__rowNum__ };
        const updatedRows = [...rows];
        updatedRows[foundRowIndex] = updatedRow;
        setRows(updatedRows);
        setScannedRow(updatedRow);
        setDialogState('success');
        setIsAlertOpen(true);
        
        // Real-time save to Google Sheets if connected
        console.log('Google Sheets save check:', {
          dataSource,
          isGoogleSheetsConnected,
          rowNum: foundRowData.__rowNum__,
          shouldSave: dataSource === 'google-sheets' && isGoogleSheetsConnected && foundRowData.__rowNum__
        });
        
        if (dataSource === 'google-sheets' && isGoogleSheetsConnected && foundRowData.__rowNum__) {
          console.log('Saving to Google Sheets, row:', foundRowData.__rowNum__);
          googleSheetsApi.saveCheckIn(foundRowData.__rowNum__, checkInTime).catch(error => {
            console.error('Failed to save check-in to Google Sheets:', error);
            // Note: We don't show user error here to avoid disrupting the check-in flow
          });
        }
      }
    } else {
      setScannedRow(undefined);
      setDialogState('not_found');
      setIsAlertOpen(true);
    }
  }, [rows, headers, isScanning, stopScan]);
  
  const tick = React.useCallback(() => {
    if (videoRef.current && videoRef.current.readyState === videoRef.current.HAVE_ENOUGH_DATA && canvasRef.current) {
        const video = videoRef.current;
        const canvas = canvasRef.current;
        const ctx = canvas.getContext("2d", { willReadFrequently: true });

        if (ctx) {
            canvas.height = video.videoHeight;
            canvas.width = video.videoWidth;
            ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
            const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
            const code = jsqr(imageData.data, imageData.width, imageData.height);

            if (code && code.data) {
              scanSourceRef.current = 'camera';
              checkInForm.setValue('uniqueCode', code.data);
              handleCheckIn({ uniqueCode: code.data }); 
              return;
            }
        }
    }
    if (animationFrameIdRef.current) {
      animationFrameIdRef.current = requestAnimationFrame(tick);
    }
  }, [checkInForm, handleCheckIn]); 

  const startScan = React.useCallback(async () => {
    setScanError(null);
    if (isScanning || animationFrameIdRef.current) return;
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } });
      if (videoRef.current) {
        videoRef.current.srcObject = stream;
        await new Promise(resolve => videoRef.current!.onloadedmetadata = resolve);
        await videoRef.current.play();
        setIsScanning(true);
        animationFrameIdRef.current = requestAnimationFrame(tick);
      }
    } catch (err) {
      console.error("Camera access error:", err);
      setScanError("Camera access denied. Please enable it in your browser settings.");
      setIsScanning(false);
    }
  }, [isScanning, tick]);

  const handleLogout = async () => {
    try {
      await fetch('/api/auth/logout', { method: 'POST' });
      router.push('/login');
      router.refresh();
    } catch (error) {
      console.error('Logout error:', error);
    }
  };

  const handleAlertClose = React.useCallback(() => {
    setIsAlertOpen(false);
    checkInForm.reset();
    
    if (isContinuous && scanSourceRef.current === 'camera') {
      setTimeout(() => startScan(), 100); 
    } else if (isContinuous && scanSourceRef.current === 'form') {
      inputRef.current?.focus();
    }
  }, [isContinuous, checkInForm, startScan]);


  React.useEffect(() => {
    return () => {
      stopScan();
    };
  }, [stopScan]);

  React.useEffect(() => {
      if (isContinuous && dialogState === 'success' && isAlertOpen) {
          const timer = setTimeout(() => {
              handleAlertClose();
          }, 1500); 
          return () => clearTimeout(timer);
      }
  }, [isAlertOpen, dialogState, isContinuous, handleAlertClose]);

  const processSheetData = (wb: WorkBook, sheetName: string) => {
    try {
      rowRefs.current = [];
      setHighlightedRowIndex(null);

      if (!wb || !wb.Sheets) {
          toast({
              variant: "destructive",
              title: "Workbook Error",
              description: "The workbook is invalid or corrupted."
          });
          return;
      }

      const worksheet = wb.Sheets[sheetName];
      if (!worksheet) {
          toast({
              variant: "destructive",
              title: "Sheet not found",
              description: `Sheet with name "${sheetName}" could not be found in the file.`
          });
          return;
      }

      const jsonData = XLSX.utils.sheet_to_json<Record<string, any>>(worksheet, {
        defval: ''
      });

      if (jsonData.length < 1) {
        toast({
            variant: "destructive",
            title: "No Data",
            description: "The selected sheet is empty or could not be read."
        });
        setHeaders([]);
        setRows([]);
        return;
      }
    
    const headerSet = new Set<string>();
    jsonData.forEach(row => {
        if (row && typeof row === 'object' && !Array.isArray(row)) {
            try {
                Object.keys(row).forEach(key => {
                    if (key && typeof key === 'string') {
                        headerSet.add(key);
                    }
                });
            } catch (error) {
                console.warn('Error processing row headers:', error);
            }
        }
    });
    const extractedHeaders = Array.from(headerSet);

    const processedRows = jsonData.map((row, index) => {
        try {
            if (row && typeof row === 'object' && !Array.isArray(row)) {
                return {
                    ...row,
                    __rowNum__: index + 2, // Assuming header is row 1, data starts at row 2
                    checkedInTime: null,
                };
            } else {
                console.warn(`Row ${index + 2} is not a valid object:`, row);
                return {
                    __rowNum__: index + 2,
                    checkedInTime: null,
                };
            }
        } catch (error) {
            console.warn(`Error processing row ${index + 2}:`, error);
            return {
                __rowNum__: index + 2,
                checkedInTime: null,
            };
        }
    });

    setHeaders(extractedHeaders);
    setRows(processedRows);
    setActiveSheetName(sheetName);
    setScannedRow(null);

    toast({
      title: "Success!",
      description: `Successfully imported ${processedRows.length} rows and ${extractedHeaders.length} columns from sheet: ${sheetName}.`,
    });
    } catch (error) {
      console.error('Error processing sheet data:', error);
      toast({
          variant: "destructive",
          title: "Processing Error",
          description: "There was an error processing the sheet data. Please check the file format."
      });
      setHeaders([]);
      setRows([]);
    }
  };

  const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;


    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        
        // Save raw file data for pristine re-reads during export
        setOriginalFileData(data);

        // Create universal handler and process file
        const handler = new UniversalExcelHandler();
        const result = await handler.readFile(data);
        
        console.log(`Using library: ${result.library}`);
        console.log(`Advanced formatting detected: ${result.hasFormatting}`);
        
        setExcelHandler(handler);
        setSheetNames(result.sheets);
        setActiveSheetName(null);
        setHeaders([]);
        setRows([]);
        setScannedRow(null);
        setHighlightedRowIndex(null);
        rowRefs.current = [];


        if (result.sheets.length === 0) {
            toast({
                variant: "destructive",
                title: "No Sheets Found",
                description: "The uploaded Excel file does not contain any sheets.",
            });
            return;
        }

        if (result.sheets.length === 1) {
            const sheetData = handler.processSheetData(result.sheets[0]);
            setHeaders(sheetData.headers);
            setRows(sheetData.rows);
            setActiveSheetName(result.sheets[0]);
            
            toast({
              title: "Success!",
              description: `File loaded with ${result.library}. Advanced formatting: ${result.hasFormatting ? 'Detected' : 'Basic'}`,
            });
        } else {
            setIsSheetSelectorOpen(true);
        }
        
      } catch (error) {
        console.error("Error processing Excel file:", error);
        toast({
          variant: "destructive",
          title: "File Error",
          description: "Could not process the Excel file. Please ensure it's a valid format.",
        });
      }
    };
    reader.onerror = () => {
        toast({
            variant: "destructive",
            title: "File Read Error",
            description: "There was an error reading the file."
        });
    }
    reader.readAsArrayBuffer(file);
    event.target.value = '';
  };
  
  const handleSheetSelect = (sheetName: string) => {
    if (sheetName && excelHandler) {
        const sheetData = excelHandler.processSheetData(sheetName);
        setHeaders(sheetData.headers);
        setRows(sheetData.rows);
        setActiveSheetName(sheetName);
        
        toast({
          title: "Success!",
          description: `Successfully imported ${sheetData.rows.length} rows from sheet: ${sheetName}.`,
        });
    }
    setIsSheetSelectorOpen(false);
  };

  const handleScanButtonClick = () => {
    if (!isScanning) {
      startScan();
    } else {
      stopScan();
    }
  };

  const handleGoogleSheetsDataLoaded = React.useCallback((data: { headers: string[]; rows: any[] }) => {
    setHeaders(data.headers);
    setRows(data.rows);
    setDataSource('google-sheets');
    setActiveSheetName('Google Sheets');
    setScannedRow(null);
    setHighlightedRowIndex(null);
    rowRefs.current = [];
    
    toast({
      title: "Google Sheets Connected!",
      description: `Successfully loaded ${data.rows.length} attendees from Google Sheets.`,
    });
  }, [toast]);

  const handleGoogleSheetsConnectionChange = React.useCallback((isConnected: boolean) => {
    setIsGoogleSheetsConnected(isConnected);
    if (!isConnected) {
      // Reset to Excel mode when disconnected
      setDataSource('excel');
      setHeaders([]);
      setRows([]);
      setActiveSheetName(null);
    }
  }, []);

  const handleExport = async () => {
    // Disable export for Google Sheets (data is saved real-time)
    if (dataSource === 'google-sheets') {
      toast({
        title: "Export Not Needed",
        description: "Check-in data is automatically saved to Google Sheets in real-time.",
      });
      return;
    }
    
    if (!excelHandler || !activeSheetName) {
      toast({
        variant: "destructive",
        title: "No Data Found",
        description: "Please upload an Excel file first."
      });
      return;
    }

    try {
      console.log('Starting export with enhanced formatting preservation...');
      
      // Export with automatic formatting preservation
      const exportBuffer = await excelHandler.exportWithCheckIns(
        activeSheetName,
        rows
      );
      
      // Create and download
      const blob = new Blob([exportBuffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });

      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      link.setAttribute('href', url);
      link.setAttribute('download', 'attendee_report_updated.xlsx');
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      toast({
        title: "Export Successful! ✨",
        description: "The updated attendee report has been downloaded with enhanced formatting preservation.",
      });
      
    } catch (error) {
      console.error("Error exporting Excel file:", error);
      toast({
        variant: "destructive",
        title: "Export Error",
        description: "Could not export the Excel file: " + (error instanceof Error ? error.message : String(error)),
      });
    }
  };


  return (
    <div className="flex min-h-screen w-full flex-col bg-muted/40" suppressHydrationWarning>
      <header className="sticky top-0 z-30 flex h-14 items-center gap-4 border-b bg-background px-4 sm:static sm:h-auto sm:border-0 sm:bg-transparent sm:px-6">
        <div className="flex items-center gap-2">
            <TicketCheck className="h-6 w-6 text-primary" />
            <h1 className="text-xl font-bold">TicketCheck Pro</h1>
        </div>
        <div className="ml-auto flex items-center gap-4">
            {dataSource === 'google-sheets' ? (
              <div className="flex items-center gap-2 text-sm text-green-600">
                <CheckCircle className="h-4 w-4" />
                Auto-saving to Google Sheets
              </div>
            ) : (
              <Button onClick={handleExport} variant="outline" size="sm" disabled={rows.length === 0}>
                  <Download className="mr-2 h-4 w-4"/>
                  Export Report
              </Button>
            )}
            <Button onClick={handleLogout} variant="outline" size="sm">
              <LogOut className="mr-2 h-4 w-4" />
              Logout
            </Button>
        </div>
      </header>
      <main className="flex-1 p-4 sm:px-6 sm:py-0">
        <div className="grid auto-rows-max items-start gap-4 md:gap-8 lg:grid-cols-2 xl:grid-cols-3">
          <div className="grid auto-rows-max items-start gap-4 md:gap-8 lg:col-span-1">
            {!isGoogleSheetsConnected && (
              <Card>
                <CardHeader>
                  <CardTitle>Upload Attendee List</CardTitle>
                  <CardDescription>
                      Select an Excel file. If it has multiple sheets, you will be prompted to choose one.
                      <br />
                      <small className="text-muted-foreground">✓ All original formatting will be preserved when downloading the updated report.</small>
                  </CardDescription>
                </CardHeader>
                <CardContent>
                  <input
                      type="file"
                      ref={fileInputRef}
                      onChange={handleFileChange}
                      className="hidden"
                      accept=".xlsx,.xls,.xlsm,.xlsb"
                      title="Upload Excel file (.xlsx, .xls, .xlsm, .xlsb)"
                  />
                  <Button onClick={() => fileInputRef.current?.click()} className="w-full">
                      <Upload className="mr-2 h-4 w-4" />
                      Upload Excel File
                  </Button>
                </CardContent>
              </Card>
            )}
            
            <GoogleSheetsConnector 
              onDataLoaded={handleGoogleSheetsDataLoaded}
              onConnectionChange={handleGoogleSheetsConnectionChange}
              googleSheetsApi={googleSheetsApi}
            />
            
            <Card>
              <CardHeader>
                <CardTitle>Check In Attendee</CardTitle>
                <CardDescription>Enter a unique code or scan a QR code to check in an attendee.</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="relative mb-4 flex aspect-square w-full items-center justify-center overflow-hidden rounded-lg border-2 border-dashed bg-muted">
                    <video ref={videoRef} className="absolute inset-0 h-full w-full object-cover" autoPlay playsInline muted />
                    
                    {isScanning ? (
                        <>
                            <div 
                                className="absolute inset-0 z-10" 
                                style={{ boxShadow: '0 0 0 9999px rgba(0, 0, 0, 0.4)' }}
                            >
                                <div className="pointer-events-none absolute top-1/2 left-1/2 h-3/4 w-3/4 max-w-[400px] max-h-[400px] -translate-x-1/2 -translate-y-1/2 rounded-lg" />
                            </div>

                            <div className="pointer-events-none relative z-20 h-3/4 w-3/4 max-w-[400px] max-h-[400px]">
                                <div className="absolute -top-1 -left-1 h-10 w-10 border-t-4 border-l-4 border-primary rounded-tl-lg"></div>
                                <div className="absolute -top-1 -right-1 h-10 w-10 border-t-4 border-r-4 border-primary rounded-tr-lg"></div>
                                <div className="absolute -bottom-1 -left-1 h-10 w-10 border-b-4 border-l-4 border-primary rounded-bl-lg"></div>
                                <div className="absolute -bottom-1 -right-1 h-10 w-10 border-b-4 border-r-4 border-primary rounded-br-lg"></div>
                            </div>
                        </>
                    ) : (
                        <div className="absolute inset-0 flex items-center justify-center bg-muted">
                            <QrCode className="h-16 w-16 text-muted-foreground/50"/>
                        </div>
                    )}
                    <canvas ref={canvasRef} className="hidden" />
                </div>
                {scanError && (
                    <Alert variant="destructive" className="mb-4">
                        <AlertTriangle className="h-4 w-4" />
                        <AlertTitle>Camera Error</AlertTitle>
                        <AlertDescriptionUI>{scanError}</AlertDescriptionUI>
                    </Alert>
                )}
                <div className="flex items-center space-x-2 mb-4">
                    <Checkbox id="continuous-scan" checked={isContinuous} onCheckedChange={(checked) => setIsContinuous(!!checked)} />
                    <Label htmlFor="continuous-scan" className="cursor-pointer">Continuous Scan</Label>
                </div>
                <Button type="button" onClick={handleScanButtonClick} className="w-full mb-4" variant="outline" disabled={rows.length === 0}>
                    <Camera className="mr-2 h-4 w-4" />
                    {isScanning ? 'Stop Camera' : 'Scan QR Code'}
                </Button>
                <Form {...checkInForm}>
                    <form 
                        onSubmit={(e) => {
                            e.preventDefault();
                            scanSourceRef.current = 'form';
                            checkInForm.handleSubmit(handleCheckIn)(e);
                        }} 
                        className="space-y-4"
                    >
                        <FormField
                            control={checkInForm.control}
                            name="uniqueCode"
                            render={({ field }) => (
                                <FormItem>
                                    <FormLabel>Unique Code</FormLabel>
                                    <FormControl>
                                        <Input 
                                            ref={(e) => {
                                                field.ref(e);
                                                (inputRef as React.MutableRefObject<HTMLInputElement | null>).current = e;
                                            }}
                                            placeholder="Paste or type code here..." 
                                            value={field.value}
                                            onChange={field.onChange}
                                            onBlur={field.onBlur}
                                            name={field.name}
                                            disabled={rows.length === 0} 
                                        />
                                    </FormControl>
                                    <FormMessage />
                                </FormItem>
                            )}
                        />
                        <Button type="submit" className="w-full bg-accent hover:bg-accent/90" disabled={rows.length === 0}>
                           <TicketCheck className="mr-2 h-4 w-4"/> Check In
                        </Button>
                    </form>
                </Form>
              </CardContent>
            </Card>
          </div>
          <div className="lg:col-span-1 xl:col-span-2">
            <Card>
              <CardHeader>
                  <CardTitle>Attendee List</CardTitle>
                  <CardDescription>A list of all imported attendees and their check-in status.</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="max-h-[600px] overflow-auto">
                    <Table>
                        <TableHeader>
                            <TableRow>
                                {headers.map(header => <TableHead key={header}>{header}</TableHead>)}
                                {headers.length > 0 && (
                                    <>
                                        <TableHead>Status</TableHead>
                                        <TableHead>Checked In At</TableHead>
                                    </>
                                )}
                            </TableRow>
                        </TableHeader>
                        <TableBody>
                            {rows.length > 0 ? (
                                rows.map((row, rowIndex) => (
                                    <TableRow 
                                      key={rowIndex}
                                      ref={(el) => {
                                          if (el) rowRefs.current[rowIndex] = el;
                                      }}
                                      className={cn(
                                        highlightedRowIndex === rowIndex && 'bg-primary/10'
                                      )}
                                    >
                                        {headers.map(header => (
                                            <TableCell key={header}>
                                                {String(row[header] ?? '')}
                                            </TableCell>
                                        ))}
                                        <TableCell>
                                            <Badge variant={row.checkedInTime ? "default" : "secondary"} className={row.checkedInTime ? "bg-accent text-accent-foreground" : ""}>
                                                {row.checkedInTime ? "Checked In" : "Pending"}
                                            </Badge>
                                        </TableCell>
                                        <TableCell suppressHydrationWarning>
                                            {formatDateSafe(row.checkedInTime)}
                                        </TableCell>
                                    </TableRow>
                                ))
                            ) : (
                                <TableRow>
                                    <TableCell colSpan={headers.length > 0 ? headers.length + 2 : 1} className="h-24 text-center">
                                        No data loaded. Please upload an Excel file.
                                    </TableCell>
                                </TableRow>
                            )}
                        </TableBody>
                    </Table>
                </div>
              </CardContent>
            </Card>
          </div>
        </div>
      </main>

      <Dialog open={isSheetSelectorOpen} onOpenChange={setIsSheetSelectorOpen}>
        <DialogContent>
            <DialogHeader>
            <DialogTitle>Select a Sheet</DialogTitle>
            <DialogDescription>
                Your Excel file contains multiple sheets. Please select the one you'd like to import.
            </DialogDescription>
            </DialogHeader>
            <div className="flex flex-col space-y-2">
            {sheetNames.map(name => (
                <Button
                key={name}
                variant="outline"
                onClick={() => handleSheetSelect(name)}
                >
                <FileSpreadsheet className="mr-2 h-4 w-4" />
                {name}
                </Button>
            ))}
            </div>
        </DialogContent>
      </Dialog>

      <AlertDialog open={isAlertOpen} onOpenChange={(open) => {
        if (!open) {
          handleAlertClose();
        }
        setIsAlertOpen(open);
      }}>
        <AlertDialogContent>
          {dialogState === 'success' && scannedRow && (
            <>
              <AlertDialogHeader className="bg-accent text-accent-foreground p-4 -mx-6 -mt-6 sm:rounded-t-lg">
                <AlertDialogTitle className="flex items-center gap-2">
                  <UserCheck className="h-6 w-6" />
                  Check-in Successful!
                </AlertDialogTitle>
                <AlertDialogDescription className="text-accent-foreground/90">
                    Welcome! Details for the attendee are below.
                </AlertDialogDescription>
              </AlertDialogHeader>
              <div className="text-sm space-y-1 max-h-60 overflow-auto">
                {headers.map((header) => (
                    <p key={header}><strong>{header}:</strong> {String(scannedRow[header] ?? '')}</p>
                ))}
                <p>
                  <strong>Checked-in:</strong>{" "}
                  <span suppressHydrationWarning>{formatDateSafe(scannedRow.checkedInTime)}</span>
                </p>
              </div>
            </>
          )}
          {dialogState === 'duplicate' && scannedRow && (
            <>
              <AlertDialogHeader className="bg-destructive text-destructive-foreground p-4 -mx-6 -mt-6 sm:rounded-t-lg">
                <AlertDialogTitle className="flex items-center gap-2">
                  <AlertTriangle className="h-8 w-8" />
                  Already Checked In!
                </AlertDialogTitle>
                <AlertDialogDescription className="text-destructive-foreground/90">
                  This ticket has already been scanned. Please verify the attendee.
                </AlertDialogDescription>
              </AlertDialogHeader>
              <div className="text-sm space-y-2 max-h-60 overflow-auto rounded-md border bg-muted p-4">
                <p className="font-bold text-lg mb-2">Original Check-in Details:</p>
                {headers.map((header) => (
                    <p key={header}><strong>{header}:</strong> {String(scannedRow[header] ?? '')}</p>
                ))}
                <p className="mt-2">
                  <strong>Initial Check-in Time:</strong>{" "}
                  <span suppressHydrationWarning>{formatDateSafe(scannedRow.checkedInTime)}</span>
                </p>
              </div>
            </>
          )}
          {dialogState === 'not_found' && (
            <AlertDialogHeader className="bg-yellow-400 text-yellow-900 p-4 -mx-6 -mt-6 sm:rounded-t-lg">
                <AlertDialogTitle className="flex items-center gap-2">
                    <AlertTriangle className="h-6 w-6" />
                    Ticket Not Found
                </AlertDialogTitle>
                <AlertDialogDescription className="text-yellow-900/90">
                    The scanned code does not match any entry in the list. Please try again.
                </AlertDialogDescription>
            </AlertDialogHeader>
          )}
          <AlertDialogFooter>
            <AlertDialogAction onClick={handleAlertClose}>
              Close
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}

