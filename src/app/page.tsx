
"use client";

import * as React from "react";
import { z } from "zod";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { format } from "date-fns";
import * as XLSX from "xlsx";
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
  FileSpreadsheet
} from "lucide-react";
import { cn } from "@/lib/utils";

// Helper functions for Excel formatting preservation
const deepCopyWorkbook = (workbook: WorkBook): WorkBook => {
  const newWorkbook = XLSX.utils.book_new();
  
  try {
    // Copy all sheets with complete formatting
    Object.keys(workbook.Sheets).forEach(sheetName => {
      const originalSheet = workbook.Sheets[sheetName];
      const newSheet: WorkSheet = {};
      
      if (!originalSheet) return;
      
      // Copy all cells and sheet properties
      Object.keys(originalSheet).forEach(key => {
        try {
          if (key.startsWith('!')) {
            // Copy sheet properties (merges, cols, rows, etc.)
            const prop = originalSheet[key];
            if (Array.isArray(prop)) {
              newSheet[key] = [...prop];
            } else if (prop && typeof prop === 'object') {
              newSheet[key] = { ...prop };
            } else {
              newSheet[key] = prop;
            }
          } else {
            // Copy cell with all formatting properties
            const cell = originalSheet[key];
            if (cell && typeof cell === 'object') {
              newSheet[key] = { ...cell };
            }
          }
        } catch (error) {
          console.warn(`Error copying cell ${key}:`, error);
        }
      });
      
      newWorkbook.Sheets[sheetName] = newSheet;
    });
    
    // Copy workbook metadata safely
    newWorkbook.SheetNames = [...workbook.SheetNames];
    if (workbook.Props) newWorkbook.Props = { ...workbook.Props };
    if (workbook.Custprops) newWorkbook.Custprops = { ...workbook.Custprops };
    if (workbook.Workbook) newWorkbook.Workbook = { ...workbook.Workbook };
    if (workbook.vbaraw) newWorkbook.vbaraw = workbook.vbaraw;
  } catch (error) {
    console.error('Error in deepCopyWorkbook:', error);
    // Fallback: return original workbook if copy fails
    return workbook;
  }
  
  return newWorkbook;
};

const detectFileFormatting = (workbook: WorkBook): boolean => {
  try {
    if (!workbook || !workbook.SheetNames) {
      return false;
    }

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      
      if (!sheet || typeof sheet !== 'object') continue;
      
      // Check for formatting indicators
      if (sheet['!cols'] || sheet['!rows'] || sheet['!merges']) {
        return true;
      }
      
      // Check cells for styles - only check actual cell addresses (like A1, B2, etc.)
      try {
        for (const cellAddress of Object.keys(sheet)) {
          if (!cellAddress.startsWith('!')) {
            try {
              const cell = sheet[cellAddress];
              // Safely check if cell exists and has formatting properties
              if (cell && typeof cell === 'object' && 
                  cell !== null && 
                  !Array.isArray(cell) &&
                  (cell.hasOwnProperty('s') || cell.hasOwnProperty('z'))) {
                return true;
              }
            } catch (cellError) {
              console.warn(`Error checking cell ${cellAddress}:`, cellError);
              continue;
            }
          }
        }
      } catch (sheetError) {
        console.warn(`Error checking sheet ${sheetName}:`, sheetError);
        continue;
      }
    }
  } catch (error) {
    console.warn('Error detecting file formatting:', error);
    // If we can't detect formatting, assume it has some
    return true;
  }
  return false;
};

const preserveCellFormatting = (sourceCell: CellObject | undefined, newValue: any, cellType: ExcelDataType = 's') => {
  const newCell: CellObject = {
    v: newValue,
    t: cellType,
  };
  
  if (sourceCell && typeof sourceCell === 'object') {
    try {
      // Preserve all formatting properties safely
      if (sourceCell.s) newCell.s = sourceCell.s; // Style
      if (sourceCell.z) newCell.z = sourceCell.z; // Number format
      if (sourceCell.l) newCell.l = sourceCell.l; // Hyperlink
      if (sourceCell.c) newCell.c = sourceCell.c; // Comments
      if (sourceCell.w) newCell.w = sourceCell.w; // Formatted text
    } catch (error) {
      console.warn('Error preserving cell formatting:', error);
    }
  }
  
  return newCell;
};

const checkInSchema = z.object({
  uniqueCode: z.string().min(1, { message: "Code is required." }),
});

type DialogState = 'success' | 'duplicate' | 'not_found';

export default function DashboardPage() {
  const { toast } = useToast();
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
  const [sheetNames, setSheetNames] = React.useState<string[]>([]);
  const [activeSheetName, setActiveSheetName] = React.useState<string | null>(null);
  const [isSheetSelectorOpen, setIsSheetSelectorOpen] = React.useState(false);
  const [highlightedRowIndex, setHighlightedRowIndex] = React.useState<number | null>(null);
  const [uploadedFileName, setUploadedFileName] = React.useState<string | null>(null);
  const [uploadedFileInfo, setUploadedFileInfo] = React.useState<{
    name: string;
    size: string;
    type: string;
    hasFormatting: boolean;
  } | null>(null);
  
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
        const updatedRow = { ...foundRowData, checkedInTime: new Date(), __rowNum__: foundRowData.__rowNum__ };
        const updatedRows = [...rows];
        updatedRows[foundRowIndex] = updatedRow;
        setRows(updatedRows);
        setScannedRow(updatedRow);
        setDialogState('success');
        setIsAlertOpen(true);
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

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    // Set file info for display
    setUploadedFileInfo({
      name: file.name,
      size: (file.size / 1024).toFixed(1) + ' KB',
      type: file.name.split('.').pop()?.toUpperCase() || 'Unknown',
      hasFormatting: true // Will be updated after reading
    });

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, {
            type: "array",
            cellStyles: true,
            cellDates: true,
            bookVBA: true,
            bookSheets: true,
            bookProps: true,
            sheetStubs: true,
            sheetRows: 0, // Read all rows
            dense: false,
        });
        const names = wb.SheetNames;
        
        setWorkbook(wb);
        setSheetNames(names);
        setActiveSheetName(null);
        setHeaders([]);
        setRows([]);
        setScannedRow(null);
        setHighlightedRowIndex(null);
        rowRefs.current = [];

        // Update file info with actual formatting detection
        try {
          setUploadedFileInfo(prev => prev ? {
            ...prev,
            hasFormatting: detectFileFormatting(wb)
          } : null);
        } catch (error) {
          console.warn('Error detecting formatting, assuming file has formatting:', error);
          setUploadedFileInfo(prev => prev ? {
            ...prev,
            hasFormatting: true
          } : null);
        }

        if (names.length === 0) {
            toast({
                variant: "destructive",
                title: "No Sheets Found",
                description: "The uploaded Excel file does not contain any sheets.",
            });
            setUploadedFileInfo(null);
            return;
        }

        if (names.length === 1) {
            processSheetData(wb, names[0]);
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
        setUploadedFileInfo(null);
      }
    };
    reader.onerror = () => {
        toast({
            variant: "destructive",
            title: "File Read Error",
            description: "There was an error reading the file."
        });
        setUploadedFileInfo(null);
    }
    reader.readAsArrayBuffer(file);
    event.target.value = '';
  };
  
  const handleSheetSelect = (sheetName: string) => {
    if (sheetName && workbook) {
        processSheetData(workbook, sheetName);
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

  const handleExport = () => {
    if (!workbook || !activeSheetName) {
        toast({
            variant: "destructive",
            title: "No Workbook Found",
            description: "Please upload an Excel file first."
        });
        return;
    }

    try {
        // Create a deep copy of the workbook to preserve all original formatting
        const workbookCopy = deepCopyWorkbook(workbook);

        // Work with the copied worksheet
        const ws = workbookCopy.Sheets[activeSheetName];
        if (!ws) {
            toast({
              variant: "destructive",
              title: "Sheet Error",
              description: `Could not find sheet "${activeSheetName}" in the workbook.`
            });
            return;
        }
        
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
        const checkInColIndex = range.e.c + 1;
        const checkInColName = XLSX.utils.encode_col(checkInColIndex);
        
        // Add header with proper formatting (copy style from adjacent header)
        const headerAddress = `${checkInColName}1`;
        const adjacentHeaderAddress = range.e.c > 0 ? `${XLSX.utils.encode_col(range.e.c)}1` : 'A1';
        const adjacentHeaderCell = ws[adjacentHeaderAddress];
        
        try {
          ws[headerAddress] = preserveCellFormatting(adjacentHeaderCell, 'Checked-In At', 's');
        } catch (error) {
          console.warn('Error setting header formatting:', error);
          ws[headerAddress] = { v: 'Checked-In At', t: 's' };
        }

        // Update data cells while preserving row formatting
        rows.forEach(row => {
            if (row.checkedInTime && row.__rowNum__) {
                const cellAddress = `${checkInColName}${row.__rowNum__}`;
                const cellValue = format(new Date(row.checkedInTime), 'yyyy-MM-dd HH:mm:ss');
                
                // Get formatting template from adjacent cell in the same row
                const adjacentCellAddress = range.e.c > 0 ? `${XLSX.utils.encode_col(range.e.c)}${row.__rowNum__}` : `A${row.__rowNum__}`;
                const adjacentCell = ws[adjacentCellAddress];
                
                try {
                  // Create cell with datetime formatting
                  const newCell = preserveCellFormatting(adjacentCell, cellValue, 's');
                  // Add specific datetime format
                  newCell.z = 'yyyy-mm-dd hh:mm:ss';
                  ws[cellAddress] = newCell;
                } catch (error) {
                  console.warn(`Error formatting cell ${cellAddress}:`, error);
                  ws[cellAddress] = { v: cellValue, t: 's' };
                }
            }
        });
        
        // Update sheet range
        if (ws['!ref']) {
            const newRange = XLSX.utils.decode_range(ws['!ref']);
            newRange.e.c = Math.max(newRange.e.c, checkInColIndex);
            ws['!ref'] = XLSX.utils.encode_range(newRange);
        }
        
        // Preserve and extend column widths
        if (workbook.Sheets[activeSheetName]['!cols']) {
            ws['!cols'] = [...workbook.Sheets[activeSheetName]['!cols']];
            // Add reasonable width for new datetime column
            if (!ws['!cols'][checkInColIndex]) {
                ws['!cols'][checkInColIndex] = { width: 20 };
            }
        } else {
            // Create column widths array if it doesn't exist
            ws['!cols'] = [];
            ws['!cols'][checkInColIndex] = { width: 20 };
        }
        
        // Preserve row heights
        if (workbook.Sheets[activeSheetName]['!rows']) {
            ws['!rows'] = [...workbook.Sheets[activeSheetName]['!rows']];
        }
        
        // Write with maximum formatting preservation
        const excelBuffer = XLSX.write(workbookCopy, { 
            bookType: 'xlsx', 
            type: 'array', 
            cellStyles: true,
            cellDates: true,
            bookVBA: true,
            compression: true,
        });
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8' });

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
            description: "The updated attendee report has been downloaded with all original formatting preserved.",
        });
    } catch (error) {
        console.error("Error exporting Excel file:", error);
        toast({
          variant: "destructive",
          title: "Export Error",
          description: "Could not export the Excel file.",
        });
    }
  };


  return (
    <div className="flex min-h-screen w-full flex-col bg-muted/40">
      <header className="sticky top-0 z-30 flex h-14 items-center gap-4 border-b bg-background px-4 sm:static sm:h-auto sm:border-0 sm:bg-transparent sm:px-6">
        <div className="flex items-center gap-2">
            <TicketCheck className="h-6 w-6 text-primary" />
            <h1 className="text-xl font-bold">TicketCheck Pro</h1>
        </div>
        <div className="ml-auto flex items-center gap-4">
            <Button onClick={handleExport} variant="outline" size="sm" disabled={rows.length === 0}>
                <Download className="mr-2 h-4 w-4"/>
                Export Report
            </Button>
        </div>
      </header>
      <main className="flex-1 p-4 sm:px-6 sm:py-0">
        <div className="grid auto-rows-max items-start gap-4 md:gap-8 lg:grid-cols-2 xl:grid-cols-3">
          <div className="grid auto-rows-max items-start gap-4 md:gap-8 lg:col-span-1">
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
            
            {uploadedFileInfo && (
              <Card>
                <CardHeader>
                  <CardTitle className="flex items-center text-sm">
                    <FileSpreadsheet className="mr-2 h-4 w-4 text-green-600" />
                    File Uploaded Successfully
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-2">
                  <div className="text-sm">
                    <strong>File:</strong> {uploadedFileInfo.name}
                  </div>
                  <div className="text-sm">
                    <strong>Size:</strong> {uploadedFileInfo.size}
                  </div>
                  <div className="text-sm">
                    <strong>Type:</strong> {uploadedFileInfo.type}
                  </div>
                  {uploadedFileInfo.hasFormatting && (
                    <div className="flex items-center gap-1 text-sm text-green-600">
                      <div className="w-2 h-2 bg-green-500 rounded-full"></div>
                      Formatting preservation enabled
                    </div>
                  )}
                  {activeSheetName && (
                    <div className="text-sm">
                      <strong>Active Sheet:</strong> {activeSheetName}
                    </div>
                  )}
                  {rows.length > 0 && (
                    <div className="text-sm">
                      <strong>Records:</strong> {rows.length} attendees loaded
                    </div>
                  )}
                </CardContent>
              </Card>
            )}
            
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
                                        <TableCell>
                                            {row.checkedInTime ? format(row.checkedInTime, 'PPpp') : 'N/A'}
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
                  {scannedRow.checkedInTime ? format(new Date(scannedRow.checkedInTime), 'PPpp') : 'N/A'}
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
                  {scannedRow.checkedInTime ? format(new Date(scannedRow.checkedInTime), 'PPpp') : 'N/A'}
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
