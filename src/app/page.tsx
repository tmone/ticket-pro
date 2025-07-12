
"use client";

import * as React from "react";
import { z } from "zod";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { format } from "date-fns";
import * as XLSX from "xlsx";
import type { WorkBook } from "xlsx";
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
import { Alert, AlertTitle, AlertDescription as AlertDialogDescriptionUI } from "@/components/ui/alert";
import { Form, FormControl, FormField, FormItem, FormLabel, FormMessage } from "@/components/ui/form";
import { Checkbox } from "@/components/ui/checkbox";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
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

const checkInSchema = z.object({
  uniqueCode: z.string().min(1, { message: "Code is required." }),
});

type DialogState = 'success' | 'duplicate' | 'not_found';

export default function DashboardPage() {
  const { toast } = useToast();
  const videoRef = React.useRef<HTMLVideoElement>(null);
  const canvasRef = React.useRef<HTMLCanvasElement>(null);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const inputRef = React.useRef<HTMLInputElement>(null);
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
  const [lastCheckedInCode, setLastCheckedInCode] = React.useState<string | null>(null);

  const [workbook, setWorkbook] = React.useState<WorkBook | null>(null);
  const [sheetNames, setSheetNames] = React.useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = React.useState<string>("");
  const [highlightedRowIndex, setHighlightedRowIndex] = React.useState<number | null>(null);
  
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
    
    // If continuous scanning from camera and the code is a duplicate of the last one, ignore it.
    if (isContinuous && scanSourceRef.current === 'camera' && uniqueCode === lastCheckedInCode) {
        // Just restart the scan without showing any dialog
        if (isContinuous && scanSourceRef.current === 'camera') {
            stopScan();
            setTimeout(() => startScan(), 100);
        }
        return; 
    }
    
    stopScan(); // Stop scanning as soon as we have a code to process.

    let codeToSearch = uniqueCode.trim();
    try {
        const url = new URL(codeToSearch);
        const params = url.searchParams;
        if (params.get("code")) {
            codeToSearch = params.get("code")!.trim();
        } else if (params.get("id")) {
             codeToSearch = params.get("id")!.trim();
        } else {
            const firstParam = params.values().next().value;
            if (firstParam) {
                codeToSearch = firstParam.trim();
            }
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
                if (params.get("code")) {
                    cellCode = params.get("code")!.trim();
                } else if (params.get("id")) {
                    cellCode = params.get("id")!.trim();
                } else {
                   const firstParam = params.values().next().value;
                   if (firstParam) {
                       cellCode = firstParam.trim();
                   }
                }
            } catch (e) { /* not a url */ }

            if (codeToSearch.toLowerCase() === cellCode.toLowerCase()) {
                foundRowIndex = i;
                break;
            }
        }
        if (foundRowIndex !== -1) {
            break;
        }
    }
    
    setHighlightedRowIndex(foundRowIndex !== -1 ? foundRowIndex : null);
    if(foundRowIndex !== -1) {
        rowRefs.current[foundRowIndex]?.scrollIntoView({
            behavior: 'smooth',
            block: 'center'
        })
    }

    if (foundRowIndex !== -1) {
      const foundRowData = rows[foundRowIndex];
      if (foundRowData.checkedInTime) {
        setScannedRow(foundRowData);
        setDialogState('duplicate');
        setIsAlertOpen(true);
      } else {
        const updatedRow = { ...foundRowData, checkedInTime: new Date() };
        const updatedRows = [...rows];
        updatedRows[foundRowIndex] = updatedRow;
        setRows(updatedRows);
        setScannedRow(updatedRow);
        setDialogState('success');
        setLastCheckedInCode(uniqueCode); 
        setIsAlertOpen(true);
      }
    } else {
      setScannedRow(undefined);
      setDialogState('not_found');
      setIsAlertOpen(true);
    }
  }, [rows, headers, isContinuous, lastCheckedInCode]);
  
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
                if (animationFrameIdRef.current) {
                  cancelAnimationFrame(animationFrameIdRef.current);
                  animationFrameIdRef.current = undefined;
                }
                scanSourceRef.current = 'camera';
                handleCheckIn({ uniqueCode: code.data });
                return; 
            }
        }
    }
    if (isScanning) {
        animationFrameIdRef.current = requestAnimationFrame(tick);
    }
  }, [handleCheckIn, isScanning]);
  
  const startScan = React.useCallback(async () => {
    setScanError(null);
    if (isScanning || animationFrameIdRef.current) return;
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } });
      if (videoRef.current) {
        videoRef.current.srcObject = stream;
        // Small delay to allow the stream to start before playing
        await new Promise(resolve => setTimeout(resolve, 100));
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

  const restartContinuousScan = React.useCallback(() => {
      stopScan();
      setTimeout(() => startScan(), 100);
  }, [startScan, stopScan]);

  React.useEffect(() => {
    return () => {
      stopScan();
    };
  }, [stopScan]);

  const processSheetData = (sheetName: string) => {
    if (!workbook) {
        toast({
            variant: "destructive",
            title: "Workbook not found",
            description: "Please upload an Excel file first."
        });
        return;
    }

    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
        toast({
            variant: "destructive",
            title: "Sheet not found",
            description: `Sheet with name "${sheetName}" could not be found in the file.`
        });
        return;
    }

    const jsonData = XLSX.utils.sheet_to_json<Record<string, any>>(worksheet);

    if (jsonData.length === 0) {
      toast({
          variant: "destructive",
          title: "No Data",
          description: "The selected sheet is empty or could not be read."
      });
      setHeaders([]);
      setRows([]);
      return;
    }
    
    const firstRow = jsonData[0];
    const extractedHeaders = Object.keys(firstRow).filter(h => h !== '__rowNum__');
    
    const initialRows = jsonData.map(row => ({...row, checkedInTime: null}));

    setHeaders(extractedHeaders);
    setRows(initialRows);
    setScannedRow(null);
    setSelectedSheet(sheetName);
    setHighlightedRowIndex(null);
    rowRefs.current = [];


    toast({
      title: "Success!",
      description: `Successfully imported ${jsonData.length} rows from sheet: ${sheetName}.`,
    });
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });
        setWorkbook(wb);
        setSheetNames(wb.SheetNames);

        // Reset previous data
        setHeaders([]);
        setRows([]);
        
        if (wb.SheetNames.length > 0) {
            setSelectedSheet(wb.SheetNames[0]);
            // Auto-process the first sheet
            const worksheet = wb.Sheets[wb.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json<Record<string, any>>(worksheet);

            if (jsonData.length > 0) {
              const firstRow = jsonData[0];
              const extractedHeaders = Object.keys(firstRow).filter(h => h !== '__rowNum__');
              const initialRows = jsonData.map(row => ({...row, checkedInTime: null}));
              setHeaders(extractedHeaders);
              setRows(initialRows);
              toast({
                title: "Success!",
                description: `Successfully imported ${jsonData.length} rows from sheet: ${wb.SheetNames[0]}.`,
              });
            } else {
              toast({
                variant: "destructive",
                title: "No Data",
                description: `The first sheet (${wb.SheetNames[0]}) is empty.`
              });
            }
        } else {
            toast({
                variant: "destructive",
                title: "No Sheets Found",
                description: "The uploaded Excel file does not contain any sheets.",
            });
            setSelectedSheet("");
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
        })
    }
    reader.readAsArrayBuffer(file);
    // Reset file input to allow uploading the same file again
    event.target.value = '';
  };
  
  const handleSheetSelect = (sheetName: string) => {
    if (sheetName) {
        processSheetData(sheetName);
    }
  };

  const handleScanButtonClick = () => {
    if (!isScanning) {
      startScan();
    } else {
      stopScan();
    }
  };

  const handleExport = () => {
    if (rows.length === 0) {
        toast({
            variant: "destructive",
            title: "No Data",
            description: "There is no data to export."
        });
        return;
    }

    const dataToExport = rows.map(row => {
        const newRow: Record<string, any> = {};
        for (const header of headers) {
            newRow[header] = row[header];
        }
        newRow['Checked-In At'] = row.checkedInTime ? format(row.checkedInTime, 'yyyy-MM-dd HH:mm:ss') : 'N/A';
        return newRow;
    });

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const csvContent = XLSX.utils.sheet_to_csv(worksheet);

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', 'attendee_report.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    toast({
        title: "Export Successful",
        description: "The attendee report has been downloaded."
    });
  };
  
  const handleAlertClose = React.useCallback(() => {
    setIsAlertOpen(false);
    checkInForm.reset();
    
    if (isContinuous) {
      if (scanSourceRef.current === 'camera') {
        restartContinuousScan();
      } else if (scanSourceRef.current === 'form') {
          inputRef.current?.focus();
      }
    }
    
  }, [isContinuous, checkInForm, restartContinuousScan]);

  React.useEffect(() => {
      if (isContinuous && dialogState === 'success' && isAlertOpen) {
          const timer = setTimeout(() => {
              handleAlertClose();
          }, 1500); 
          return () => clearTimeout(timer);
      }
  }, [isAlertOpen, dialogState, isContinuous, handleAlertClose]);

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
                    Select an Excel file. If it has multiple sheets, you can choose which one to use.
                </CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <input
                    type="file"
                    ref={fileInputRef}
                    onChange={handleFileChange}
                    className="hidden"
                    accept=".xlsx, .xls, .csv"
                />
                <Button onClick={() => fileInputRef.current?.click()} className="w-full">
                    <Upload className="mr-2 h-4 w-4" />
                    Upload Excel File
                </Button>
                {sheetNames.length > 1 && (
                    <div className="space-y-2">
                        <Label htmlFor="sheet-select">Select a sheet</Label>
                        <Select onValueChange={handleSheetSelect} value={selectedSheet}>
                            <SelectTrigger id="sheet-select">
                                <FileSpreadsheet className="mr-2 h-4 w-4" />
                                <SelectValue placeholder="Select a sheet" />
                            </SelectTrigger>
                            <SelectContent>
                                {sheetNames.map(name => (
                                    <SelectItem key={name} value={name}>{name}</SelectItem>
                                ))}
                            </SelectContent>
                        </Select>
                    </div>
                )}
              </CardContent>
            </Card>
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
                        <AlertDialogDescriptionUI>{scanError}</AlertDialogDescriptionUI>
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
                                            placeholder="Paste or type code here..." {...field} disabled={rows.length === 0} />
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

    