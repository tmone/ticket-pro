"use client";

import * as React from "react";
import { useRouter } from "next/navigation";
import { z } from "zod";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { format } from "date-fns";
import * as XLSX from "xlsx";
import jsQR from "jsqr";

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
import {
  TicketCheck,
  LogOut,
  Link,
  QrCode,
  Download,
  UserCheck,
  AlertTriangle,
  Camera,
  RefreshCw,
} from "lucide-react";
import { fetchGoogleSheetData } from "./actions";

const checkInSchema = z.object({
  uniqueCode: z.string().min(1, { message: "Code is required." }),
});

const sheetUrlSchema = z.object({
    url: z.string().url({ message: "Please enter a valid Google Sheet URL." }),
});

type DialogState = 'success' | 'duplicate' | 'not_found';

export default function DashboardPage() {
  const router = useRouter();
  const { toast } = useToast();
  const videoRef = React.useRef<HTMLVideoElement>(null);
  const canvasRef = React.useRef<HTMLCanvasElement>(null);

  const [isAuthenticated, setIsAuthenticated] = React.useState(false);
  const [headers, setHeaders] = React.useState<string[]>([]);
  const [rows, setRows] = React.useState<Record<string, any>[]>([]);
  const [scannedRow, setScannedRow] = React.useState<Record<string, any> | null | undefined>(null);
  const [isAlertOpen, setIsAlertOpen] = React.useState(false);
  const [dialogState, setDialogState] = React.useState<DialogState>('not_found');
  
  const [isScanning, setIsScanning] = React.useState(false);
  const [scanError, setScanError] = React.useState<string | null>(null);
  const [isContinuous, setIsContinuous] = React.useState(false);
  const [googleSheetUrl, setGoogleSheetUrl] = React.useState<string>("");
  const [isFetching, setIsFetching] = React.useState(false);

  const checkInForm = useForm<z.infer<typeof checkInSchema>>({
    resolver: zodResolver(checkInSchema),
    defaultValues: { uniqueCode: "" },
  });

  const sheetUrlForm = useForm<z.infer<typeof sheetUrlSchema>>({
    resolver: zodResolver(sheetUrlSchema),
    defaultValues: { url: "" },
  });
  
  const stopScan = React.useCallback(() => {
    setIsScanning(false);
    if (videoRef.current && videoRef.current.srcObject) {
      const stream = videoRef.current.srcObject as MediaStream;
      stream.getTracks().forEach((track) => track.stop());
      videoRef.current.srcObject = null;
    }
  }, []);
  
  const startScan = React.useCallback(async () => {
    setScanError(null);
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } });
      if (videoRef.current) {
        videoRef.current.srcObject = stream;
        await videoRef.current.play();
        setIsScanning(true);
      }
    } catch (err) {
      console.error("Camera access error:", err);
      setScanError("Camera access denied. Please enable it in your browser settings.");
      setIsScanning(false);
    }
  }, []);

  React.useEffect(() => {
    const authStatus = sessionStorage.getItem("isAuthenticated");
    if (authStatus !== "true") {
      router.push("/login");
    } else {
      setIsAuthenticated(true);
    }
    
    // Cleanup camera on component unmount
    return () => {
      stopScan();
    };
  }, [router, stopScan]);

  const handleLogout = () => {
    sessionStorage.removeItem("isAuthenticated");
    router.push("/login");
  };

  const processSheetData = (jsonData: Record<string, any>[]) => {
      if (jsonData.length === 0) {
        throw new Error("No data found in the Google Sheet.");
      }
      
      const firstRow = jsonData[0];
      const extractedHeaders = Object.keys(firstRow);
      
      const initialRows = jsonData.map(row => ({...row, checkedInTime: null}));

      setHeaders(extractedHeaders);
      setRows(initialRows);
      setScannedRow(null);

      toast({
        title: "Success!",
        description: `Successfully imported ${jsonData.length} rows.`,
      });
  };

  const handleGoogleSheetFetch = async (data: z.infer<typeof sheetUrlSchema>) => {
    setIsFetching(true);
    setGoogleSheetUrl(data.url);
    try {
      const result = await fetchGoogleSheetData(data.url);
      if (result.error) {
        throw new Error(result.error);
      }
      processSheetData(result.data!);
    } catch (error: any) {
      toast({
        variant: "destructive",
        title: "Import Failed",
        description: error.message || "Could not fetch or process data. Please check the URL and sheet permissions.",
      });
    } finally {
      setIsFetching(false);
    }
  };


  const handleCheckIn = React.useCallback((data: z.infer<typeof checkInSchema>) => {
    const { uniqueCode } = data;
    let searchCode = uniqueCode.trim();

    try {
      const url = new URL(searchCode);
      const firstParam = url.searchParams.values().next().value;
      if (firstParam) {
        searchCode = firstParam.trim();
      }
    } catch (e) {
      // Not a URL, continue with the original code
    }

    const rowIndex = rows.findIndex(row =>
      Object.values(row).some(
        cellValue => String(cellValue).trim().toLowerCase() === searchCode.toLowerCase()
      )
    );

    if (rowIndex !== -1) {
      const foundRow = rows[rowIndex];
      const now = new Date();
      
      if (foundRow.checkedInTime) {
        const checkedInDate = new Date(foundRow.checkedInTime);
        const secondsSinceCheckIn = (now.getTime() - checkedInDate.getTime()) / 1000;

        if (secondsSinceCheckIn < 60) {
          // Within 60s grace period, treat as success
          setScannedRow(foundRow);
          setDialogState('success');
        } else {
          // After 60s, it's a duplicate
          setScannedRow(foundRow);
          setDialogState('duplicate');
        }
      } else {
        // First time check-in
        const updatedRow = { ...foundRow, checkedInTime: now };
        const updatedRows = [...rows];
        updatedRows[rowIndex] = updatedRow;
        setRows(updatedRows);
        setScannedRow(updatedRow);
        setDialogState('success');
      }
    } else {
      setScannedRow(undefined);
      setDialogState('not_found');
    }

    setIsAlertOpen(true);
  }, [rows]);

  const tick = React.useCallback(() => {
    if (isScanning && videoRef.current && videoRef.current.readyState === videoRef.current.HAVE_ENOUGH_DATA && canvasRef.current) {
        const video = videoRef.current;
        const canvas = canvasRef.current;
        const ctx = canvas.getContext("2d");

        if (ctx) {
            canvas.height = video.videoHeight;
            canvas.width = video.videoWidth;
            ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
            const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
            const code = jsQR(imageData.data, imageData.width, imageData.height);

            if (code && code.data) {
                stopScan();
                checkInForm.setValue("uniqueCode", code.data, { shouldValidate: true });
                checkInForm.handleSubmit(handleCheckIn)();
                return;
            }
        }
    }
    requestAnimationFrame(tick);
  }, [isScanning, stopScan, checkInForm, handleCheckIn]);

  React.useEffect(() => {
    if (isScanning) {
      const animationFrameId = requestAnimationFrame(tick);
      return () => cancelAnimationFrame(animationFrameId);
    }
  }, [isScanning, tick]);

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
    if (link.href) {
        URL.revokeObjectURL(link.href);
    }
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', 'attendee_report.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    toast({
        title: "Export Successful",
        description: "The attendee report has been downloaded."
    });
  };
  
  const handleAlertClose = React.useCallback(() => {
      setIsAlertOpen(false);
      checkInForm.reset();
      // If continuous mode is on and we just closed an error/warning dialog, restart scanning.
      if (isContinuous && (dialogState === 'duplicate' || dialogState === 'not_found')) {
           // Use a timeout to defer the startScan call, allowing the state to update properly first.
           setTimeout(() => startScan(), 0);
      }
  }, [isContinuous, dialogState, checkInForm, startScan]);

  React.useEffect(() => {
      // For successful scans in continuous mode, auto-close the dialog and restart scanning.
      if (isContinuous && dialogState === 'success' && isAlertOpen) {
          const timer = setTimeout(() => {
              setIsAlertOpen(false);
              checkInForm.reset();
              startScan();
          }, 1000); // 1-second delay

          return () => clearTimeout(timer); // Cleanup timer
      }
  }, [isAlertOpen, dialogState, isContinuous, startScan, checkInForm]);


  if (!isAuthenticated) {
    return <div className="flex h-screen w-full items-center justify-center">Loading...</div>;
  }
  
  return (
    <div className="flex min-h-screen w-full flex-col bg-muted/40">
      <header className="sticky top-0 z-30 flex h-14 items-center gap-4 border-b bg-background px-4 sm:static sm:h-auto sm:border-0 sm:bg-transparent sm:px-6">
        <div className="flex items-center gap-2">
            <TicketCheck className="h-6 w-6 text-primary" />
            <h1 className="text-xl font-bold">TicketCheck Pro</h1>
        </div>
        <div className="ml-auto flex items-center gap-2">
            <Button onClick={handleExport} variant="outline" size="sm">
                <Download className="mr-2 h-4 w-4"/>
                Export Report
            </Button>
            <Button onClick={handleLogout} variant="ghost" size="icon">
                <LogOut className="h-5 w-5" />
            </Button>
        </div>
      </header>
      <main className="flex-1 p-4 sm:px-6 sm:py-0">
        <div className="grid auto-rows-max items-start gap-4 md:gap-8 lg:grid-cols-2 xl:grid-cols-3">
          <div className="grid auto-rows-max items-start gap-4 md:gap-8 lg:col-span-1">
            <Card>
              <CardHeader>
                <CardTitle>Google Sheet Data</CardTitle>
                <CardDescription>
                    Paste the URL of a public Google Sheet. Make sure it's shared with "Anyone with the link".
                </CardDescription>
              </CardHeader>
              <CardContent>
                <Form {...sheetUrlForm}>
                    <form onSubmit={sheetUrlForm.handleSubmit(handleGoogleSheetFetch)} className="space-y-4">
                        <FormField
                            control={sheetUrlForm.control}
                            name="url"
                            render={({ field }) => (
                                <FormItem>
                                    <FormLabel>Sheet URL</FormLabel>
                                    <FormControl>
                                        <Input placeholder="https://docs.google.com/spreadsheets/d/..." {...field} />
                                    </FormControl>
                                    <FormMessage />
                                </FormItem>
                            )}
                        />
                        <Button type="submit" className="w-full" disabled={isFetching}>
                           {isFetching ? <RefreshCw className="mr-2 h-4 w-4 animate-spin"/> : <Link className="mr-2 h-4 w-4" />}
                           {isFetching ? 'Fetching Data...' : 'Load Data from Sheet'}
                        </Button>
                    </form>
                </Form>
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
                    <Label htmlFor="continuous-scan" className="cursor-pointer">Liên tục</Label>
                </div>
                <Button type="button" onClick={handleScanButtonClick} className="w-full mb-4" variant="outline">
                    <Camera className="mr-2 h-4 w-4" />
                    {isScanning ? 'Stop Camera' : 'Scan QR Code'}
                </Button>
                <Form {...checkInForm}>
                    <form onSubmit={checkInForm.handleSubmit(handleCheckIn)} className="space-y-4">
                        <FormField
                            control={checkInForm.control}
                            name="uniqueCode"
                            render={({ field }) => (
                                <FormItem>
                                    <FormLabel>Unique Code</FormLabel>
                                    <FormControl>
                                        <Input placeholder="Paste or type code here..." {...field} />
                                    </FormControl>
                                    <FormMessage />
                                </FormItem>
                            )}
                        />
                        <Button type="submit" className="w-full bg-accent hover:bg-accent/90">
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
                  <CardTitle>Data List</CardTitle>
                  <CardDescription>A list of all imported rows and their check-in status.</CardDescription>
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
                                    <TableRow key={rowIndex}>
                                        {headers.map(header => <TableCell key={header}>{String(row[header] ?? '')}</TableCell>)}
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
                                        No data loaded yet. Please provide a Google Sheet URL.
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

      <AlertDialog open={isAlertOpen} onOpenChange={setIsAlertOpen}>
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
                {Object.entries(scannedRow).filter(([key]) => key !== 'checkedInTime').map(([key, value]) => (
                    <p key={key}><strong>{key}:</strong> {String(value)}</p>
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
                {Object.entries(scannedRow).filter(([key]) => key !== 'checkedInTime').map(([key, value]) => (
                    <p key={key}><strong>{key}:</strong> {String(value)}</p>
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
