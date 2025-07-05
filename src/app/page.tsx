"use client";

import * as React from "react";
import { useRouter } from "next/navigation";
import { z } from "zod";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { format } from "date-fns";
import * as XLSX from "xlsx";

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
import { Form, FormControl, FormField, FormItem, FormLabel, FormMessage } from "@/components/ui/form";
import {
  TicketCheck,
  LogOut,
  Upload,
  QrCode,
  Download,
  UserCheck,
  AlertTriangle,
} from "lucide-react";

const checkInSchema = z.object({
  uniqueCode: z.string().min(1, { message: "Code is required." }),
});

type DialogState = 'success' | 'duplicate' | 'not_found';

export default function DashboardPage() {
  const router = useRouter();
  const { toast } = useToast();
  
  const [isAuthenticated, setIsAuthenticated] = React.useState(false);
  const [headers, setHeaders] = React.useState<string[]>([]);
  const [rows, setRows] = React.useState<Record<string, any>[]>([]);
  const [scannedRow, setScannedRow] = React.useState<Record<string, any> | null | undefined>(null);
  const [isAlertOpen, setIsAlertOpen] = React.useState(false);
  const [dialogState, setDialogState] = React.useState<DialogState>('not_found');
  
  const form = useForm<z.infer<typeof checkInSchema>>({
    resolver: zodResolver(checkInSchema),
    defaultValues: { uniqueCode: "" },
  });

  React.useEffect(() => {
    const authStatus = sessionStorage.getItem("isAuthenticated");
    if (authStatus !== "true") {
      router.push("/login");
    } else {
      setIsAuthenticated(true);
    }
  }, [router]);

  const handleLogout = () => {
    sessionStorage.removeItem("isAuthenticated");
    router.push("/login");
  };

  const handleDataUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    const fileName = file.name.toLowerCase();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) {
          throw new Error("Could not read file.");
        }
        
        if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls')) {
          throw new Error("Unsupported file type. Please upload an Excel file.");
        }

        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData: Record<string, any>[] = XLSX.utils.sheet_to_json(worksheet);

        if (jsonData.length === 0) {
            throw new Error("No data found in the Excel file.");
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
      } catch (error: any) {
        toast({
          variant: "destructive",
          title: "Import Failed",
          description: error.message || "Please check the file format and try again.",
        });
      }
    };
    
    reader.onerror = () => {
        toast({
            variant: "destructive",
            title: "File Read Error",
            description: "There was an error reading the file.",
        });
    };

    if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
      reader.readAsArrayBuffer(file);
    } else {
        toast({
            variant: "destructive",
            title: "Unsupported File",
            description: "Please upload an Excel (.xlsx, .xls) file.",
        });
    }
  };
  
  const handleCheckIn = (data: z.infer<typeof checkInSchema>) => {
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
        cellValue => String(cellValue).trim() === searchCode
      )
    );

    if (rowIndex !== -1) {
      const foundRow = rows[rowIndex];
      if (!foundRow.checkedInTime) {
        // First time check-in
        const checkInTime = new Date();
        const updatedRow = { ...foundRow, checkedInTime: checkInTime };
        const updatedRows = [...rows];
        updatedRows[rowIndex] = updatedRow;
        setRows(updatedRows);
        setScannedRow(updatedRow);
        setDialogState('success');
      } else {
        // Already checked in
        setScannedRow(foundRow);
        setDialogState('duplicate');
      }
    } else {
      setScannedRow(undefined); // Not found
      setDialogState('not_found');
    }

    setIsAlertOpen(true);
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
                <CardTitle>Upload Data</CardTitle>
                <CardDescription>Upload an Excel file. The first row should contain headers.</CardDescription>
              </CardHeader>
              <CardContent>
                <Input id="data-upload" type="file" accept=".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={handleDataUpload} />
              </CardContent>
            </Card>
            <Card>
              <CardHeader>
                <CardTitle>Check In Attendee</CardTitle>
                <CardDescription>Enter a unique code from your data to find and check in an attendee.</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="relative mb-4 flex aspect-square w-full items-center justify-center rounded-lg border-2 border-dashed bg-muted animate-border-flash">
                    <QrCode className="h-16 w-16 text-muted-foreground/50"/>
                </div>
                <Form {...form}>
                    <form onSubmit={form.handleSubmit(handleCheckIn)} className="space-y-4">
                        <FormField
                            control={form.control}
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
                                        No data uploaded yet.
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
                  {scannedRow.checkedInTime ? format(scannedRow.checkedInTime, 'PPpp') : 'N/A'}
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
                  {scannedRow.checkedInTime ? format(scannedRow.checkedInTime, 'PPpp') : 'N/A'}
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
            <AlertDialogAction onClick={() => {
              setIsAlertOpen(false);
              form.reset();
            }}>
              Close
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}
