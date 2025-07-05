"use client";

import * as React from "react";
import { useRouter } from "next/navigation";
import { z } from "zod";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { format } from "date-fns";

import type { Ticket } from "@/lib/types";

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
import { Form, FormControl, FormDescription, FormField, FormItem, FormLabel, FormMessage } from "@/components/ui/form";
import {
  TicketCheck,
  LogOut,
  Upload,
  QrCode,
  Download,
  UserCheck,
  XCircle,
  Info,
  Users
} from "lucide-react";

const checkInSchema = z.object({
  uniqueCode: z.string().min(1, { message: "QR code is required." }),
});

export default function DashboardPage() {
  const [isAuthenticated, setIsAuthenticated] = React.useState(false);
  const [tickets, setTickets] = React.useState<Ticket[]>([]);
  const [scannedTicket, setScannedTicket] = React.useState<Ticket | null | undefined>(null); // null: found, undefined: not found
  const [isAlertOpen, setIsAlertOpen] = React.useState(false);
  
  const router = useRouter();
  const { toast } = useToast();

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
    reader.onload = (e) => {
      const text = e.target?.result as string;
      try {
        const parsedTickets = parseCSV(text);
        setTickets(parsedTickets);
        toast({
          title: "Success!",
          description: `Successfully imported ${parsedTickets.length} tickets.`,
        });
      } catch (error) {
        toast({
          variant: "destructive",
          title: "Import Failed",
          description: "Please check the CSV format and try again.",
        });
      }
    };
    reader.readAsText(file);
  };
  
  const parseCSV = (csvText: string): Ticket[] => {
      const lines = csvText.trim().split('\n');
      if (lines.length < 2) return [];
      
      const headers = lines[0].split(',').map(h => h.trim().toLowerCase());
      const requiredHeaders = ['name', 'phone', 'email', 'seatrow', 'seatnumber', 'uniquecode'];
      const missingHeaders = requiredHeaders.filter(h => !headers.includes(h));

      if (missingHeaders.length > 0) {
          throw new Error(`Missing required CSV columns: ${missingHeaders.join(', ')}`);
      }

      return lines.slice(1).map((line) => {
          const values = line.split(',');
          const ticketData: any = {};
          headers.forEach((header, index) => {
              ticketData[header.replace(/\s/g, '')] = values[index]?.trim();
          });

          return {
              name: ticketData.name,
              phone: ticketData.phone,
              email: ticketData.email,
              seat: {
                  row: ticketData.seatrow,
                  number: ticketData.seatnumber,
              },
              uniqueCode: ticketData.uniquecode,
              checkedInTime: null,
          };
      });
  };


  const handleCheckIn = (data: z.infer<typeof checkInSchema>) => {
    const { uniqueCode } = data;
    const foundTicket = tickets.find((t) => t.uniqueCode === uniqueCode);
    setScannedTicket(foundTicket);
    setIsAlertOpen(true);

    if (foundTicket && !foundTicket.checkedInTime) {
      setTickets(
        tickets.map((t) =>
          t.uniqueCode === uniqueCode
            ? { ...t, checkedInTime: new Date() }
            : t
        )
      );
    }
  };
  
  const handleExport = () => {
    if (tickets.length === 0) {
        toast({
            variant: "destructive",
            title: "No Data",
            description: "There is no ticket data to export."
        });
        return;
    }
    const headers = 'Name,Email,Phone,Seat,Unique Code,Checked-In At';
    const rows = tickets.map(t => 
        [
            `"${t.name}"`,
            `"${t.email}"`,
            `"${t.phone}"`,
            `"${t.seat.row}${t.seat.number}"`,
            `"${t.uniqueCode}"`,
            `"${t.checkedInTime ? format(t.checkedInTime, 'yyyy-MM-dd HH:mm:ss') : 'N/A'}"`
        ].join(',')
    );

    const csvContent = [headers, ...rows].join('\n');
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
  
  const form = useForm<z.infer<typeof checkInSchema>>({
    resolver: zodResolver(checkInSchema),
    defaultValues: { uniqueCode: "" },
  });

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
                <CardTitle>Upload Ticket Data</CardTitle>
                <CardDescription>Upload a CSV file with ticket information. Columns: name, phone, email, seatRow, seatNumber, uniqueCode</CardDescription>
              </CardHeader>
              <CardContent>
                <Input id="csv-upload" type="file" accept=".csv" onChange={handleDataUpload} />
              </CardContent>
            </Card>
            <Card>
              <CardHeader>
                <CardTitle>Scan Ticket</CardTitle>
                <CardDescription>Simulate a QR code scan by entering the unique code from the ticket.</CardDescription>
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
                  <CardTitle>Attendee List</CardTitle>
                  <CardDescription>A list of all tickets and their check-in status.</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="max-h-[600px] overflow-auto">
                    <Table>
                        <TableHeader>
                            <TableRow>
                                <TableHead>Name</TableHead>
                                <TableHead>Seat</TableHead>
                                <TableHead>Status</TableHead>
                                <TableHead>Checked In At</TableHead>
                            </TableRow>
                        </TableHeader>
                        <TableBody>
                            {tickets.length > 0 ? (
                                tickets.map(ticket => (
                                    <TableRow key={ticket.uniqueCode}>
                                        <TableCell className="font-medium">{ticket.name}</TableCell>
                                        <TableCell>{ticket.seat.row}{ticket.seat.number}</TableCell>
                                        <TableCell>
                                            <Badge variant={ticket.checkedInTime ? "default" : "secondary"} className={ticket.checkedInTime ? "bg-accent text-accent-foreground" : ""}>
                                                {ticket.checkedInTime ? "Checked In" : "Pending"}
                                            </Badge>
                                        </TableCell>
                                        <TableCell>
                                            {ticket.checkedInTime ? format(ticket.checkedInTime, 'PPpp') : 'N/A'}
                                        </TableCell>
                                    </TableRow>
                                ))
                            ) : (
                                <TableRow>
                                    <TableCell colSpan={4} className="h-24 text-center">
                                        No tickets uploaded yet.
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
          {scannedTicket ? (
            <>
              <AlertDialogHeader>
                {scannedTicket.checkedInTime ? (
                  <>
                    <AlertDialogTitle className="flex items-center gap-2">
                      <UserCheck className="h-6 w-6 text-green-500" />
                      Check-in Successful!
                    </AlertDialogTitle>
                    <AlertDialogDescription>Welcome, {scannedTicket.name}!</AlertDialogDescription>
                  </>
                ) : (
                  <>
                    <AlertDialogTitle className="flex items-center gap-2">
                      <Info className="h-6 w-6 text-blue-500" />
                      Already Checked In
                    </AlertDialogTitle>
                    <AlertDialogDescription>
                      This ticket for {scannedTicket.name} was already used.
                    </AlertDialogDescription>
                  </>
                )}
              </AlertDialogHeader>
              <div className="text-sm">
                <p><strong>Name:</strong> {scannedTicket.name}</p>
                <p><strong>Seat:</strong> {scannedTicket.seat.row}{scannedTicket.seat.number}</p>
                <p><strong>Email:</strong> {scannedTicket.email}</p>
                <p>
                  <strong>Initial Check-in:</strong>{" "}
                  {format(scannedTicket.checkedInTime || new Date(), 'PPpp')}
                </p>
              </div>
            </>
          ) : (
            <>
              <AlertDialogHeader>
                <AlertDialogTitle className="flex items-center gap-2">
                    <XCircle className="h-6 w-6 text-red-500" />
                    Ticket Not Found
                </AlertDialogTitle>
                <AlertDialogDescription>
                    The scanned QR code does not match any ticket in the list. Please try again.
                </AlertDialogDescription>
              </AlertDialogHeader>
            </>
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
