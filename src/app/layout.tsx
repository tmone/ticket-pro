import type {Metadata} from 'next';
import { Toaster } from "@/components/ui/toaster"
import { Suspense } from 'react';
import { RedirectHandler } from '@/components/redirect-handler';
import './globals.css';

export const metadata: Metadata = {
  title: 'TicketCheck Pro',
  description: 'Effortless ticket validation and check-in management.',
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" suppressHydrationWarning>
      <head>
        <link rel="preconnect" href="https://fonts.googleapis.com" />
        <link rel="preconnect" href="https://fonts.gstatic.com" crossOrigin="anonymous" />
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet" />
      </head>
      <body className="font-body antialiased" suppressHydrationWarning={true}>
        <Suspense fallback={null}>
          <RedirectHandler />
        </Suspense>
        {children}
        <Toaster />
      </body>
    </html>
  );
}
