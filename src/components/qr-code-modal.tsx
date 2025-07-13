import * as React from "react";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Copy, Download, Ticket, Loader2 } from "lucide-react";
import { useToast } from "@/hooks/use-toast";

interface QRCodeModalProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  data: string;
}

export function QRCodeModal({ open, onOpenChange, data }: QRCodeModalProps) {
  const { toast } = useToast();
  const [ticketSvg, setTicketSvg] = React.useState<string>('');
  const [isLoading, setIsLoading] = React.useState(false);

  React.useEffect(() => {
    if (data && open) {
      generateTicketPreview();
    }
  }, [data, open]);

  const generateTicketPreview = async () => {
    setIsLoading(true);
    try {
      const response = await fetch('/api/generate-ticket', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ qrData: data }),
      });

      if (response.ok) {
        const svgContent = await response.text();
        setTicketSvg(svgContent);
      }
    } catch (error) {
      console.error('Failed to generate ticket preview:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleCopyData = async () => {
    try {
      await navigator.clipboard.writeText(data);
      toast({
        title: "Copied!",
        description: "QR code data copied to clipboard",
      });
    } catch (error) {
      toast({
        title: "Error",
        description: "Failed to copy to clipboard",
        variant: "destructive",
      });
    }
  };

  const handleDownloadTicket = () => {
    if (ticketSvg) {
      const blob = new Blob([ticketSvg], { type: 'image/svg+xml' });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `ticket-${data.slice(0, 10)}.svg`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
    }
  };


  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-4xl max-h-[90vh]">
        <DialogHeader>
          <DialogTitle>Ticket Preview</DialogTitle>
        </DialogHeader>
        <div className="flex flex-col items-center space-y-4">
          {isLoading ? (
            <div className="flex items-center justify-center h-64">
              <Loader2 className="h-8 w-8 animate-spin" />
              <span className="ml-2">Generating ticket preview...</span>
            </div>
          ) : ticketSvg ? (
            <div className="w-full h-[60vh] overflow-auto border rounded-lg">
              <div 
                className="w-full h-full"
                style={{ width: '100%', height: '100%' }}
                dangerouslySetInnerHTML={{ 
                  __html: ticketSvg.replace(
                    /<svg([^>]*)>/,
                    '<svg$1 width="100%" height="100%" style="max-width: 100%; max-height: 100%;">'
                  )
                }}
              />
            </div>
          ) : (
            <div className="text-center text-muted-foreground">
              Failed to generate ticket preview
            </div>
          )}
          
          <div className="w-full">
            <p className="text-sm text-muted-foreground mb-2">QR Code Data:</p>
            <p className="text-sm font-mono bg-muted p-2 rounded border break-all">
              {data}
            </p>
          </div>

          <div className="flex gap-2 w-full">
            <Button
              variant="outline"
              size="sm"
              onClick={handleCopyData}
              className="flex-1"
            >
              <Copy className="mr-2 h-4 w-4" />
              Copy Data
            </Button>
            <Button
              onClick={handleDownloadTicket}
              className="flex-1"
              disabled={!ticketSvg}
            >
              <Download className="mr-2 h-4 w-4" />
              Download Ticket
            </Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}