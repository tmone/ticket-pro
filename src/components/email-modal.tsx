"use client";

import * as React from "react";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import { Send, Loader2, TestTube } from "lucide-react";
import { useToast } from "@/hooks/use-toast";

interface EmailModalProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  selectedEmails: { email: string; name?: string; qrData: string; rowNumber: string }[];
}

export function EmailModal({ open, onOpenChange, selectedEmails }: EmailModalProps) {
  const { toast } = useToast();
  const [subject, setSubject] = React.useState("Your Event Ticket");
  const [message, setMessage] = React.useState(`Dear {name},

Thank you for registering for our event. Please find your ticket attached.

Best regards,
Event Team`);
  const [senderEmail, setSenderEmail] = React.useState("");
  const [senderName, setSenderName] = React.useState("");
  const [bccList, setBccList] = React.useState<string[]>([]);

  // Load sender info from environment on mount
  React.useEffect(() => {
    const loadSenderInfo = async () => {
      try {
        const response = await fetch('/api/get-sender-info');
        if (response.ok) {
          const data = await response.json();
          setSenderName(data.senderName);
          setSenderEmail(data.senderEmail);
          setBccList(data.bccList || []);
        }
      } catch (error) {
        console.error('Failed to load sender info:', error);
      }
    };
    
    if (open) {
      loadSenderInfo();
    }
  }, [open]);
  const [attachTicket, setAttachTicket] = React.useState(true);
  const [isSending, setIsSending] = React.useState(false);
  const [isTesting, setIsTesting] = React.useState(false);

  const handleSendEmails = async () => {
    if (!senderEmail || !subject || !message) {
      toast({
        title: "Error",
        description: "Please fill in all required fields",
        variant: "destructive",
      });
      return;
    }

    setIsSending(true);

    try {
      const response = await fetch('/api/send-emails', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          emails: selectedEmails,
          subject,
          message,
          senderEmail,
          senderName,
          attachTicket,
        }),
      });

      if (response.ok) {
        toast({
          title: "Success!",
          description: `Emails sent to ${selectedEmails.length} recipients`,
        });
        onOpenChange(false);
      } else {
        const error = await response.json();
        throw new Error(error.error || 'Failed to send emails');
      }
    } catch (error) {
      console.error('Send emails error:', error);
      toast({
        title: "Error",
        description: error instanceof Error ? error.message : "Failed to send emails",
        variant: "destructive",
      });
    } finally {
      setIsSending(false);
    }
  };

  const handleTestEmail = async () => {
    if (!senderEmail) {
      toast({
        title: "Error",
        description: "Please enter sender email first",
        variant: "destructive",
      });
      return;
    }

    setIsTesting(true);

    try {
      const response = await fetch('/api/test-email', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ testEmail: senderEmail }),
      });

      const result = await response.json();

      if (response.ok) {
        const bccInfo = result.config?.bccList?.length > 0 
          ? ` (BCC: ${result.config.bccList.join(', ')})`
          : '';
        toast({
          title: "Test Email Sent!",
          description: `Test email sent to ${senderEmail}${bccInfo}. Check your inbox.`,
        });
      } else {
        throw new Error(result.error || 'Failed to send test email');
      }
    } catch (error) {
      console.error('Test email error:', error);
      toast({
        title: "Test Failed",
        description: error instanceof Error ? error.message : "Failed to send test email",
        variant: "destructive",
      });
    } finally {
      setIsTesting(false);
    }
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-2xl max-h-[90vh] overflow-auto">
        <DialogHeader>
          <DialogTitle>Send Emails to Selected Recipients</DialogTitle>
        </DialogHeader>
        
        <div className="space-y-4">
          <div className="text-sm text-muted-foreground">
            Sending to {selectedEmails.length} recipients
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            <div>
              <Label htmlFor="sender-name">Sender Name</Label>
              <Input
                id="sender-name"
                value={senderName}
                onChange={(e) => setSenderName(e.target.value)}
                placeholder="Your Name"
              />
            </div>
            <div>
              <Label htmlFor="sender-email">Sender Email *</Label>
              <div className="flex gap-2">
                <Input
                  id="sender-email"
                  type="email"
                  value={senderEmail}
                  onChange={(e) => setSenderEmail(e.target.value)}
                  placeholder="your@email.com"
                  required
                  className="flex-1"
                />
                <Button
                  type="button"
                  variant="outline"
                  size="sm"
                  onClick={handleTestEmail}
                  disabled={isTesting || !senderEmail}
                >
                  {isTesting ? (
                    <Loader2 className="h-4 w-4 animate-spin" />
                  ) : (
                    <TestTube className="h-4 w-4" />
                  )}
                </Button>
              </div>
              <div className="text-xs text-muted-foreground mt-1">
                Click test button to verify email configuration
              </div>
            </div>
          </div>
          
          <div>
            <Label htmlFor="subject">Subject *</Label>
            <Input
              id="subject"
              value={subject}
              onChange={(e) => setSubject(e.target.value)}
              placeholder="Email subject"
              required
            />
          </div>
          
          <div>
            <Label htmlFor="message">Message *</Label>
            <Textarea
              id="message"
              value={message}
              onChange={(e) => setMessage(e.target.value)}
              placeholder="Email message body"
              rows={8}
              required
              className="resize-none"
            />
            <div className="text-xs text-muted-foreground mt-1">
              Use {"{name}"} to insert recipient name
            </div>
          </div>
          
          <div className="flex items-center space-x-2">
            <Checkbox
              id="attach-ticket"
              checked={attachTicket}
              onCheckedChange={(checked) => setAttachTicket(!!checked)}
            />
            <Label htmlFor="attach-ticket" className="text-sm">
              Attach ticket as JPG image
            </Label>
          </div>
          
          <div className="bg-muted p-3 rounded-lg">
            <div className="text-sm font-medium mb-2">Preview Recipients:</div>
            <div className="max-h-32 overflow-auto space-y-1">
              {selectedEmails.slice(0, 5).map((recipient, index) => (
                <div key={index} className="text-xs">
                  {recipient.name ? `${recipient.name} (${recipient.email})` : recipient.email}
                </div>
              ))}
              {selectedEmails.length > 5 && (
                <div className="text-xs text-muted-foreground">
                  ... and {selectedEmails.length - 5} more
                </div>
              )}
              {bccList.length > 0 && (
                <div className="text-xs text-muted-foreground mt-2 pt-2 border-t">
                  <strong>BCC:</strong> {bccList.join(', ')}
                </div>
              )}
            </div>
          </div>
          
          <div className="flex gap-2 justify-end">
            <Button
              variant="outline"
              onClick={() => onOpenChange(false)}
              disabled={isSending}
            >
              Cancel
            </Button>
            <Button
              onClick={handleSendEmails}
              disabled={isSending}
            >
              {isSending ? (
                <>
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  Sending...
                </>
              ) : (
                <>
                  <Send className="mr-2 h-4 w-4" />
                  Send Emails
                </>
              )}
            </Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}