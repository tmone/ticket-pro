"use client";

import * as React from "react";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter, AlertDialogHeader, AlertDialogTitle } from "@/components/ui/alert-dialog";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import { Send, Loader2, TestTube, Save } from "lucide-react";
import { useToast } from "@/hooks/use-toast";
import { RichTextEditor } from "@/components/rich-text-editor";

interface EmailModalProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onSuccess?: () => void;
  selectedEmails: { email: string; name?: string; qrData: string; rowNumber: string; originalRowIndex?: number; rowData?: Record<string, any> }[];
  spreadsheetId?: string;
  sheetName?: string;
  emailColumn?: string;
  headers?: string[];
}

export function EmailModal({ open, onOpenChange, onSuccess, selectedEmails, spreadsheetId, sheetName, emailColumn, headers }: EmailModalProps) {
  const { toast } = useToast();
  const [subject, setSubject] = React.useState("");
  const [message, setMessage] = React.useState("");
  const [senderEmail, setSenderEmail] = React.useState("");
  const [senderName, setSenderName] = React.useState("");
  const [bccList, setBccList] = React.useState<string[]>([]);
  const [attachTicket, setAttachTicket] = React.useState(true);
  const [appendTicketInline, setAppendTicketInline] = React.useState(false);
  const [isSending, setIsSending] = React.useState(false);
  const [sendingProgress, setSendingProgress] = React.useState({ current: 0, total: 0 });
  const [isTesting, setIsTesting] = React.useState(false);
  const [showResendWarning, setShowResendWarning] = React.useState(false);
  const [previouslySentEmails, setPreviouslySentEmails] = React.useState<string[]>([]);
  const [isSaving, setIsSaving] = React.useState(false);
  
  // Handle dialog close
  const handleClose = React.useCallback((newOpen: boolean) => {
    if (!newOpen) {
      // Reset any pending states when closing
      setShowResendWarning(false);
      setIsSending(false);
      setIsTesting(false);
      setIsSaving(false);
    }
    onOpenChange(newOpen);
  }, [onOpenChange]);

  // Load email template and sender info from server on mount
  React.useEffect(() => {
    const loadEmailTemplate = async () => {
      try {
        const response = await fetch('/api/get-email-template');
        if (response.ok) {
          const data = await response.json();
          setSubject(data.subject || 'Your Event Ticket');
          setMessage(data.message || '');
          setSenderName(data.senderName || '');
          setSenderEmail(data.senderEmail || '');
          setBccList(data.bccList || []);
        }
      } catch (error) {
        console.error('Failed to load email template:', error);
        // Set fallback values
        setSubject('Your Event Ticket');
        setMessage(`<p>Dear {name},</p>
<p>Thank you for registering for our event. Please find your ticket attached.</p>
<p>Best regards,<br>Event Team</p>`);
      }
    };
    
    if (open) {
      loadEmailTemplate();
    }
  }, [open]);

  // Check for previously sent emails
  React.useEffect(() => {
    if (open && selectedEmails.length > 0 && spreadsheetId && sheetName) {
      checkPreviouslySentEmails();
    }
  }, [open, selectedEmails, spreadsheetId, sheetName]);

  const checkPreviouslySentEmails = async () => {
    try {
      const response = await fetch('/api/check-sent-emails', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          spreadsheetId,
          sheetName,
          emailAddresses: selectedEmails.map(e => e.email),
          emailColumn
        }),
      });

      if (response.ok) {
        const data = await response.json();
        setPreviouslySentEmails(data.previouslySentEmails || []);
      }
    } catch (error) {
      console.error('Failed to check previously sent emails:', error);
    }
  };

  const handleSendEmails = async (forceResend = false) => {
    if (!senderEmail || !subject || !message) {
      toast({
        title: "Error",
        description: "Please fill in all required fields",
        variant: "destructive",
      });
      return;
    }

    // Check for resend without force
    if (!forceResend && previouslySentEmails.length > 0) {
      setShowResendWarning(true);
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
          appendTicketInline,
          spreadsheetId,
          sheetName,
          emailColumnIndex: headers?.findIndex(h => h === emailColumn) || -1,
          emailSentColumnName: 'Email Sent'
        }),
      });

      if (response.ok) {
        const data = await response.json();
        const updateInfo = data.updatedSuccessRows || data.updatedErrorRows 
          ? ` (Updated ${data.updatedSuccessRows || 0} success + ${data.updatedErrorRows || 0} error rows in Google Sheets)`
          : '';
        toast({
          title: "Success!",
          description: `Sent ${data.successCount} emails successfully${data.failureCount > 0 ? `, ${data.failureCount} failed` : ''}${updateInfo}`,
        });
        // Call success callback if provided
        if (onSuccess) {
          onSuccess();
        }
        handleClose(false);
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
      setShowResendWarning(false);
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

  const handleSaveTemplate = async () => {
    setIsSaving(true);

    try {
      const response = await fetch('/api/save-email-template', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          subject,
          message,
          senderEmail,
          senderName,
        }),
      });

      if (response.ok) {
        toast({
          title: "Template Saved!",
          description: "Email template has been saved successfully.",
        });
      } else {
        const error = await response.json();
        throw new Error(error.error || 'Failed to save template');
      }
    } catch (error) {
      console.error('Save template error:', error);
      toast({
        title: "Save Failed",
        description: error instanceof Error ? error.message : "Failed to save template",
        variant: "destructive",
      });
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <>
      <Dialog open={open} onOpenChange={handleClose}>
        <DialogContent className="sm:max-w-2xl max-h-[90vh] overflow-auto">
          <DialogHeader>
            <DialogTitle>Send Emails</DialogTitle>
            <div className="text-sm text-muted-foreground mt-1">
              {selectedEmails.length} recipient{selectedEmails.length !== 1 ? 's' : ''} selected
            </div>
          </DialogHeader>
        
        <div className="space-y-4">
          
          <div className="grid grid-cols-2 gap-4">
            <div>
              <Label htmlFor="sender-name">Sender Name</Label>
              <Input
                id="sender-name"
                value={senderName}
                placeholder="Your Name"
                readOnly
                className="bg-muted"
              />
            </div>
            <div>
              <Label htmlFor="sender-email">Sender Email *</Label>
              <div className="flex gap-2">
                <Input
                  id="sender-email"
                  type="email"
                  value={senderEmail}
                  placeholder="your@email.com"
                  required
                  readOnly
                  className="flex-1 bg-muted"
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
            <RichTextEditor
              value={message}
              onChange={setMessage}
              placeholder="Email message body"
              headers={headers}
            />
          </div>
          
          <div className="space-y-2">
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
            
            <div className="flex items-center space-x-2">
              <Checkbox
                id="append-ticket-inline"
                checked={appendTicketInline}
                onCheckedChange={(checked) => setAppendTicketInline(!!checked)}
              />
              <Label htmlFor="append-ticket-inline" className="text-sm">
                Append ticket as JPG image at bottom letter
              </Label>
            </div>
          </div>
          
          
          {previouslySentEmails.length > 0 && (
            <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-3">
              <div className="text-sm font-medium text-yellow-800 mb-1">
                ⚠️ Some emails have been sent before
              </div>
              <div className="text-xs text-yellow-700">
                {previouslySentEmails.length} recipient(s) have already received emails. 
                Continuing will resend to all selected recipients.
              </div>
            </div>
          )}
          

          <div className="flex gap-2 justify-between">
            <Button
              variant="outline"
              onClick={handleSaveTemplate}
              disabled={isSending || isSaving}
            >
              {isSaving ? (
                <>
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  Saving...
                </>
              ) : (
                <>
                  <Save className="mr-2 h-4 w-4" />
                  Save Template
                </>
              )}
            </Button>
            
            <div className="flex gap-2">
              <Button
                variant="outline"
                onClick={() => handleClose(false)}
                disabled={isSending}
              >
                Cancel
              </Button>
              <Button
                onClick={() => handleSendEmails(false)}
                disabled={isSending}
              >
                {isSending ? (
                  <>
                    <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                    {sendingProgress.current > 0 
                      ? `Sending... (${sendingProgress.current}/${sendingProgress.total})`
                      : 'Sending...'
                    }
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
        </div>
        </DialogContent>
      </Dialog>

      <AlertDialog open={showResendWarning} onOpenChange={setShowResendWarning}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Resend Confirmation</AlertDialogTitle>
            <AlertDialogDescription>
              {previouslySentEmails.length} of the selected recipients have already received emails:
              <div className="mt-2 max-h-32 overflow-auto bg-muted p-2 rounded">
                {previouslySentEmails.map((email, index) => (
                  <div key={index} className="text-xs">{email}</div>
                ))}
              </div>
              <div className="mt-2">
                Are you sure you want to resend emails to all selected recipients? This may be considered spam.
              </div>
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel disabled={isSending}>Cancel</AlertDialogCancel>
            <AlertDialogAction 
              onClick={() => handleSendEmails(true)}
              disabled={isSending}
              className="bg-red-600 hover:bg-red-700"
            >
              {isSending ? (
                <>
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  Sending...
                </>
              ) : (
                'Yes, Resend All'
              )}
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </>
  );
}