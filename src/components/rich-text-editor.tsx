"use client";

import React from 'react';
import { Button } from '@/components/ui/button';
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu';
import { ChevronDown } from 'lucide-react';
import { Textarea } from '@/components/ui/textarea';

interface RichTextEditorProps {
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  headers?: string[];
}

export function RichTextEditor({ value, onChange, placeholder, headers = [] }: RichTextEditorProps) {
  const textareaRef = React.useRef<HTMLTextAreaElement>(null);

  const insertPlaceholder = (placeholder: string) => {
    const textarea = textareaRef.current;
    if (!textarea) return;

    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const text = textarea.value;
    const before = text.substring(0, start);
    const after = text.substring(end);
    
    const newValue = `${before}{${placeholder}}${after}`;
    onChange(newValue);
    
    // Set cursor position after the inserted placeholder
    setTimeout(() => {
      textarea.focus();
      const newPosition = start + placeholder.length + 2;
      textarea.setSelectionRange(newPosition, newPosition);
    }, 0);
  };

  const systemPlaceholders = [
    { label: 'Sender Name', value: 'senderName' },
    { label: 'Sender Email', value: 'senderEmail' },
    { label: 'Ticket Code', value: 'ticketCode' },
    { label: 'Event Date', value: 'eventDate' },
  ];

  const insertHtmlTag = (tag: string, attributes?: string) => {
    const textarea = textareaRef.current;
    if (!textarea) return;

    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const text = textarea.value;
    const selectedText = text.substring(start, end);
    const before = text.substring(0, start);
    const after = text.substring(end);
    
    let newValue = '';
    if (tag === 'br') {
      newValue = `${before}<br>${after}`;
    } else if (selectedText) {
      newValue = `${before}<${tag}${attributes ? ' ' + attributes : ''}>${selectedText}</${tag}>${after}`;
    } else {
      newValue = `${before}<${tag}${attributes ? ' ' + attributes : ''}></${tag}>${after}`;
    }
    
    onChange(newValue);
    
    // Set cursor position
    setTimeout(() => {
      textarea.focus();
      if (tag === 'br') {
        const newPosition = start + 4;
        textarea.setSelectionRange(newPosition, newPosition);
      } else if (!selectedText) {
        const newPosition = start + tag.length + (attributes ? attributes.length + 1 : 0) + 2;
        textarea.setSelectionRange(newPosition, newPosition);
      }
    }, 0);
  };

  return (
    <div className="space-y-2">
      <div className="flex gap-2 mb-2 flex-wrap">
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button variant="outline" size="sm">
              Insert Column <ChevronDown className="ml-1 h-4 w-4" />
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent>
            {headers.map((header) => (
              <DropdownMenuItem
                key={header}
                onClick={() => insertPlaceholder(header)}
              >
                {header}
              </DropdownMenuItem>
            ))}
          </DropdownMenuContent>
        </DropdownMenu>

        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button variant="outline" size="sm">
              System Fields <ChevronDown className="ml-1 h-4 w-4" />
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent>
            {systemPlaceholders.map((item) => (
              <DropdownMenuItem
                key={item.value}
                onClick={() => insertPlaceholder(item.value)}
              >
                {item.label}
              </DropdownMenuItem>
            ))}
          </DropdownMenuContent>
        </DropdownMenu>
        
        <div className="flex gap-1">
          <Button
            variant="outline"
            size="sm"
            onClick={() => insertHtmlTag('b')}
            title="Bold"
          >
            <b>B</b>
          </Button>
          <Button
            variant="outline"
            size="sm"
            onClick={() => insertHtmlTag('i')}
            title="Italic"
          >
            <i>I</i>
          </Button>
          <Button
            variant="outline"
            size="sm"
            onClick={() => insertHtmlTag('u')}
            title="Underline"
          >
            <u>U</u>
          </Button>
          <Button
            variant="outline"
            size="sm"
            onClick={() => insertHtmlTag('br')}
            title="Line Break"
          >
            ↵
          </Button>
          <Button
            variant="outline"
            size="sm"
            onClick={() => insertHtmlTag('a', 'href=""')}
            title="Link"
          >
            🔗
          </Button>
        </div>
      </div>
      
      <div className="text-xs text-muted-foreground">
        You can use HTML tags for formatting. Example: &lt;b&gt;bold&lt;/b&gt;, &lt;i&gt;italic&lt;/i&gt;, &lt;br&gt; for line break
      </div>

      <Textarea
        ref={textareaRef}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className="min-h-[300px] font-mono text-sm"
      />
      
      <div className="text-xs text-muted-foreground space-y-1">
        <div>💡 Placeholders: {`{columnName}`} will be replaced with actual values from your spreadsheet</div>
        <div>📝 HTML Preview: The email will be sent as HTML with your formatting preserved</div>
      </div>
    </div>
  );
}