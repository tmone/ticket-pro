import { NextResponse } from 'next/server';

export async function GET() {
  const bccList = process.env.SENDER_BCC 
    ? process.env.SENDER_BCC.replace(/"/g, '').split(';').filter(email => email.trim())
    : [];
    
  return NextResponse.json({
    senderName: process.env.SENDER_NAME || '',
    senderEmail: process.env.SENDER_EMAIL || process.env.EMAIL_USER || '',
    bccList: bccList
  });
}