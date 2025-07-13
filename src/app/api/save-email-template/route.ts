import { NextRequest, NextResponse } from 'next/server';
import fs from 'fs/promises';
import path from 'path';

export async function POST(request: NextRequest) {
  try {
    const { subject, message, senderEmail, senderName } = await request.json();

    if (!subject || !message) {
      return NextResponse.json(
        { error: 'Subject and message are required' },
        { status: 400 }
      );
    }

    // Create the email template content
    const emailContent = `Subject: ${subject}

${message}`;

    // Define the path to the email template file
    const templatePath = path.join(process.cwd(), 'public', 'templates', 'email.eml');
    
    // Ensure the directory exists
    const templateDir = path.dirname(templatePath);
    await fs.mkdir(templateDir, { recursive: true });
    
    // Write the content to the file
    await fs.writeFile(templatePath, emailContent, 'utf-8');
    
    console.log('Email template saved:', {
      path: templatePath,
      subject,
      messageLength: message.length,
      senderEmail,
      senderName
    });

    return NextResponse.json({
      success: true,
      message: 'Email template saved successfully'
    });

  } catch (error) {
    console.error('Save email template error:', error);
    return NextResponse.json(
      { 
        error: 'Failed to save email template', 
        details: error instanceof Error ? error.message : String(error) 
      },
      { status: 500 }
    );
  }
}