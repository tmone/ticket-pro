import { google } from 'googleapis';
import { NextRequest, NextResponse } from 'next/server';

const OAUTH2_CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const OAUTH2_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const OAUTH2_REDIRECT_URI = process.env.GOOGLE_REDIRECT_URI || 'http://localhost:9002/api/auth/callback/google';

export async function GET(req: NextRequest) {
  if (!OAUTH2_CLIENT_ID || !OAUTH2_CLIENT_SECRET) {
    return NextResponse.json({ error: "Google OAuth2 credentials are not configured in environment variables. Please set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET in your .env.local file." }, { status: 500 });
  }

  const oauth2Client = new google.auth.OAuth2(
    OAUTH2_CLIENT_ID,
    OAUTH2_CLIENT_SECRET,
    OAUTH2_REDIRECT_URI
  );

  // Request full access to read/write sheets and user profile info
  const scopes = [
    'https://www.googleapis.com/auth/userinfo.profile',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/spreadsheets' // Changed from .readonly to full access
  ];

  const authorizationUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline', // 'offline' gets a refresh token, important for long-lived sessions
    scope: scopes,
    include_granted_scopes: true,
  });

  return NextResponse.redirect(authorizationUrl);
}
