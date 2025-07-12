import { google } from 'googleapis';
import { NextRequest, NextResponse } from 'next/server';

const OAUTH2_CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const OAUTH2_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const OAUTH2_REDIRECT_URI = process.env.GOOGLE_REDIRECT_URI || 'http://localhost:9002/api/auth/callback/google';

if (!OAUTH2_CLIENT_ID || !OAUTH2_CLIENT_SECRET) {
  throw new Error("Google OAuth2 credentials are not configured in environment variables.");
}

export async function GET(req: NextRequest) {
  const oauth2Client = new google.auth.OAuth2(
    OAUTH2_CLIENT_ID,
    OAUTH2_CLIENT_SECRET,
    OAUTH2_REDIRECT_URI
  );

  const scopes = [
    'https://www.googleapis.com/auth/userinfo.profile',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    // Add 'https://www.googleapis.com/auth/spreadsheets' for write access later
  ];

  const authorizationUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline', // 'offline' gets a refresh token
    scope: scopes,
    include_granted_scopes: true,
  });

  return NextResponse.redirect(authorizationUrl);
}
