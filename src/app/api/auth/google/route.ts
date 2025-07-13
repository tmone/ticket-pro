import { NextRequest, NextResponse } from 'next/server';

export async function GET(request: NextRequest) {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const redirectUri = process.env.GOOGLE_REDIRECT_URI || 'http://localhost:9002/api/auth/callback/google';
  
  if (!clientId) {
    return NextResponse.json({ error: 'Google Client ID not configured' }, { status: 500 });
  }

  // Google OAuth authorization URL
  const authUrl = new URL('https://accounts.google.com/o/oauth2/v2/auth');
  
  authUrl.searchParams.append('client_id', clientId);
  authUrl.searchParams.append('redirect_uri', redirectUri);
  authUrl.searchParams.append('response_type', 'code');
  authUrl.searchParams.append('scope', 'https://www.googleapis.com/auth/spreadsheets');
  authUrl.searchParams.append('access_type', 'offline');
  authUrl.searchParams.append('prompt', 'consent'); // Force to get refresh token
  
  // Store the original URL to redirect back after auth
  const { searchParams } = new URL(request.url);
  const returnUrl = searchParams.get('returnUrl') || '/';
  authUrl.searchParams.append('state', encodeURIComponent(returnUrl));
  
  return NextResponse.redirect(authUrl.toString());
}