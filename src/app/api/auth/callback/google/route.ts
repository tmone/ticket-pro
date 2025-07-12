import { google } from 'googleapis';
import { NextRequest, NextResponse } from 'next/server';
import { getIronSession } from 'iron-session';
import { cookies } from 'next/headers';
import type { SessionData } from '@/lib/session';
import { sessionOptions } from '@/lib/session';

const OAUTH2_CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const OAUTH2_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const OAUTH2_REDIRECT_URI = process.env.GOOGLE_REDIRECT_URI || 'http://localhost:9002/api/auth/callback/google';

export async function GET(req: NextRequest) {
  if (!OAUTH2_CLIENT_ID || !OAUTH2_CLIENT_SECRET) {
    const errorMessage = "Google OAuth2 credentials are not configured in environment variables. Please set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET in your .env.local file.";
    console.error(`AUTH_CALLBACK_ERROR: ${errorMessage}`);
    return NextResponse.json({ error: errorMessage }, { status: 500 });
  }

  const session = await getIronSession<SessionData>(cookies(), sessionOptions);
  
  const code = req.nextUrl.searchParams.get('code');
  if (!code) {
    const errorMessage = 'Authorization code not found in callback URL.';
    console.error(`AUTH_CALLBACK_ERROR: ${errorMessage}`);
    return NextResponse.json({ error: errorMessage }, { status: 400 });
  }

  try {
    const oauth2Client = new google.auth.OAuth2(
      OAUTH2_CLIENT_ID,
      OAUTH2_CLIENT_SECRET,
      OAUTH2_REDIRECT_URI
    );

    let tokens;
    try {
      const response = await oauth2Client.getToken(code);
      tokens = response.tokens;
      console.log('Successfully received tokens from Google.');
    } catch (error: any) {
        console.error('AUTH_CALLBACK_ERROR: Error getting token from Google:', error.response?.data || error.message);
        return NextResponse.json({ 
            error: 'Failed to exchange authorization code for tokens.',
            details: error.response?.data || error.message
        }, { status: 400 });
    }
    
    oauth2Client.setCredentials(tokens);

    const oauth2 = google.oauth2({
      auth: oauth2Client,
      version: 'v2',
    });

    let userInfo;
    try {
        const { data } = await oauth2.userinfo.get();
        userInfo = data;
        console.log('Successfully fetched user info from Google.');
    } catch(error: any) {
        console.error('AUTH_CALLBACK_ERROR: Error fetching user info from Google:', error.response?.data || error.message);
        return NextResponse.json({ 
            error: 'Failed to fetch user profile information from Google.',
            details: error.response?.data || error.message
        }, { status: 500 });
    }

    session.isLoggedIn = true;
    session.tokens = tokens;
    session.name = userInfo.name || 'User';
    session.email = userInfo.email || '';
    session.picture = userInfo.picture || '';

    await session.save();
    console.log('Session saved successfully for user:', userInfo.email);

    return NextResponse.redirect(new URL('/', req.nextUrl));

  } catch (error: any) {
    console.error('AUTH_CALLBACK_ERROR: An unexpected error occurred during Google OAuth callback:', error);
    return NextResponse.json({ 
        error: 'An unexpected error occurred during the authentication process.',
        details: error.message
    }, { status: 500 });
  }
}
