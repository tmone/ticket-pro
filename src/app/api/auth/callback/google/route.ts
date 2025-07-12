import { google } from 'googleapis';
import { NextRequest, NextResponse } from 'next/server';
import { getIronSession } from 'iron-session';
import { cookies } from 'next/headers';
import type { SessionData } from '@/lib/session';
import { sessionOptions } from '@/lib/session';

const OAUTH2_CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const OAUTH2_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const OAUTH2_REDIRECT_URI = process.env.GOOGLE_REDIRECT_URI || 'http://localhost:9002/api/auth/callback/google';

if (!OAUTH2_CLIENT_ID || !OAUTH2_CLIENT_SECRET) {
  throw new Error("Google OAuth2 credentials are not configured in environment variables.");
}

export async function GET(req: NextRequest) {
  const session = await getIronSession<SessionData>(cookies(), sessionOptions);
  
  const code = req.nextUrl.searchParams.get('code');
  if (!code) {
    return NextResponse.json({ error: 'Authorization code not found.' }, { status: 400 });
  }

  try {
    const oauth2Client = new google.auth.OAuth2(
      OAUTH2_CLIENT_ID,
      OAUTH2_CLIENT_SECRET,
      OAUTH2_REDIRECT_URI
    );

    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);

    const oauth2 = google.oauth2({
      auth: oauth2Client,
      version: 'v2',
    });

    const { data: userInfo } = await oauth2.userinfo.get();

    // Save tokens and user info in the session
    session.isLoggedIn = true;
    session.tokens = tokens;
    session.name = userInfo.name || 'User';
    session.email = userInfo.email || '';
    session.picture = userInfo.picture || '';

    await session.save();

    // Redirect to the dashboard
    return NextResponse.redirect(new URL('/', req.url));

  } catch (error) {
    console.error('Error during Google OAuth callback:', error);
    return NextResponse.json({ error: 'Authentication failed.' }, { status: 500 });
  }
}
