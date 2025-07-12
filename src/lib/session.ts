import type { IronSessionOptions } from 'iron-session';
import type { TokenSet } from 'google-auth-library';

// This is the shape of our session data
export interface SessionData {
  isLoggedIn: boolean;
  tokens?: TokenSet;
  name?: string;
  email?: string;
  picture?: string;
}

// Ensure you have IRON_SESSION_PASSWORD set in your .env.local file
// It should be a secret string of at least 32 characters
const password = process.env.IRON_SESSION_PASSWORD;
if (!password) {
  throw new Error('IRON_SESSION_PASSWORD environment variable not set.');
}

export const sessionOptions: IronSessionOptions = {
  cookieName: 'ticketcheck_pro_session',
  password: password,
  cookieOptions: {
    // secure: true should be used in production (HTTPS)
    secure: process.env.NODE_ENV === 'production',
    // httpOnly: true, // Prevents client-side JS from accessing the cookie
  },
};
