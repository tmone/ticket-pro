import { NextRequest, NextResponse } from 'next/server';
import { cookies } from 'next/headers';

// Get password from environment variables
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'password123';

export async function POST(request: NextRequest) {
  try {
    const { password } = await request.json();

    // Validate password
    if (password === ADMIN_PASSWORD) {
      // Create a simple token (in production, use JWT or proper session management)
      const token = Buffer.from(`admin:${Date.now()}`).toString('base64');
      
      // Set cookie
      cookies().set('auth-token', token, {
        httpOnly: true,
        secure: process.env.NODE_ENV === 'production',
        sameSite: 'lax',
        maxAge: 60 * 60 * 24 * 7, // 7 days
        path: '/',
      });

      return NextResponse.json({ success: true });
    } else {
      return NextResponse.json(
        { error: 'Invalid password' },
        { status: 401 }
      );
    }
  } catch (error) {
    console.error('Login error:', error);
    return NextResponse.json(
      { error: 'Internal server error' },
      { status: 500 }
    );
  }
}