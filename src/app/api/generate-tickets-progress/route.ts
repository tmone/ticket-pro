import { NextRequest, NextResponse } from 'next/server';

// Extend global interface
declare global {
  var ticketProgress: Record<string, { current: number; total: number; completed: boolean }> | undefined;
}

export async function GET(request: NextRequest) {
  const sessionId = request.nextUrl.searchParams.get('sessionId');
  
  if (!sessionId) {
    return NextResponse.json({ error: 'Session ID required' }, { status: 400 });
  }

  // Get progress from global store
  const progress = global.ticketProgress?.[sessionId] || { current: 0, total: 0, completed: false };
  
  return NextResponse.json(progress);
}

export async function POST(request: NextRequest) {
  const { sessionId, current, total, completed } = await request.json();
  
  if (!sessionId) {
    return NextResponse.json({ error: 'Session ID required' }, { status: 400 });
  }

  // Store progress in global store
  global.ticketProgress = global.ticketProgress || {};
  global.ticketProgress[sessionId] = { current, total, completed: !!completed };
  
  // Clean up completed sessions after 30 seconds
  if (completed) {
    setTimeout(() => {
      if (global.ticketProgress?.[sessionId]) {
        delete global.ticketProgress[sessionId];
      }
    }, 30000);
  }
  
  return NextResponse.json({ success: true });
}