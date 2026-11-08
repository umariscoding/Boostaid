import { NextRequest, NextResponse } from 'next/server'
import { SESSION_COOKIE, verifySessionToken } from '@/lib/session'

type Ctx<P> = { params: Promise<P> }

/**
 * Wraps a route handler with a session check and a catch-all error handler.
 * The proxy already guards these routes; this is defence in depth.
 */
export function withAuth<P = Record<string, never>>(
  handler: (req: NextRequest, ctx: Ctx<P>) => Promise<Response>
) {
  return async (req: NextRequest, ctx: Ctx<P>) => {
    if (!verifySessionToken(req.cookies.get(SESSION_COOKIE)?.value)) {
      return NextResponse.json({ error: 'Unauthorised' }, { status: 401 })
    }
    try {
      return await handler(req, ctx)
    } catch (err) {
      if (err instanceof RangeError) {
        return NextResponse.json({ error: err.message }, { status: 400 })
      }
      console.error(err)
      return NextResponse.json({ error: 'Something went wrong' }, { status: 500 })
    }
  }
}

export const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
