import { NextRequest, NextResponse } from 'next/server'
import { SESSION_COOKIE, readSession, verifySessionToken } from '@/lib/session'

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
      if (err instanceof ForbiddenError) {
        return NextResponse.json({ error: err.message }, { status: 403 })
      }
      if (err instanceof RangeError) {
        return NextResponse.json({ error: err.message }, { status: 400 })
      }
      console.error(err)
      return NextResponse.json({ error: 'Something went wrong' }, { status: 500 })
    }
  }
}

export const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i

/** The signed-in user's id (from the session cookie), or null. */
export function currentUserId(req: NextRequest): string | null {
  return readSession(req.cookies.get(SESSION_COOKIE)?.value)?.userId ?? null
}

/**
 * BoostAid's own admin login only. Checked on the server for every admin route (the menu hiding the link is not enough).
 * Throws a 403 for anyone else.
 */
export async function requireAdmin(req: NextRequest): Promise<{ id: string; email: string; name: string }> {
  const { getDb, users } = await import('@/lib/db')
  const { eq } = await import('drizzle-orm')
  const id = currentUserId(req)
  const [u] = id ? await getDb().select().from(users).where(eq(users.id, id)) : []
  if (!u || u.role !== 'admin') throw new ForbiddenError()
  return { id: u.id, email: u.email, name: u.name }
}

export class ForbiddenError extends Error {
  constructor() {
    super('Only the BoostAid admin can do this')
  }
}
