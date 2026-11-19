import { NextRequest, NextResponse } from 'next/server'
import { timingSafeEqual } from 'node:crypto'
import { appBase } from '@/lib/email'
import { processDue } from '@/lib/recovery-db'

/**
 * Runs daily (Vercel Cron): one reminder after 14 days, "No response" 14 days later. Protected by CRON_SECRET, which
 * Vercel sends as a Bearer token. With no CRON_SECRET set nobody can call it. `?asOf=2026-12-01` runs it as of that
 * date (for testing; still needs the secret).
 */
export async function GET(req: NextRequest) {
  const secret = process.env.CRON_SECRET
  const given = (req.headers.get('authorization') ?? '').replace(/^Bearer\s+/i, '')
  const ok = !!secret && given.length === secret.length && timingSafeEqual(Buffer.from(given), Buffer.from(secret))
  if (!ok) return NextResponse.json({ error: 'Unauthorised' }, { status: 401 })
  const asOf = req.nextUrl.searchParams.get('asOf')
  const now = asOf && /^\d{4}-\d{2}-\d{2}$/.test(asOf) ? new Date(asOf + 'T12:00:00Z') : new Date()
  try {
    return NextResponse.json({ now: now.toISOString(), ...(await processDue(appBase(req.nextUrl.origin), now)) })
  } catch (e) {
    console.error(e)
    return NextResponse.json({ error: 'The daily job failed' }, { status: 500 })
  }
}
