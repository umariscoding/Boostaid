import { NextRequest, NextResponse } from 'next/server'
import { and, eq, inArray, or } from 'drizzle-orm'
import { charities, getDb, recoveryLog, recoveryRequests, reviewRows } from '@/lib/db'
import { DECLARATION_TEXT, checkAnswers } from '@/lib/recovery'
import { hashCode } from '@/lib/recovery-codes'
import { giftAidOn } from '@/lib/claims'

/**
 * The donor's form, reached from the email. No login: the secret code in the link is the key. Each link works once and
 * expires after 60 days. Nothing about the donor is in the link, and nothing is shown until the code is valid.
 */
type Found =
  | { state: 'invalid' }
  | { state: 'expired' | 'used'; charity: string; contactEmail: string | null }
  | { state: 'done'; charity: string; contactEmail: string | null }
  | { state: 'open'; charityId: string; charity: string; contactEmail: string | null; rootId: string; pending: (typeof reviewRows.$inferSelect)[] }

async function find(code: string): Promise<Found> {
  if (!/^[A-Za-z0-9]{6,32}$/.test(code)) return { state: 'invalid' }
  const db = getDb()
  const [req] = await db.select().from(recoveryRequests).where(eq(recoveryRequests.codeHash, hashCode(code)))
  if (!req) return { state: 'invalid' }
  const rootId = req.parentId ?? req.id
  const [charity] = await db.select().from(charities).where(eq(charities.id, req.charityId))
  const base = { charity: charity?.name ?? 'the charity', contactEmail: charity?.contactEmail ?? null }
  const [root] = req.parentId ? await db.select().from(recoveryRequests).where(eq(recoveryRequests.id, rootId)) : [req]
  if (req.usedAt || root?.usedAt) return { state: 'used', ...base }
  if (req.expiresAt <= new Date()) return { state: 'expired', ...base }
  const pending = await db.select().from(reviewRows).where(and(eq(reviewRows.requestId, rootId), inArray(reviewRows.status, ['sent', 'reminded'])))
  if (!pending.length) return { state: 'done', ...base }
  return { state: 'open', charityId: req.charityId, rootId, pending, ...base }
}

export async function GET(_req: NextRequest, { params }: { params: Promise<{ code: string }> }) {
  const { code } = await params
  const f = await find(code)
  if (f.state === 'invalid') return NextResponse.json({ state: 'invalid' }, { status: 404 })
  if (f.state !== 'open') return NextResponse.json({ state: f.state, charity: f.charity, contactEmail: f.contactEmail })
  return NextResponse.json({
    state: 'open',
    charity: f.charity,
    declaration: DECLARATION_TEXT.replace('{charity}', f.charity),
    twoPeople: f.pending.some((r) => /two people|joint/i.test(r.reason)),
    donations: f.pending.map((r) => ({ date: r.donationDate, amount: Number(r.amount), giftAid: giftAidOn(Number(r.amount)) })),
  })
}

export async function POST(req: NextRequest, { params }: { params: Promise<{ code: string }> }) {
  const { code } = await params
  const body = await req.json().catch(() => null)
  const f = await find(code)
  if (f.state === 'invalid') return NextResponse.json({ error: 'This link is not valid.' }, { status: 404 })
  if (f.state !== 'open') return NextResponse.json({ error: 'This link has already been used or has expired.', state: f.state }, { status: 410 })
  if (!body || typeof body !== 'object') return NextResponse.json({ error: 'Invalid request.' }, { status: 400 })
  if (body.declaration !== true) return NextResponse.json({ error: 'Tick the Gift Aid declaration.' }, { status: 400 })

  const s = (v: unknown) => (typeof v === 'string' ? v.slice(0, 120) : '')
  const answers = { title: s(body.title), first: s(body.first), last: s(body.last), house: s(body.house), postcode: s(body.postcode) }
  // Every donation goes through the cleaning engine's normal checks before it becomes "Ready – recovered"
  const checks = f.pending.map((r) => checkAnswers(answers, r.donationDate ?? '2026-01-01', Number(r.amount)))
  const bad = checks.find((c) => !c.ok)
  if (bad && !bad.ok) return NextResponse.json({ error: bad.error }, { status: 400 })
  const cleaned = checks[0].ok ? checks[0].cleaned : answers

  const ip = (req.headers.get('x-forwarded-for') ?? '').split(',')[0].trim().slice(0, 64)
  const now = new Date()
  const db = getDb()
  await db.transaction(async (tx) => {
    // Only rows still waiting for this donor are changed (a row fixed another way in the meantime is left alone)
    const changed = await tx
      .update(reviewRows)
      .set({
        status: 'recovered',
        fixedBy: 'donor',
        repliedAt: now,
        fixedTitle: cleaned.title,
        fixedFirst: cleaned.first,
        fixedLast: cleaned.last,
        fixedHouse: cleaned.house,
        fixedPostcode: cleaned.postcode,
      })
      .where(and(inArray(reviewRows.id, f.pending.map((r) => r.id)), inArray(reviewRows.status, ['sent', 'reminded'])))
      .returning({ id: reviewRows.id })
    await tx.update(recoveryRequests).set({ usedAt: now }).where(or(eq(recoveryRequests.id, f.rootId), eq(recoveryRequests.parentId, f.rootId)))
    await tx.insert(recoveryLog).values({
      charityId: f.charityId,
      requestId: f.rootId,
      type: 'form_submitted',
      detail: JSON.stringify({ rows: changed.length, alsoCoverOtherDonations: body.all === true, declaration: DECLARATION_TEXT.replace('{charity}', f.charity), ip }),
    })
  })
  return NextResponse.json({ ok: true, first: cleaned.first, charity: f.charity })
}
