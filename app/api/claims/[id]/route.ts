import { NextResponse } from 'next/server'
import { asc, eq, sql } from 'drizzle-orm'
import { batches, claimEvents, claims, donations, getDb } from '@/lib/db'
import type { ClaimStatus } from '@/lib/db'
import { UUID_RE, withAuth } from '@/lib/api'
import { toClaimDto } from '@/lib/claims'

type Params = { id: string }

export const GET = withAuth<Params>(async (_req, { params }) => {
  const { id } = await params
  if (!UUID_RE.test(id)) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const db = getDb()

  const [row] = await db
    .select({ claim: claims, filename: batches.filename })
    .from(claims)
    .innerJoin(batches, eq(batches.id, claims.batchId))
    .where(eq(claims.id, id))
  if (!row) return NextResponse.json({ error: 'Not found' }, { status: 404 })

  const [rows, events] = await Promise.all([
    db.select().from(donations).where(eq(donations.claimId, id)).orderBy(asc(donations.lastName), asc(donations.firstName)),
    db.select().from(claimEvents).where(eq(claimEvents.claimId, id)).orderBy(asc(claimEvents.at)),
  ])

  return NextResponse.json({
    claim: toClaimDto(row.claim, row.filename),
    donations: rows.map((d) => ({
      id: d.id,
      title: d.title,
      firstName: d.firstName,
      lastName: d.lastName,
      address: d.address,
      postcode: d.postcode,
      donationDate: d.donationDate,
      amount: Number(d.amount),
    })),
    events: events.map((e) => ({ id: e.id, status: e.status, at: e.at.toISOString(), note: e.note })),
  })
})

const STATUSES: ClaimStatus[] = ['pending', 'submitted', 'received']

/** Update status, HMRC reference and/or notes. Status changes are recorded in the history. */
export const PATCH = withAuth<Params>(async (req, { params }) => {
  const { id } = await params
  if (!UUID_RE.test(id)) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const body = await req.json().catch(() => null)
  if (!body || typeof body !== 'object') return NextResponse.json({ error: 'Invalid body' }, { status: 400 })

  const db = getDb()
  const updated = await db.transaction(async (tx) => {
    const [current] = await tx.select().from(claims).where(eq(claims.id, id)).for('update')
    if (!current) return null

    const patch: Partial<typeof claims.$inferInsert> = {}
    if ('hmrcRef' in body) patch.hmrcRef = body.hmrcRef ? String(body.hmrcRef).trim().slice(0, 100) : null
    if ('notes' in body) patch.notes = body.notes ? String(body.notes).trim().slice(0, 2000) : null

    if ('status' in body) {
      const status = body.status as ClaimStatus
      if (!STATUSES.includes(status)) throw new RangeError('Invalid status')
      if (status !== current.status) {
        const now = new Date()
        patch.status = status
        if (status === 'pending') {
          patch.submittedAt = null
          patch.receivedAt = null
        } else if (status === 'submitted') {
          patch.submittedAt = current.submittedAt ?? now
          patch.receivedAt = null
        } else {
          patch.submittedAt = current.submittedAt ?? now
          patch.receivedAt = now
        }
        await tx.insert(claimEvents).values({
          claimId: id,
          status,
          at: now,
          note: typeof body.note === 'string' ? body.note.trim().slice(0, 500) || null : null,
        })
      }
    }

    if (Object.keys(patch).length === 0) return current
    const [next] = await tx.update(claims).set(patch).where(eq(claims.id, id)).returning()
    return next
  })

  if (!updated) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  return NextResponse.json({ claim: toClaimDto(updated) })
})

/** Only claims that are still pending can be deleted (e.g. a test upload). */
export const DELETE = withAuth<Params>(async (_req, { params }) => {
  const { id } = await params
  if (!UUID_RE.test(id)) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const db = getDb()

  const outcome = await db.transaction(async (tx) => {
    const [current] = await tx.select().from(claims).where(eq(claims.id, id)).for('update')
    if (!current) return 'missing' as const
    if (current.status !== 'pending') return 'locked' as const
    await tx.delete(claims).where(eq(claims.id, id))
    // Remove the batch too once it has no claims left
    await tx.delete(batches).where(
      sql`${batches.id} = ${current.batchId} and not exists (select 1 from ${claims} where ${claims.batchId} = ${current.batchId})`
    )
    return 'deleted' as const
  })

  if (outcome === 'missing') return NextResponse.json({ error: 'Not found' }, { status: 404 })
  if (outcome === 'locked') {
    return NextResponse.json({ error: 'Only pending claims can be deleted' }, { status: 409 })
  }
  return NextResponse.json({ success: true })
})
