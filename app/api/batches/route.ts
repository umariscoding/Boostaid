import { NextResponse } from 'next/server'
import { sql } from 'drizzle-orm'
import { batches, claimEvents, claims, donations, getDb } from '@/lib/db'
import { withAuth } from '@/lib/api'
import { normalisePostcode, ukDateToIso } from '@/lib/claims'

interface IncomingDonation {
  title?: unknown
  firstName?: unknown
  lastName?: unknown
  address?: unknown
  postcode?: unknown
  donationDate?: unknown
  amount?: unknown
}

const str = (v: unknown, max = 500) => String(v ?? '').slice(0, max).trim()
const count = (v: unknown) => (Number.isFinite(Number(v)) ? Math.max(0, Math.trunc(Number(v))) : 0)

/** Saves a processed upload: one batch, one claim per HMRC file, and every donation in it. */
export const POST = withAuth(async (req) => {
  const body = await req.json().catch(() => null)
  if (!body || !Array.isArray(body.claims) || body.claims.length === 0 || body.claims.length > 500) {
    return NextResponse.json({ error: 'No claims supplied' }, { status: 400 })
  }

  const prepared: { rows: (typeof donations.$inferInsert)[]; total: number }[] = []
  for (const c of body.claims as { donations?: IncomingDonation[] }[]) {
    if (!Array.isArray(c.donations) || c.donations.length === 0 || c.donations.length > 1000) {
      return NextResponse.json({ error: 'Each claim needs 1-1000 donations' }, { status: 400 })
    }
    let total = 0
    const rows = c.donations.map((d) => {
      const amount = Number(d.amount)
      if (!Number.isFinite(amount) || amount < 0) throw new RangeError('Invalid donation amount')
      total += Math.round(amount * 100)
      const postcode = str(d.postcode, 20)
      return {
        title: str(d.title, 20),
        firstName: str(d.firstName, 100),
        lastName: str(d.lastName, 100),
        address: str(d.address, 300),
        postcode,
        postcodeNorm: normalisePostcode(postcode),
        donationDate: ukDateToIso(d.donationDate),
        amount: amount.toFixed(2),
      } satisfies Omit<typeof donations.$inferInsert, 'claimId'>
    })
    prepared.push({ rows: rows as (typeof donations.$inferInsert)[], total: total / 100 })
  }

  const db = getDb()
  const stats = body.stats ?? {}

  const result = await db.transaction(async (tx) => {
    const [batch] = await tx
      .insert(batches)
      .values({
        filename: str(body.filename, 255) || 'upload.xlsx',
        totalRecords: count(stats.totalRecords),
        eligibleRecords: count(stats.eligibleRecords),
        validRecords: count(stats.validRecords),
        invalidRecords: count(stats.invalidRecords),
      })
      .returning()

    const year = new Date().getFullYear()
    const created: { id: string; ref: string; chunkIndex: number; donationCount: number; totalAmount: number }[] = []

    for (let i = 0; i < prepared.length; i++) {
      const { rows, total } = prepared[i]
      const seq = await tx.execute<{ n: string }>(sql`select nextval('claim_ref_seq')::text as n`)
      const ref = `GA-${year}-${seq.rows[0].n.padStart(4, '0')}`

      const [claim] = await tx
        .insert(claims)
        .values({
          ref,
          batchId: batch.id,
          chunkIndex: i + 1,
          donationCount: rows.length,
          totalAmount: total.toFixed(2),
        })
        .returning()

      await tx.insert(claimEvents).values({ claimId: claim.id, status: 'pending', note: 'Claim created' })
      await tx.insert(donations).values(rows.map((r) => ({ ...r, claimId: claim.id })))
      created.push({ id: claim.id, ref, chunkIndex: i + 1, donationCount: rows.length, totalAmount: total })
    }

    // Donations that already appear in an earlier batch (same donor, postcode, date and amount)
    const dupes = await tx.execute<{ ref: string; n: string }>(sql`
      select oc.ref as ref, count(*)::text as n
      from ${donations} d
      join ${claims} nc on nc.id = d.claim_id and nc.batch_id = ${batch.id}
      join ${donations} o
        on o.postcode_norm = d.postcode_norm
       and o.amount = d.amount
       and o.donation_date is not distinct from d.donation_date
       and lower(o.first_name) = lower(d.first_name)
       and lower(o.last_name) = lower(d.last_name)
      join ${claims} oc on oc.id = o.claim_id and oc.batch_id <> ${batch.id}
      group by oc.ref
      order by oc.ref
    `)

    return {
      batchId: batch.id,
      claims: created,
      duplicates: {
        count: dupes.rows.reduce((s, r) => s + Number(r.n), 0),
        claimRefs: dupes.rows.map((r) => r.ref),
      },
    }
  })

  return NextResponse.json(result, { status: 201 })
})

