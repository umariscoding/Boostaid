import { NextResponse } from 'next/server'
import { sql } from 'drizzle-orm'
import { batches, charities, claimEvents, claims, donations, getDb } from '@/lib/db'
import { UUID_RE, withAuth } from '@/lib/api'
import { isoDateOrNull, normalisePostcode } from '@/lib/claims'
import { findAlreadyClaimed } from '@/lib/duplicates'

interface IncomingDonation {
  title?: unknown
  firstName?: unknown
  lastName?: unknown
  address?: unknown
  postcode?: unknown
  donationDate?: unknown
  amount?: unknown
  platform?: unknown
  donationId?: unknown
  hmrcRow?: unknown
}

const str = (v: unknown, max = 500) => String(v ?? '').slice(0, max).trim()
const count = (v: unknown) => (Number.isFinite(Number(v)) ? Math.max(0, Math.trunc(Number(v))) : 0)

/**
 * Saves a cleaned claim: one batch for the charity, one claim per HMRC schedule file, and every donation in it.
 * A donation that is already claimed for this charity (same platform + donation ID) is refused outright.
 */
export const POST = withAuth(async (req) => {
  const body = await req.json().catch(() => null)
  if (!body || !Array.isArray(body.claims) || body.claims.length === 0 || body.claims.length > 500) {
    return NextResponse.json({ error: 'No claims supplied' }, { status: 400 })
  }

  const charityId = String(body.charityId ?? '')
  if (!UUID_RE.test(charityId)) throw new RangeError('Choose which charity this upload is for')
  const [charity] = await getDb().select().from(charities).where(sql`${charities.id} = ${charityId}`)
  if (!charity) return NextResponse.json({ error: 'That charity no longer exists' }, { status: 404 })

  type Prepared = { kind: string; fileName: string; rows: Omit<typeof donations.$inferInsert, 'claimId'>[]; total: number }
  const prepared: Prepared[] = []
  for (const c of body.claims as { kind?: unknown; fileName?: unknown; donations?: IncomingDonation[] }[]) {
    if (!Array.isArray(c.donations) || c.donations.length === 0 || c.donations.length > 1000) {
      return NextResponse.json({ error: 'Each claim needs 1-1000 donations' }, { status: 400 })
    }
    const kind = c.kind === 'Overseas' ? 'Overseas' : 'UK'
    let total = 0
    const rows = c.donations.map((d) => {
      const amount = Number(d.amount)
      if (!Number.isFinite(amount) || amount < 0) throw new RangeError('Invalid donation amount')
      total += Math.round(amount * 100)
      const postcode = str(d.postcode, 20)
      const id = str(d.donationId, 200)
      return {
        title: str(d.title, 20),
        firstName: str(d.firstName, 100),
        lastName: str(d.lastName, 100),
        address: str(d.address, 300),
        postcode,
        postcodeNorm: normalisePostcode(postcode),
        donationDate: isoDateOrNull(d.donationDate),
        amount: amount.toFixed(2),
        platform: str(d.platform, 100),
        platformDonationId: id || null,
        hmrcRow: Number.isFinite(Number(d.hmrcRow)) ? Math.trunc(Number(d.hmrcRow)) : null,
      }
    })
    prepared.push({ kind, fileName: str(c.fileName, 200), rows, total: total / 100 })
  }

  const db = getDb()
  const stats = body.stats ?? {}

  const result = await db.transaction(async (tx) => {
    // Stop the same platform donation being claimed twice (the Clean screen checks first; this is the safety net)
    const all = prepared.flatMap((p) => p.rows.map((r) => ({ platform: r.platform ?? '', donationId: r.platformDonationId ?? '' })))
    const already = await findAlreadyClaimed(tx as unknown as typeof db, charityId, all)
    if (already.length) return { blocked: already }

    const [batch] = await tx
      .insert(batches)
      .values({
        charityId,
        filename: str(body.filename, 255) || 'upload',
        totalRecords: count(stats.totalRecords),
        eligibleRecords: count(stats.eligibleRecords),
        validRecords: count(stats.validRecords),
        invalidRecords: count(stats.invalidRecords),
      })
      .returning()

    const year = new Date().getFullYear()
    const created: { id: string; ref: string; kind: string; fileName: string; chunkIndex: number; donationCount: number; totalAmount: number }[] = []

    for (let i = 0; i < prepared.length; i++) {
      const { rows, total, kind, fileName } = prepared[i]
      const seq = await tx.execute<{ n: string }>(sql`select nextval('claim_ref_seq')::text as n`)
      const ref = `GA-${year}-${seq.rows[0].n.padStart(4, '0')}`

      const [claim] = await tx
        .insert(claims)
        .values({ ref, batchId: batch.id, chunkIndex: i + 1, kind, fileName: fileName || null, donationCount: rows.length, totalAmount: total.toFixed(2) })
        .returning()

      await tx.insert(claimEvents).values({ claimId: claim.id, status: 'pending', note: 'Claim created' })
      await tx.insert(donations).values(rows.map((r) => ({ ...r, claimId: claim.id })))
      created.push({ id: claim.id, ref, kind, fileName, chunkIndex: i + 1, donationCount: rows.length, totalAmount: total })
    }

    // Donations with no platform ID (e.g. a cash book) can't be matched exactly. Warn if one looks identical
    // (same donor, postcode, date and amount) to a donation already in an earlier claim for this charity.
    const similar = await tx.execute<{ ref: string; n: string }>(sql`
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
      join ${batches} ob on ob.id = oc.batch_id and ob.charity_id = ${charityId}
      where d.platform_donation_id is null
      group by oc.ref
      order by oc.ref
    `)

    return {
      batchId: batch.id,
      charity: { id: charity.id, name: charity.name },
      claims: created,
      possibleRepeats: { count: similar.rows.reduce((s, r) => s + Number(r.n), 0), claimRefs: similar.rows.map((r) => r.ref) },
    }
  })

  if ('blocked' in result) {
    return NextResponse.json(
      { error: 'Some of these donations are already claimed for this charity', claimed: result.blocked },
      { status: 409 }
    )
  }
  return NextResponse.json(result, { status: 201 })
})
