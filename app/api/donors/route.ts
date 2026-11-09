import { NextResponse } from 'next/server'
import { and, desc, eq, gte, lte, sql, type SQL } from 'drizzle-orm'
import { claims, donations, getDb } from '@/lib/db'
import { withAuth } from '@/lib/api'
import { escapeLike, normalisePostcode } from '@/lib/claims'

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/

/**
 * Donor lookup. Optional filters (with none, the most recent donations are returned):
 *   q         name words - every word must appear in the donor's name
 *   postcode  postcode prefix, spaces/case ignored
 *   from, to  donation date range (yyyy-MM-dd)
 *   amount    exact donation amount
 */
export const GET = withAuth(async (req) => {
  const sp = req.nextUrl.searchParams
  const q = (sp.get('q') ?? '').trim().slice(0, 100)
  const postcode = normalisePostcode(sp.get('postcode') ?? '').slice(0, 10)
  const from = sp.get('from') ?? ''
  const to = sp.get('to') ?? ''
  const amount = sp.get('amount') ?? ''

  if (from && !ISO_DATE.test(from)) throw new RangeError('Invalid "from" date')
  if (to && !ISO_DATE.test(to)) throw new RangeError('Invalid "to" date')
  if (amount && !Number.isFinite(Number(amount))) throw new RangeError('Invalid amount')

  const conditions: SQL[] = []
  // Punctuation is ignored on both sides so "oneill" finds O'Neill and "smithjones" finds Smith-Jones
  const stripPunct = (v: string) => v.replace(/['.\-\u2019]/g, '')
  for (const word of q.split(/\s+/).map(stripPunct).filter(Boolean).slice(0, 6)) {
    conditions.push(
      sql`regexp_replace(${donations.firstName} || ' ' || ${donations.lastName}, '[''.-]', '', 'g') ilike ${'%' + escapeLike(word) + '%'}`
    )
  }
  if (postcode) conditions.push(sql`${donations.postcodeNorm} like ${escapeLike(postcode) + '%'}`)
  if (from) conditions.push(gte(donations.donationDate, from))
  if (to) conditions.push(lte(donations.donationDate, to))
  if (amount) conditions.push(eq(donations.amount, Number(amount).toFixed(2)))

  const limit = 200
  const where = conditions.length ? and(...conditions) : undefined
  const donorKey = sql`lower(${donations.firstName}) || '|' || lower(${donations.lastName}) || '|' || ${donations.postcodeNorm}`
  const [{ totalDonors }] = await getDb()
    .select({ totalDonors: sql<number>`count(distinct ${donorKey})::int` })
    .from(donations)
    .where(where)
  const rows = await getDb()
    .select({
      id: donations.id,
      title: donations.title,
      firstName: donations.firstName,
      lastName: donations.lastName,
      address: donations.address,
      postcode: donations.postcode,
      donationDate: donations.donationDate,
      amount: donations.amount,
      claimId: claims.id,
      claimRef: claims.ref,
      hmrcRef: claims.hmrcRef,
      status: claims.status,
      submittedAt: claims.submittedAt,
      receivedAt: claims.receivedAt,
    })
    .from(donations)
    .innerJoin(claims, eq(claims.id, donations.claimId))
    .where(where)
    .orderBy(desc(donations.donationDate), donations.lastName)
    .limit(limit + 1)

  return NextResponse.json({
    truncated: rows.length > limit,
    totalDonors,
    results: rows.slice(0, limit).map((r) => ({
      ...r,
      amount: Number(r.amount),
      submittedAt: r.submittedAt?.toISOString() ?? null,
      receivedAt: r.receivedAt?.toISOString() ?? null,
    })),
  })
})
