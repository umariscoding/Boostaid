import { NextResponse } from 'next/server'
import { and, desc, eq, gte, inArray, lte, sql, type SQL } from 'drizzle-orm'
import { claims, donations, getDb } from '@/lib/db'
import { withAuth } from '@/lib/api'
import { escapeLike, normalisePostcode } from '@/lib/claims'

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/

/**
 * Donor lookup, one row per donor, paged. Optional filters (with none, the most recent donors are listed):
 *   q         name words - every word must appear in the donor's name
 *   postcode  postcode prefix, spaces/case ignored
 *   from, to  donation date range (yyyy-MM-dd)
 *   amount    exact donation amount
 *   page, pageSize  1-based page and size (max 100)
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

  const page = Math.max(1, Math.trunc(Number(sp.get('page'))) || 1)
  const pageSize = Math.min(100, Math.max(1, Math.trunc(Number(sp.get('pageSize'))) || 25))

  const db = getDb()
  const where = conditions.length ? and(...conditions) : undefined
  // A donor is the same first name + last name + postcode
  const donorKey = sql<string>`lower(${donations.firstName}) || '|' || lower(${donations.lastName}) || '|' || ${donations.postcodeNorm}`

  const [{ total }] = await db
    .select({ total: sql<number>`count(distinct ${donorKey})::int` })
    .from(donations)
    .where(where)

  // One page of donors, most recently donated first
  const pageDonors = await db
    .select({ key: donorKey })
    .from(donations)
    .where(where)
    .groupBy(donorKey)
    .orderBy(sql`max(${donations.donationDate}) desc nulls last`, donorKey)
    .limit(pageSize)
    .offset((page - 1) * pageSize)
  const keys = pageDonors.map((d) => d.key)

  // Every matching donation for the donors on this page
  const rows = keys.length
    ? await db
        .select({
          key: donorKey,
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
        .where(and(where, inArray(donorKey, keys)))
        .orderBy(desc(donations.donationDate))
    : []

  const byKey = new Map<string, typeof rows>()
  for (const r of rows) byKey.set(r.key, [...(byKey.get(r.key) ?? []), r])

  const donors = keys.map((key) => {
    const list = byKey.get(key) ?? []
    const first = list[0]
    const claimMap = new Map<string, { id: string; ref: string; status: string }>()
    for (const r of list) claimMap.set(r.claimId, { id: r.claimId, ref: r.claimRef, status: r.status })
    return {
      key,
      name: first ? [first.title, first.firstName, first.lastName].filter(Boolean).join(' ') : '',
      house: first?.address ?? '',
      postcode: first?.postcode ?? '',
      donationCount: list.length,
      total: Math.round(list.reduce((sum, r) => sum + Number(r.amount), 0) * 100) / 100,
      claims: [...claimMap.values()],
      donations: list.map((r) => ({
        id: r.id,
        donationDate: r.donationDate,
        amount: Number(r.amount),
        claimId: r.claimId,
        claimRef: r.claimRef,
        hmrcRef: r.hmrcRef,
        status: r.status,
        submittedAt: r.submittedAt?.toISOString() ?? null,
        receivedAt: r.receivedAt?.toISOString() ?? null,
      })),
    }
  })

  return NextResponse.json({ donors, total, page, pageSize })
})
