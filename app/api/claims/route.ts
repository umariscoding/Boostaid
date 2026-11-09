import { NextResponse } from 'next/server'
import { and, desc, eq, inArray, sql } from 'drizzle-orm'
import { batches, charities, claims, donations, getDb } from '@/lib/db'
import { withAuth } from '@/lib/api'
import { toClaimDto } from '@/lib/claims'
import { claimsOfCharity, parseCharityParam } from '@/lib/scope'

/** GET /api/claims?view=active|archive|all (active = pending + submitted, archive = received) */
export const GET = withAuth(async (req) => {
  const view = req.nextUrl.searchParams.get('view') ?? 'active'
  const summaryOnly = req.nextUrl.searchParams.get('summary') === '1'
  const page = Math.max(1, Math.trunc(Number(req.nextUrl.searchParams.get('page'))) || 1)
  const pageSize = Math.min(100, Math.max(1, Math.trunc(Number(req.nextUrl.searchParams.get('pageSize'))) || 25))
  const charityId = parseCharityParam(req.nextUrl.searchParams.get('charity'))
  const db = getDb()

  const statusFilter =
    view === 'archive'
      ? eq(claims.status, 'received')
      : view === 'active'
        ? inArray(claims.status, ['pending', 'submitted'])
        : undefined
  const scope = charityId ? claimsOfCharity(charityId) : undefined
  const where = and(statusFilter, scope)

  const totals = await db
    .select({
      status: claims.status,
      count: sql<number>`count(*)::int`,
      amount: sql<string>`coalesce(sum(${claims.totalAmount}), 0)::text`,
    })
    .from(claims)
    .where(scope)
    .groupBy(claims.status)

  const summary = { pending: { count: 0, amount: 0 }, submitted: { count: 0, amount: 0 }, received: { count: 0, amount: 0 } }
  for (const t of totals) summary[t.status] = { count: t.count, amount: Number(t.amount) }

  if (summaryOnly) return NextResponse.json({ summary })

  const [{ total }] = await db.select({ total: sql<number>`count(*)::int` }).from(claims).where(where)

  const rows = await db
    .select({
      claim: claims,
      filename: batches.filename,
      charityId: charities.id,
      charityName: charities.name,
      periodFrom: sql<string | null>`(select min(${donations.donationDate})::text from ${donations} where ${donations.claimId} = ${claims.id})`,
      periodTo: sql<string | null>`(select max(${donations.donationDate})::text from ${donations} where ${donations.claimId} = ${claims.id})`,
    })
    .from(claims)
    .innerJoin(batches, eq(batches.id, claims.batchId))
    .innerJoin(charities, eq(charities.id, batches.charityId))
    .where(where)
    .orderBy(desc(claims.createdAt), desc(claims.chunkIndex))
    .limit(pageSize)
    .offset((page - 1) * pageSize)

  return NextResponse.json({
    claims: rows.map((r) => toClaimDto(r.claim, r.filename, { from: r.periodFrom, to: r.periodTo }, { id: r.charityId, name: r.charityName })),
    summary,
    total,
    page,
    pageSize,
  })
})
