import { NextResponse } from 'next/server'
import { desc, eq, inArray, sql } from 'drizzle-orm'
import { batches, claims, getDb } from '@/lib/db'
import { withAuth } from '@/lib/api'
import { toClaimDto } from '@/lib/claims'

/** GET /api/claims?view=active|archive|all (active = pending + submitted, archive = received) */
export const GET = withAuth(async (req) => {
  const view = req.nextUrl.searchParams.get('view') ?? 'active'
  const summaryOnly = req.nextUrl.searchParams.get('summary') === '1'
  const db = getDb()

  const where =
    view === 'archive'
      ? eq(claims.status, 'received')
      : view === 'active'
        ? inArray(claims.status, ['pending', 'submitted'])
        : undefined

  const totals = await db
    .select({
      status: claims.status,
      count: sql<number>`count(*)::int`,
      amount: sql<string>`coalesce(sum(${claims.totalAmount}), 0)::text`,
    })
    .from(claims)
    .groupBy(claims.status)

  const summary = { pending: { count: 0, amount: 0 }, submitted: { count: 0, amount: 0 }, received: { count: 0, amount: 0 } }
  for (const t of totals) summary[t.status] = { count: t.count, amount: Number(t.amount) }

  if (summaryOnly) return NextResponse.json({ summary })

  const rows = await db
    .select({ claim: claims, filename: batches.filename })
    .from(claims)
    .innerJoin(batches, eq(batches.id, claims.batchId))
    .where(where)
    .orderBy(desc(claims.createdAt), desc(claims.chunkIndex))
    .limit(1000)

  return NextResponse.json({
    claims: rows.map((r) => toClaimDto(r.claim, r.filename)),
    summary,
  })
})
