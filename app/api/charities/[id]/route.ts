import { NextResponse } from 'next/server'
import { eq, sql } from 'drizzle-orm'
import { charities, getDb } from '@/lib/db'
import { UUID_RE, withAuth } from '@/lib/api'

/**
 * Deletes a charity and everything that belongs to it: uploads, claims, donations and status history
 * (the foreign keys cascade). The response says how much was removed.
 */
export const DELETE = withAuth<{ id: string }>(async (_req, { params }) => {
  const { id } = await params
  if (!UUID_RE.test(id)) return NextResponse.json({ error: 'Not found' }, { status: 404 })

  const result = await getDb().transaction(async (tx) => {
    const [charity] = await tx.select().from(charities).where(eq(charities.id, id)).for('update')
    if (!charity) return null

    const counts = await tx.execute<{ claims: number; donations: number }>(sql`
      select
        (select count(*)::int from claims cl join batches b on b.id = cl.batch_id where b.charity_id = ${id}) as claims,
        (select count(*)::int from donations d join claims cl on cl.id = d.claim_id join batches b on b.id = cl.batch_id
          where b.charity_id = ${id}) as donations
    `)
    await tx.delete(charities).where(eq(charities.id, id))
    return { name: charity.name, claims: counts.rows[0].claims, donations: counts.rows[0].donations }
  })

  if (!result) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  return NextResponse.json({ success: true, deleted: result })
})
