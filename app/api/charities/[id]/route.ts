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

    // HMRC expects Gift Aid records to be kept for 6 years: a charity with claims sent to HMRC can't be deleted
    const sent = await tx.execute<{ n: number }>(sql`
      select count(*)::int as n from claims cl join batches b on b.id = cl.batch_id
      where b.charity_id = ${id} and cl.status in ('submitted', 'received')
    `)
    if (sent.rows[0].n > 0) return { locked: sent.rows[0].n } as const
    const gasdsSent = await tx.execute<{ n: number }>(sql`select count(*)::int as n from gasds_claims where charity_id = ${id}`)
    if (gasdsSent.rows[0].n > 0) return { locked: gasdsSent.rows[0].n } as const

    const counts = await tx.execute<{ claims: number; donations: number }>(sql`
      select
        (select count(*)::int from claims cl join batches b on b.id = cl.batch_id where b.charity_id = ${id}) as claims,
        (select count(*)::int from donations d join claims cl on cl.id = d.claim_id join batches b on b.id = cl.batch_id
          where b.charity_id = ${id}) as donations
    `)
    // Receipts that were never claimed are only records: they go with the charity
    await tx.execute(sql`delete from gasds_receipts where charity_id = ${id}`)
    await tx.delete(charities).where(eq(charities.id, id))
    return { name: charity.name, claims: counts.rows[0].claims, donations: counts.rows[0].donations }
  })

  if (!result) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  if ('locked' in result) {
    return NextResponse.json(
      { error: `This charity has ${result.locked} claim${result.locked === 1 ? '' : 's'} sent to HMRC (Gift Aid or small donations). Gift Aid records must be kept for 6 years, so it can't be deleted.` },
      { status: 409 }
    )
  }
  return NextResponse.json({ success: true, deleted: result })
})
