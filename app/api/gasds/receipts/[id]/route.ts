import { NextResponse } from 'next/server'
import { eq } from 'drizzle-orm'
import { gasdsReceipts, getDb } from '@/lib/db'
import { UUID_RE, withAuth } from '@/lib/api'
import { countTotal, type Count } from '@/lib/gasds'
import { toReceipt } from '@/lib/gasds-db'

const whole = (v: unknown) => Math.max(0, Math.min(100000, Math.trunc(Number(v) || 0)))

/** Step 2: the notes and coins counted for a collection, and/or its bank slip number. Not allowed once claimed. */
export const PATCH = withAuth<{ id: string }>(async (req, { params }) => {
  const { id } = await params
  if (!UUID_RE.test(id)) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const body = await req.json().catch(() => null)
  if (!body || typeof body !== 'object') throw new RangeError('Invalid body')
  const db = getDb()
  const updated = await db.transaction(async (tx) => {
    const [cur] = await tx.select().from(gasdsReceipts).where(eq(gasdsReceipts.id, id)).for('update')
    if (!cur) return null
    if (cur.claimId) throw new RangeError('This receipt is in a claim sent to HMRC, so it is locked')
    const patch: Partial<typeof gasdsReceipts.$inferInsert> = {}
    if ('count' in body) {
      if (body.count === null) patch.count = null
      else {
        const c: Count = {
          n50: whole(body.count?.n50),
          n20: whole(body.count?.n20),
          n10: whole(body.count?.n10),
          n5: whole(body.count?.n5),
          coins: Math.max(0, Math.min(1e7, Math.round((Number(body.count?.coins) || 0) * 100) / 100)),
        }
        patch.count = countTotal(c) > 0 ? c : null
      }
    }
    if ('slip' in body) patch.slip = typeof body.slip === 'string' && body.slip.trim() ? body.slip.trim().slice(0, 60) : null
    if (!Object.keys(patch).length) return cur
    const [next] = await tx.update(gasdsReceipts).set(patch).where(eq(gasdsReceipts.id, id)).returning()
    return next
  })
  if (!updated) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  return NextResponse.json({ receipt: toReceipt(updated) })
})
