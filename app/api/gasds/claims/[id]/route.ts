import { NextResponse } from 'next/server'
import { asc, eq } from 'drizzle-orm'
import { charities, gasdsClaims, gasdsReceipts, getDb } from '@/lib/db'
import { UUID_RE, withAuth } from '@/lib/api'
import { toClaimDto, toReceipt } from '@/lib/gasds-db'

/** The evidence for one claim, for an auditor: the claim, the Gift Aid figure, and every receipt in it with its counts. */
export const GET = withAuth<{ id: string }>(async (_req, { params }) => {
  const { id } = await params
  if (!UUID_RE.test(id)) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const db = getDb()
  const [row] = await db
    .select({ claim: gasdsClaims, charity: charities.name })
    .from(gasdsClaims)
    .innerJoin(charities, eq(charities.id, gasdsClaims.charityId))
    .where(eq(gasdsClaims.id, id))
  if (!row) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const receipts = await db.select().from(gasdsReceipts).where(eq(gasdsReceipts.claimId, id)).orderBy(asc(gasdsReceipts.receiptDate), asc(gasdsReceipts.receiptNo))
  return NextResponse.json({
    claim: toClaimDto(row.claim),
    charity: row.charity,
    receipts: receipts.map((r) => toReceipt(r, row.claim.hmrcRef)),
  })
})

/** Mark received: the amount HMRC actually paid. Once received the claim is locked. */
export const PATCH = withAuth<{ id: string }>(async (req, { params }) => {
  const { id } = await params
  if (!UUID_RE.test(id)) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const body = await req.json().catch(() => null)
  if (!body || typeof body !== 'object') throw new RangeError('Invalid body')
  const paid = Number(body.receivedAmount)
  if (body.status !== 'received') throw new RangeError('Only marking a claim received is allowed')
  if (body.receivedAmount === undefined || body.receivedAmount === '' || !Number.isFinite(paid) || paid <= 0 || paid > 1e9) {
    throw new RangeError('Enter the amount HMRC actually paid')
  }
  const db = getDb()
  const updated = await db.transaction(async (tx) => {
    const [cur] = await tx.select().from(gasdsClaims).where(eq(gasdsClaims.id, id)).for('update')
    if (!cur) return null
    if (cur.status === 'received') throw new RangeError('This claim is marked received, so it is locked for audit and cannot be changed')
    const [next] = await tx
      .update(gasdsClaims)
      .set({ status: 'received', receivedAmount: (Math.round(paid * 100) / 100).toFixed(2), receivedAt: new Date() })
      .where(eq(gasdsClaims.id, id))
      .returning()
    return next
  })
  if (!updated) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  return NextResponse.json({ claim: toClaimDto(updated) })
})
