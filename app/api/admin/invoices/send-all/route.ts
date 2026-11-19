import { NextResponse } from 'next/server'
import { withAuth, requireAdmin } from '@/lib/api'
import { claimsWithRecovery } from '@/lib/admin-db'
import { getDb, invoices } from '@/lib/db'
import { sendInvoice } from '@/lib/invoices'

/** Sends every invoice that is ready (received, not yet invoiced). Each one stands alone: one failure never blocks the rest. */
export const POST = withAuth(async (req) => {
  await requireAdmin(req)
  const [all, done] = await Promise.all([claimsWithRecovery(), getDb().select({ claimId: invoices.claimId }).from(invoices)])
  const invoiced = new Set(done.map((d) => d.claimId))
  const ready = all.filter((c) => c.status === 'received' && !invoiced.has(c.claimId))
  const sent: string[] = []
  const failed: { ref: string; charity: string; error: string }[] = []
  for (const c of ready) {
    try {
      await sendInvoice(c.claimId)
      sent.push(c.ref)
    } catch (e) {
      failed.push({ ref: c.ref, charity: c.charityName, error: (e as Error).message })
    }
  }
  return NextResponse.json({ sent: sent.length, failed })
})
