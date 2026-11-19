import { NextResponse } from 'next/server'
import { desc } from 'drizzle-orm'
import { adminNotifications, getDb, invoices } from '@/lib/db'
import { withAuth, requireAdmin } from '@/lib/api'
import { claimsWithRecovery } from '@/lib/admin-db'
import { emailConfigured } from '@/lib/email'

const STUCK_DAYS = 42 // submitted to HMRC more than 6 weeks ago and still not marked received

/** Everything on BoostAid's admin page, across all charities. Admin login only. */
export const GET = withAuth(async (req) => {
  await requireAdmin(req)
  const now = new Date()
  const [claimRows, invs, notes] = await Promise.all([
    claimsWithRecovery(),
    getDb().select().from(invoices),
    getDb().select().from(adminNotifications).orderBy(desc(adminNotifications.createdAt)).limit(40),
  ])
  const invByClaim = new Map(invs.map((i) => [i.claimId, i]))
  const r2 = (n: number) => Math.round(n * 100) / 100

  const rows = claimRows
    .map((c) => {
      const inv = invByClaim.get(c.claimId)
      const fee = inv ? Number(inv.fee) : c.fee
      const stage = inv
        ? inv.status === 'paid' ? 'paid' : inv.dueAt < now ? 'overdue' : 'sent'
        : c.status === 'received' ? 'received' : 'pending'
      const stuck = c.status === 'submitted' && !!c.submittedAt && now.getTime() - c.submittedAt.getTime() > STUCK_DAYS * 86400000
      const stageAt = inv ? (inv.status === 'paid' ? inv.paidAt : inv.sentAt) : c.status === 'received' ? c.receivedAt : c.submittedAt ?? c.createdAt
      return {
        claimId: c.claimId,
        charityId: c.charityId,
        charity: c.charityName,
        ref: c.ref,
        claimStatus: c.status,
        recoveredGiftAid: c.recoveredGiftAid,
        recoveredPaid: inv ? Number(inv.recoveredPaid) : c.recoveredPaid,
        short: c.status === 'received' && c.received != null && c.received < c.expected,
        fee,
        stage,
        stageAt: stageAt?.toISOString() ?? null,
        stuck,
        invoice: inv ? { id: inv.id, number: inv.number, status: inv.status, sentAt: inv.sentAt.toISOString(), dueAt: inv.dueAt.toISOString(), paidAt: inv.paidAt?.toISOString() ?? null, delivered: inv.delivered, sentTo: inv.sentTo } : null,
      }
    })
  const live = rows.filter((r) => r.stage !== 'paid')
  const order = { received: 0, overdue: 1, sent: 2, pending: 3, paid: 4 } as const
  live.sort((a, b) => order[a.stage as keyof typeof order] - order[b.stage as keyof typeof order])

  const sum = (f: (r: (typeof rows)[number]) => boolean) => r2(rows.filter(f).reduce((s, r) => s + r.fee, 0))
  const earnedRows = rows.filter((r) => r.claimStatus === 'received')
  const byCharity = new Map<string, { charity: string; earned: number; owed: number }>()
  for (const r of earnedRows) {
    const e = byCharity.get(r.charityId) ?? { charity: r.charity, earned: 0, owed: 0 }
    e.earned = r2(e.earned + r.fee)
    if (r.stage !== 'paid') e.owed = r2(e.owed + r.fee)
    byCharity.set(r.charityId, e)
  }

  return NextResponse.json({
    emailConfigured: emailConfigured(),
    tiles: {
      earned: sum((r) => r.claimStatus === 'received'),
      owed: sum((r) => r.claimStatus === 'received' && r.stage !== 'paid'),
      overdue: sum((r) => r.stage === 'overdue'),
      waiting: sum((r) => r.claimStatus !== 'received'),
    },
    rows: live,
    readyToInvoice: live.filter((r) => r.stage === 'received').length,
    byCharity: [...byCharity.values()].sort((a, b) => b.earned - a.earned),
    archive: rows
      .filter((r) => r.stage === 'paid' && r.invoice)
      .map((r) => ({ charity: r.charity, number: r.invoice!.number, ref: r.ref, sentAt: r.invoice!.sentAt, paidAt: r.invoice!.paidAt, fee: r.fee }))
      .sort((a, b) => (b.paidAt ?? '').localeCompare(a.paidAt ?? '')),
    notifications: notes.map((n) => ({ id: n.id, type: n.type, text: n.text, at: n.createdAt.toISOString() })),
  })
})
