'use client'

import { useEffect, useState } from 'react'
import Link from 'next/link'
import { Card, th, thr, td, tdr } from '@/components/ui/primitives'
import { apiJson, fmtDay, gbp } from '@/components/claims/shared'
import { feeFigures } from '@/lib/fees'
import { giftAidOn, type ClaimDto } from '@/lib/claims'

const STUCK_DAYS = 42

/**
 * The charity's view of BoostAid's fee: one row per claim that includes recovered Gift Aid. The fee is only due once HMRC has
 * paid. Also nudges the charity when a claim has been with HMRC for over 6 weeks: "Has HMRC paid it? Mark it received."
 */
export default function FeeReport({ charityId, refreshKey = 0 }: { charityId: string; refreshKey?: number }) {
  const [claims, setClaims] = useState<ClaimDto[] | null>(null)
  const [now] = useState(() => Date.now())

  useEffect(() => {
    let off = false
    apiJson<{ claims: ClaimDto[] }>(`/api/claims?view=all&pageSize=100&charity=${charityId}`)
      .then((d) => !off && setClaims(d.claims.filter((c) => c.recoveredCount > 0)))
      .catch(() => !off && setClaims([]))
    return () => { off = true }
  }, [charityId, refreshKey])

  if (!claims?.length) return null
  const stuck = claims.filter((c) => c.status === 'submitted' && c.submittedAt && now - Date.parse(c.submittedAt) > STUCK_DAYS * 86400000)
  const rows = claims.map((c) => {
    const rec = giftAidOn(c.recoveredAmount)
    return { c, rec, f: feeFigures(rec, c.giftAidAmount, c.status === 'received' ? c.receivedAmount : null) }
  })
  const due = rows.filter((r) => r.c.status === 'received').reduce((s, r) => s + r.f.fee, 0)

  return (
    <section className="mt-8 space-y-3">
      <div>
        <h2 className="font-display text-xl font-semibold text-ink">Fee report</h2>
        <p className="text-sm text-muted">BoostAid charges 5% of the Gift Aid it recovered, only once HMRC has paid the claim. {due > 0 ? `${gbp(due)} is due on received claims.` : 'Nothing is due yet.'}</p>
      </div>
      {stuck.map((c) => (
        <div key={c.id} className="rounded-xl bg-warn-soft px-4 py-3 text-sm text-warn">
          Has HMRC paid {c.ref}? It was submitted on {fmtDay(c.submittedAt)}. <Link href={`/claims/${c.id}`} className="font-semibold underline">Open it and mark it received</Link> when the money arrives.
        </div>
      ))}
      <Card className="overflow-x-auto">
        <table className="w-full min-w-[680px] text-base">
          <thead>
            <tr className="border-b border-line">
              <th className={th}>Claim</th><th className={thr}>Recovered Gift Aid</th><th className={thr}>Fee (5%)</th><th className={th}>Fee status</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(({ c, rec, f }) => (
              <tr key={c.id} className="border-b border-line-soft last:border-0">
                <td className={td}><Link href={`/claims/${c.id}`} className="font-semibold text-brand hover:underline">{c.ref}</Link></td>
                <td className={`${tdr} text-ink`}>{gbp(c.status === 'received' ? f.recoveredPaid : rec)}{c.status === 'received' && f.recoveredPaid < rec && <div className="text-xs text-muted">of {gbp(rec)}, HMRC paid less</div>}</td>
                <td className={`${tdr} font-semibold text-ink`}>{gbp(f.fee)}</td>
                <td className={td}>{c.status === 'received' ? <span className="font-semibold text-warn">Due</span> : <span className="text-muted">Not due until HMRC pays</span>}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>
    </section>
  )
}
