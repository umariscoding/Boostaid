import { Card } from '@/components/ui/primitives'
import { gbp } from '@/components/claims/shared'
import { feeFigures } from '@/lib/fees'
import { giftAidOn, type ClaimDto } from '@/lib/claims'

/**
 * For a claim that includes Gift Aid BoostAid recovered by emailing donors: what was ready, what was recovered, and the 5% fee.
 * The fee is only due once HMRC has paid and the claim is marked received, and it follows what HMRC actually paid.
 */
export default function RecoverySummary({ claim }: { claim: ClaimDto }) {
  if (claim.recoveredCount <= 0) return null
  const recovered = giftAidOn(claim.recoveredAmount)
  const ready = Math.round((claim.giftAidAmount - recovered) * 100) / 100
  const received = claim.status === 'received' ? claim.receivedAmount : null
  const f = feeFigures(recovered, claim.giftAidAmount, received)
  const short = received != null && f.recoveredPaid < recovered
  const readyCount = claim.donationCount - claim.recoveredCount

  return (
    <Card className="p-6">
      <h3 className="font-display text-xl font-semibold text-ink">Gift Aid recovered by BoostAid</h3>
      <dl className="mt-4 grid max-w-3xl grid-cols-[1fr_auto] gap-x-8 gap-y-2 text-base text-ink">
        <dt>Gift Aid on donations that were ready ({readyCount.toLocaleString()})</dt><dd className="text-right font-semibold tabular-nums">{gbp(ready)}</dd>
        <dt>Gift Aid recovered by BoostAid ({claim.recoveredCount.toLocaleString()})</dt><dd className="text-right font-semibold tabular-nums">{gbp(recovered)}</dd>
        <dt>Total Gift Aid in this claim</dt><dd className="text-right font-semibold tabular-nums">{gbp(claim.giftAidAmount)}</dd>
        {short && (
          <>
            <dt>Recovered Gift Aid HMRC actually paid (it paid {gbp(claim.giftAidAmount - (received ?? 0))} less than claimed)</dt>
            <dd className="text-right font-semibold tabular-nums">{gbp(f.recoveredPaid)}</dd>
          </>
        )}
        <dt className="border-t border-line pt-2 font-semibold">Extra recovered by BoostAid: {gbp(f.recoveredPaid)} Gift Aid. Fee at 5%</dt>
        <dd className="border-t border-line pt-2 text-right font-bold tabular-nums">{gbp(f.fee)}</dd>
      </dl>
      <div className={`mt-4 rounded-xl px-4 py-3 text-sm ${received != null ? 'bg-warn-soft text-warn' : 'bg-info-soft text-info'}`}>
        {received != null
          ? `HMRC paid ${gbp(received)}. Without BoostAid you would have had about ${gbp(f.withoutBoostAid ?? 0)}. The ${gbp(f.fee)} fee is now due.`
          : `The fee is only due once HMRC pays and the claim is marked received. There is no fee on the ${gbp(ready)} of Gift Aid on donations that were ready.`}
      </div>
    </Card>
  )
}
