'use client'

import { useEffect, useState } from 'react'
import { Loader2 } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Dialog } from '@/components/ui/dialog'
import { inputClass, labelClass } from '@/components/ui/primitives'
import { apiJson, gbp } from '@/components/claims/shared'
import type { ClaimDto } from '@/lib/claims'

/** "Mark received" asks what HMRC actually paid, and shows the difference from what was expected. */
export default function ReceivedDialog({
  claim,
  onClose,
  onDone,
}: {
  claim: { id: string; ref: string; giftAidAmount: number } | null
  onClose: () => void
  onDone: (updated: ClaimDto) => void
}) {
  const [amount, setAmount] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    if (claim) {
      setAmount(claim.giftAidAmount.toFixed(2))
      setError(null)
    }
  }, [claim])

  const paid = amount.trim() === '' ? NaN : Number(amount)
  const valid = Number.isFinite(paid) && paid >= 0
  const diff = claim && valid ? Math.round((paid - claim.giftAidAmount) * 100) / 100 : 0

  const submit = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!claim || !valid) return
    setBusy(true)
    setError(null)
    try {
      const { claim: updated } = await apiJson<{ claim: ClaimDto }>(`/api/claims/${claim.id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status: 'received', receivedAmount: paid }),
      })
      onDone(updated)
      onClose()
    } catch (err) {
      setError((err as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <Dialog open={!!claim} onClose={() => !busy && onClose()}>
      {claim && (
        <form onSubmit={submit}>
          <h2 className="font-display text-xl font-semibold text-ink">Mark {claim.ref} as received</h2>
          <p className="mt-2 text-muted">Type the amount HMRC actually paid into the bank. We expected {gbp(claim.giftAidAmount)}.</p>
          <div className="mt-5">
            <label className={labelClass} htmlFor="received-amount">Actual amount received (£)</label>
            <input
              id="received-amount"
              autoFocus
              required
              type="number"
              inputMode="decimal"
              step="0.01"
              min="0"
              className={inputClass}
              value={amount}
              onChange={(e) => setAmount(e.target.value)}
              autoComplete="off"
            />
            {valid && (
              <p className={`mt-1.5 text-sm ${diff === 0 ? 'text-muted' : 'text-warn'}`}>
                {diff === 0
                  ? 'Matches what we expected.'
                  : `${gbp(paid)} (expected ${gbp(claim.giftAidAmount)}) · ${diff < 0 ? 'short by' : 'over by'} ${gbp(Math.abs(diff))}`}
              </p>
            )}
          </div>
          {error && <div className="mt-4 rounded-xl border border-bad-line bg-bad-soft px-4 py-3 text-sm text-bad">{error}</div>}
          <div className="mt-6 flex justify-end gap-3">
            <Button type="button" variant="outline" onClick={onClose} disabled={busy}>Cancel</Button>
            <Button type="submit" variant="success" disabled={busy || !valid}>
              {busy && <Loader2 className="h-4 w-4 animate-spin" />} Mark received
            </Button>
          </div>
        </form>
      )}
    </Dialog>
  )
}
