'use client'

import { useState } from 'react'
import { Loader2 } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Dialog } from '@/components/ui/dialog'
import { inputClass, labelClass } from '@/components/ui/primitives'
import { apiJson, gbp } from '@/components/claims/shared'
import type { ClaimDto } from '@/lib/claims'

type Target = { id: string; ref: string; giftAidAmount: number }

/**
 * "Mark received" asks for the amount HMRC actually paid. It can be a few pounds off the claim
 * (rounding, an adjustment, a deduction), so the real figure is kept next to the expected one.
 */
export default function ReceivedDialog({
  claim,
  onClose,
  onDone,
}: {
  claim: Target | null
  onClose: () => void
  onDone: (updated: ClaimDto) => void
}) {
  return (
    <Dialog open={!!claim} onClose={onClose}>
      {/* keyed so the box starts from this claim's expected amount each time */}
      {claim && <ReceivedForm key={claim.id} claim={claim} onClose={onClose} onDone={onDone} />}
    </Dialog>
  )
}

function ReceivedForm({ claim, onClose, onDone }: { claim: Target; onClose: () => void; onDone: (updated: ClaimDto) => void }) {
  const [value, setValue] = useState(claim.giftAidAmount.toFixed(2))
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const amount = Number(value)
  const valid = value.trim() !== '' && Number.isFinite(amount) && amount > 0
  const diff = valid ? Math.round((amount - claim.giftAidAmount) * 100) / 100 : 0

  const submit = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!valid) return
    setBusy(true)
    setError(null)
    try {
      const { claim: updated } = await apiJson<{ claim: ClaimDto }>(`/api/claims/${claim.id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status: 'received', receivedAmount: Math.round(amount * 100) / 100 }),
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
    <form onSubmit={submit}>
      <h2 className="font-display text-xl font-semibold text-ink">Mark {claim.ref} as received</h2>
      <p className="mt-2 text-muted">Expected from HMRC: {gbp(claim.giftAidAmount)}</p>
      <div className="mt-5">
        <label className={labelClass} htmlFor="received-amount">Actual amount received (£)</label>
        <input
          id="received-amount"
          autoFocus
          className={inputClass}
          type="number"
          inputMode="decimal"
          min="0.01"
          step="0.01"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          required
        />
        <p className="mt-1.5 text-sm text-muted">
          {!valid
            ? 'Enter the amount HMRC paid into the bank.'
            : diff === 0
              ? 'Matches the claim.'
              : `${gbp(Math.abs(diff))} ${diff < 0 ? 'less' : 'more'} than expected.`}
        </p>
      </div>
      {error && <div className="mt-4 rounded-xl border border-bad-line bg-bad-soft px-4 py-3 text-sm text-bad">{error}</div>}
      <div className="mt-6 flex justify-end gap-3">
        <Button type="button" variant="outline" onClick={onClose} disabled={busy}>Cancel</Button>
        <Button type="submit" variant="success" disabled={busy || !valid}>
          {busy && <Loader2 className="h-4 w-4 animate-spin" />} Mark received
        </Button>
      </div>
    </form>
  )
}
