'use client'

import { useEffect, useState } from 'react'
import { Loader2 } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Dialog } from '@/components/ui/dialog'
import { inputClass, labelClass } from '@/components/ui/primitives'
import { apiJson } from '@/components/claims/shared'
import type { ClaimDto } from '@/lib/claims'

/**
 * "Mark submitted" asks for the reference HMRC gave when the file was uploaded to Charities Online,
 * so it is stored on the claim from the start.
 */
export default function SubmitDialog({
  claim,
  onClose,
  onDone,
}: {
  claim: { id: string; ref: string; hmrcRef?: string | null } | null
  onClose: () => void
  onDone: (updated: ClaimDto) => void
}) {
  const [hmrcRef, setHmrcRef] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    if (claim) {
      setHmrcRef(claim.hmrcRef ?? '')
      setError(null)
    }
  }, [claim])

  const submit = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!claim) return
    if (!hmrcRef.trim()) {
      setError('Type the HMRC reference to save.')
      return
    }
    setBusy(true)
    setError(null)
    try {
      const { claim: updated } = await apiJson<{ claim: ClaimDto }>(`/api/claims/${claim.id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status: 'submitted', hmrcRef: hmrcRef.trim() }),
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
          <h2 className="font-display text-xl font-semibold text-ink">Mark {claim.ref} as submitted</h2>
          <p className="mt-2 text-muted">
            Once you have uploaded this file to HMRC Charities Online, add the reference HMRC gave you. It is saved on the claim
            so you can find it later.
          </p>
          <div className="mt-5">
            <label className={labelClass} htmlFor="submit-hmrc-ref">HMRC reference</label>
            <input
              id="submit-hmrc-ref"
              autoFocus
              className={inputClass}
              value={hmrcRef}
              onChange={(e) => setHmrcRef(e.target.value)}
              maxLength={100}
              required
              placeholder="e.g. the reference shown after the upload"
              autoComplete="off"
            />
            <p className="mt-1.5 text-sm text-muted">Required. You can correct it later on the claim.</p>
          </div>
          {error && <div className="mt-4 rounded-xl border border-bad-line bg-bad-soft px-4 py-3 text-sm text-bad">{error}</div>}
          <div className="mt-6 flex justify-end gap-3">
            <Button type="button" variant="outline" onClick={onClose} disabled={busy}>Cancel</Button>
            <Button type="submit" variant="dark" disabled={busy || !hmrcRef.trim()}>
              {busy && <Loader2 className="h-4 w-4 animate-spin" />} Mark submitted
            </Button>
          </div>
        </form>
      )}
    </Dialog>
  )
}
