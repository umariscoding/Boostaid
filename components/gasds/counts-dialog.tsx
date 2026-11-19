'use client'

import { useState } from 'react'
import { Button } from '@/components/ui/button'
import { Dialog } from '@/components/ui/dialog'
import { inputClass } from '@/components/ui/primitives'
import { apiJson, gbp } from '@/components/claims/shared'
import { countTotal, hasCount, p, ukDate, type Count, type GReceipt } from '@/lib/gasds'

const FIELDS: { k: keyof Count; label: string; step: string }[] = [
  { k: 'n50', label: '£50 notes', step: '1' },
  { k: 'n20', label: '£20 notes', step: '1' },
  { k: 'n10', label: '£10 notes', step: '1' },
  { k: 'n5', label: '£5 notes', step: '1' },
  { k: 'coins', label: 'Coins (£)', step: '0.01' },
]

/** Step 2: for each cash collection in the claim, the notes and coins counted (or a photo of the count sheet) and the bank slip. */
export default function CountsDialog({
  open,
  receipts,
  onClose,
  onChanged,
}: {
  open: boolean
  receipts: GReceipt[]
  onClose: () => void
  /** Called after anything is saved, so the page can reload. */
  onChanged: () => void
}) {
  return (
    <Dialog open={open} onClose={onClose} wide>
      {open && (
        <div className="space-y-4">
          <h2 className="font-display text-xl font-semibold text-ink">Add the counts</h2>
          <p className="text-sm text-muted">
            For each collection, type the notes and coins that were counted, or attach a photo of the count sheet. Any £50 notes are taken
            off, because HMRC doesn’t allow them.
          </p>
          <div className="space-y-3">
            {receipts.map((r) => <CountBox key={r.id} r={r} onChanged={onChanged} />)}
          </div>
          <div className="flex justify-end"><Button variant="dark" onClick={onClose}>Done</Button></div>
        </div>
      )}
    </Dialog>
  )
}

function CountBox({ r, onChanged }: { r: GReceipt; onChanged: () => void }) {
  const [draft, setDraft] = useState<Record<keyof Count, string>>({
    n50: r.count?.n50 ? String(r.count.n50) : '',
    n20: r.count?.n20 ? String(r.count.n20) : '',
    n10: r.count?.n10 ? String(r.count.n10) : '',
    n5: r.count?.n5 ? String(r.count.n5) : '',
    coins: r.count?.coins ? String(r.count.coins) : '',
  })
  const [slip, setSlip] = useState(r.slip ?? '')
  const [error, setError] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)

  const asCount = (): Count => ({
    n50: Math.max(0, Math.trunc(Number(draft.n50) || 0)),
    n20: Math.max(0, Math.trunc(Number(draft.n20) || 0)),
    n10: Math.max(0, Math.trunc(Number(draft.n10) || 0)),
    n5: Math.max(0, Math.trunc(Number(draft.n5) || 0)),
    coins: Math.max(0, Number(draft.coins) || 0),
  })
  const total = countTotal(asCount())
  // What should have been counted: the bank's figure if we have it, otherwise the receipt
  const expected = r.counted ?? r.amount
  const fix = async (body: unknown) => {
    setBusy(true)
    setError(null)
    try {
      await apiJson(`/api/gasds/receipts/${r.id}`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
      onChanged()
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(false)
    }
  }
  const attach = async (file: File | undefined) => {
    if (!file) return
    setBusy(true)
    setError(null)
    try {
      await apiJson(`/api/gasds/receipts/${r.id}/photo`, { method: 'POST', headers: { 'Content-Type': file.type || 'application/octet-stream', 'x-file-name': encodeURIComponent(file.name) }, body: file })
      onChanged()
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(false)
    }
  }
  const ok = hasCount(r)

  return (
    <div className={`space-y-3 rounded-2xl border p-4 ${ok ? 'border-good-line bg-good-soft' : 'border-line'}`}>
      <div className="flex flex-wrap justify-between gap-2 font-semibold text-ink">
        <span>{ukDate(r.date)} · {r.notes || 'Collection'}</span>
        <span className="tabular-nums">{gbp(r.amount)}</span>
      </div>
      <div className="grid grid-cols-2 gap-2 sm:grid-cols-5">
        {FIELDS.map((f) => (
          <label key={f.k} className="flex flex-col gap-1 text-sm font-semibold text-ink">
            {f.label}
            <input
              className={inputClass}
              type="number"
              min="0"
              step={f.step}
              inputMode={f.k === 'coins' ? 'decimal' : 'numeric'}
              value={draft[f.k]}
              onChange={(e) => setDraft({ ...draft, [f.k]: e.target.value })}
              onBlur={() => fix({ count: total > 0 ? asCount() : null })}
            />
          </label>
        ))}
      </div>
      <div className="flex flex-wrap items-center justify-between gap-2 text-sm">
        <span className="space-x-2">
          {total > 0 && (p(total) === p(expected)
            ? <span className="rounded-md bg-good-soft px-2 py-1 font-bold text-good">Adds up ✓</span>
            : <span className="rounded-md bg-bad-soft px-2 py-1 font-bold text-bad">Adds up to {gbp(total)}, but {r.counted != null ? 'the bank counted' : 'the receipt says'} {gbp(expected)}</span>)}
          {asCount().n50 > 0 && <span className="rounded-md bg-warn-soft px-2 py-1 font-bold text-warn">{gbp(asCount().n50 * 50)} in £50 notes taken off</span>}
        </span>
        <label className="text-muted">or attach a photo <input type="file" accept="image/jpeg,image/png,image/webp,application/pdf" onChange={(e) => attach(e.target.files?.[0])} className="max-w-[220px] text-xs" /></label>
      </div>
      {r.photo && <p className="text-sm text-muted">Photo attached: {r.photo}</p>}
      <label className="flex max-w-xs flex-col gap-1 text-sm text-muted">
        Bank slip number{r.slip ? ' (from the banking file)' : ''}
        <input className={inputClass} value={slip} onChange={(e) => setSlip(e.target.value)} onBlur={() => slip.trim() !== (r.slip ?? '') && fix({ slip })} placeholder="e.g. 300101" maxLength={60} />
      </label>
      {busy && <p className="text-xs text-muted">Saving…</p>}
      {error && <p role="alert" className="text-sm font-semibold text-bad">{error}</p>}
    </div>
  )
}
