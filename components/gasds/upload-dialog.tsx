'use client'

import { useMemo, useState } from 'react'
import { Download, FileUp, Loader2 } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Dialog } from '@/components/ui/dialog'
import { apiJson, gbp } from '@/components/claims/shared'
import { readGasdsFiles, type ReadResult } from '@/lib/gasds-read'
import { p, previewUpload, showNo, taxYearLabel, topUpOf, ukDate, who, type GReceipt, type UploadOutcome } from '@/lib/gasds'

export interface UploadContext {
  charityId: string
  today: string
  stored: GReceipt[]
  claimedByYear: Record<number, number>
  giftAidByYear: Record<number, number>
}

const ICON = { in: '✅', wait: '⏳', out: '❌' } as const
const COLOUR = { in: 'text-good', wait: 'text-warn', out: 'text-muted' } as const

/** Step 1: choose the files, see what would happen to every receipt, then upload. */
export default function UploadDialog({
  open,
  ctx,
  onClose,
  onDone,
}: {
  open: boolean
  ctx: UploadContext
  onClose: () => void
  onDone: (message: string, bestYear: number | null) => void
}) {
  return (
    <Dialog open={open} onClose={onClose} wide>
      {open && <UploadForm ctx={ctx} onClose={onClose} onDone={onDone} />}
    </Dialog>
  )
}

function UploadForm({ ctx, onClose, onDone }: { ctx: UploadContext; onClose: () => void; onDone: UploadDialogDone }) {
  const [read, setRead] = useState<ReadResult | null>(null)
  const [reading, setReading] = useState(false)
  const [names, setNames] = useState<string[]>([])
  const [over, setOver] = useState(false)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const pick = async (files: FileList | null) => {
    if (!files?.length) return
    setNames([...files].map((f) => f.name))
    setReading(true)
    setError(null)
    try {
      const list = await Promise.all([...files].map(async (f) => ({ name: f.name, data: new Uint8Array(await f.arrayBuffer()) })))
      setRead(readGasdsFiles(list))
    } finally {
      setReading(false)
    }
  }

  const outcomes: UploadOutcome[] = useMemo(
    () => (read ? previewUpload(ctx.stored, read.receipts, read.banking, ctx.claimedByYear, ctx.giftAidByYear, ctx.today) : []),
    [read, ctx]
  )

  const tally = useMemo(() => {
    const t = { inN: 0, inP: 0, wait: 0, out: 0, seen: 0, why: new Map<string, number>(), fullYears: new Set<number>(), lowGiftAid: new Set<number>() }
    for (const o of outcomes) {
      if (o.action === 'same' || o.action === 'claimed') t.seen++
      if (o.status.group === 'in') { t.inN++; t.inP += p(o.used) }
      else if (o.status.group === 'wait') t.wait++
      else if (o.status.label === 'Not included') {
        t.out++
        t.why.set(o.status.reason, (t.why.get(o.status.reason) ?? 0) + 1)
        if (o.status.reason.endsWith('has reached its limit')) t.fullYears.add(o.receipt.year)
        if (o.status.reason.startsWith('not enough Gift Aid')) t.lowGiftAid.add(o.receipt.year)
      }
    }
    return t
  }, [outcomes])

  const shorts = (read?.banking ?? []).flatMap((b) =>
    b.items.filter((it) => p(it.counted) !== p(it.amount)).map((it) => `${it.no} was ${gbp(Math.abs(it.counted - it.amount))} ${it.counted < it.amount ? 'short' : 'over'} at the bank (slip ${b.slip}), so ${gbp(it.counted)} counts`)
  )
  const nothing = !read || (!read.receipts.length && !read.banking.length)

  const upload = async () => {
    if (!read) return
    setBusy(true)
    setError(null)
    try {
      const r = await apiJson<{ added: number; updated: number }>('/api/gasds/receipts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ charityId: ctx.charityId, receipts: read.receipts, banking: read.banking }),
      })
      const years = new Map<number, number>()
      for (const o of outcomes) if (o.action === 'new') years.set(o.receipt.year, (years.get(o.receipt.year) ?? 0) + 1)
      const best = [...years.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? null
      const n = r.added + r.updated
      onDone(n ? `Uploaded. ${r.added} receipt${r.added === 1 ? '' : 's'} added${r.updated ? `, ${r.updated} updated` : ''}.` : 'Nothing new in those files.', best)
      onClose()
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="space-y-4">
      <h2 className="font-display text-xl font-semibold text-ink">Upload your receipts</h2>
      {/* The whole box is the button: click it, or drop files on it */}
      <label
        htmlFor="gasds-files"
        onDragOver={(e) => { e.preventDefault(); setOver(true) }}
        onDragLeave={() => setOver(false)}
        onDrop={(e) => { e.preventDefault(); setOver(false); pick(e.dataTransfer.files) }}
        className={`block cursor-pointer rounded-2xl border-2 border-dashed p-8 text-center transition-colors ${over ? 'border-brand bg-brand-soft/60' : 'border-[#cfc9e8] bg-[#faf9ff] hover:border-brand/50 hover:bg-brand-soft/40'}`}
      >
        <div className="mx-auto mb-3 flex h-12 w-12 items-center justify-center rounded-full border border-line bg-white"><FileUp className="h-5 w-5 text-brand" /></div>
        <div className="font-display text-lg font-semibold text-ink">{names.length ? `${names.length} file${names.length === 1 ? '' : 's'} chosen` : 'Drop your files here'}</div>
        <div className="mt-1 text-sm text-muted">{names.length ? 'Click to choose different files' : 'or click to browse. CSV or Excel, and you can pick several at once.'}</div>
        <span className="mt-4 inline-flex h-11 items-center rounded-full bg-ink px-6 text-sm font-medium text-white">{names.length ? 'Choose other files' : 'Choose files'}</span>
        <input id="gasds-files" type="file" multiple accept=".csv,.xlsx,.xls,text/csv" onChange={(e) => { pick(e.target.files); e.target.value = '' }} className="sr-only" />
      </label>
      {names.length > 0 && <ul className="space-y-0.5 text-sm text-muted">{names.map((n) => <li key={n} className="truncate">✓ {n}</li>)}</ul>}
      <p className="text-sm text-muted">
        Your receipt book exports (receipt list and banking), or the BoostAid template. They are read on this computer first, so you can check them before anything is saved.
      </p>
      <a href="/templates/BoostAid_GASDS_template.xlsx" className="inline-flex items-center gap-1.5 text-sm font-medium text-brand underline">
        <Download className="h-4 w-4" /> Download the BoostAid template
      </a>
      {reading && <p className="text-sm text-muted">Reading…</p>}
      {read?.errors.length ? <div role="alert" className="rounded-xl border border-bad-line bg-bad-soft px-4 py-3 text-sm font-medium text-bad">{read.errors.join(' ')}</div> : null}

      {read && !nothing && (
        <div className="space-y-3">
          <div className={`rounded-2xl border px-4 py-4 ${tally.inP ? 'border-good-line bg-good-soft text-good' : 'border-line bg-white text-ink'}`}>
            <div>{tally.inP ? 'This adds' : 'This adds nothing new to claim'}</div>
            {tally.inP > 0 && (
              <>
                <div className="font-display text-3xl font-bold tabular-nums">{gbp(tally.inP / 100)}</div>
                <div>HMRC would pay {gbp(topUpOf(tally.inP / 100))} on it.</div>
              </>
            )}
          </div>
          <div className="flex flex-wrap gap-2 text-sm font-semibold">
            <span className="rounded-full bg-good-soft px-3 py-1.5 text-good">✅ {tally.inN} included</span>
            <span className="rounded-full bg-warn-soft px-3 py-1.5 text-warn">⏳ {tally.wait} waiting for the bank</span>
            <span className="rounded-full bg-wash px-3 py-1.5 text-muted">❌ {tally.out} not included</span>
            {tally.seen > 0 && <span className="rounded-full bg-wash px-3 py-1.5 text-muted">{tally.seen} seen before</span>}
          </div>
          {tally.out > 0 && (
            <p className="text-sm text-muted">Not included: {[...tally.why.entries()].map(([k, n]) => `${n} ${k}`).join(' · ')}.</p>
          )}
          {tally.fullYears.size > 0 && (
            <div className="rounded-xl bg-info-soft px-4 py-3 text-sm text-info">
              ✓ The limit is already reached for {[...tally.fullYears].map(taxYearLabel).join(' and ')}. Those receipts are saved as records only.
            </div>
          )}
          {tally.lowGiftAid.size > 0 && (
            <div className="rounded-xl bg-bad-soft px-4 py-3 text-sm text-bad">
              Not enough Gift Aid has been claimed for {[...tally.lowGiftAid].map(taxYearLabel).join(' and ')}. HMRC only pays on small donations up to 10 times it.
              Those receipts are saved as records, and count once more Gift Aid is claimed (or added under “How it is worked out”).
            </div>
          )}
          {read.banking.length > 0 && (
            <p className="text-sm text-muted">Bank slips found: {read.banking.map((b) => b.slip).join(', ')}. They are added to the receipts automatically.</p>
          )}
          {shorts.map((s) => <div key={s} className="rounded-xl bg-warn-soft px-4 py-3 text-sm text-warn">{s}.</div>)}

          <details className="rounded-xl border border-line">
            <summary className="cursor-pointer px-4 py-3 font-semibold text-ink">See every receipt ({outcomes.length})</summary>
            <div className="max-h-80 overflow-auto">
              <table className="w-full min-w-[600px] text-sm">
                <thead className="sticky top-0 bg-white text-left text-muted">
                  <tr><th className="px-4 py-2 font-medium">Receipt</th><th className="px-4 py-2 font-medium">Date</th><th className="px-4 py-2 font-medium">From</th><th className="px-4 py-2 text-right font-medium">Amount</th><th className="px-4 py-2 font-medium">Included?</th></tr>
                </thead>
                <tbody>
                  {outcomes.slice(0, 500).map((o) => (
                    <tr key={o.receipt.no} className="border-t border-line-soft">
                      <td className="px-4 py-2">{showNo(o.receipt.no)}</td>
                      <td className="px-4 py-2 whitespace-nowrap">{ukDate(o.receipt.date)}</td>
                      <td className="px-4 py-2">{who(o.receipt)}{!o.receipt.gift && <div className="text-xs text-muted">collection</div>}</td>
                      <td className="px-4 py-2 text-right tabular-nums">{gbp(o.receipt.amount)}</td>
                      <td className={`px-4 py-2 font-semibold ${COLOUR[o.status.group]}`}>{ICON[o.status.group]} {o.status.label}{o.status.reason && <div className="text-xs font-normal text-muted">{o.status.reason}</div>}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {outcomes.length > 500 && <p className="px-4 py-2 text-xs text-muted">Showing the first 500 of {outcomes.length}.</p>}
            </div>
          </details>
        </div>
      )}

      {error && <div role="alert" className="rounded-xl border border-bad-line bg-bad-soft px-4 py-3 text-sm font-medium text-bad">{error}</div>}
      <div className="flex justify-end gap-3">
        <Button type="button" variant="outline" onClick={onClose} disabled={busy}>Cancel</Button>
        <Button type="button" variant="dark" onClick={upload} disabled={busy || nothing}>
          {busy && <Loader2 className="h-4 w-4 animate-spin" />} Upload
        </Button>
      </div>
    </div>
  )
}

type UploadDialogDone = (message: string, bestYear: number | null) => void
