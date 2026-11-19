'use client'

import { useCallback, useEffect, useMemo, useState } from 'react'
import { Loader2 } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Card, StatTile, inputClass, labelClass, th, thr, td, tdr } from '@/components/ui/primitives'
import { Dialog } from '@/components/ui/dialog'
import { Pagination } from '@/components/ui/pagination'
import { Skeleton } from '@/components/ui/skeleton'
import { apiJson, fmtDay, gbp } from '@/components/claims/shared'
import { STATUS_LABEL } from '@/lib/recovery'
import type { RecoveryRowDto } from '@/app/api/recovery/route'

interface Data {
  rows: RecoveryRowDto[]
  emailConfigured: boolean
  tiles: {
    heldBack: { count: number; giftAid: number }
    waiting: { count: number; giftAid: number }
    recovered: { count: number; giftAid: number }
    noResponse: { count: number }
  }
}
interface Preview { donors: number; donations: number; from: string; replyTo: string | null; to: string; subject: string; html: string; emailConfigured: boolean }
interface SendResult { sent: number; failed: number; emailConfigured: boolean; testLinks: { to: string; name: string; url: string }[] }

const PILL: Record<string, string> = {
  none: 'bg-bad-pill text-bad',
  sent: 'bg-info-pill text-info',
  reminded: 'bg-warn-pill text-warn',
  no_response: 'bg-[#f1eff6] text-muted',
  recovered: 'bg-good-pill text-good',
}
const ukDate = (iso: string | null) => (iso ? iso.split('-').reverse().join('/') : '')

/**
 * Donations held back in Needs review, with the donor's email. Tick donors, press Email donors, and their reply fixes the row.
 * Recovered donations then join the next claim.
 */
export default function RecoveryPanel({
  charityId,
  refreshKey = 0,
  onCreateClaim,
}: {
  charityId: string
  refreshKey?: number
  /** Starts a claim from the recovered donations. */
  onCreateClaim?: () => void
}) {
  const [data, setData] = useState<{ charityId: string; value: Data } | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [selected, setSelected] = useState<Set<string>>(new Set())
  const [preview, setPreview] = useState<Preview | null>(null)
  const [busy, setBusy] = useState(false)
  const [result, setResult] = useState<SendResult | null>(null)
  const [fixRow, setFixRow] = useState<RecoveryRowDto | null>(null)
  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = useState(25)

  const load = useCallback(async () => {
    try {
      const value = await apiJson<Data>(`/api/recovery?charity=${charityId}`)
      setData({ charityId, value })
      setError(null)
    } catch (e) {
      setError((e as Error).message)
    }
  }, [charityId])

  useEffect(() => {
    setSelected(new Set())
    setPage(1)
    load()
  }, [load, refreshKey])

  const d = data && data.charityId === charityId ? data.value : null
  const rows = useMemo(() => d?.rows ?? [], [d])
  const canSend = useMemo(() => rows.filter((r) => r.status === 'none' && r.email), [rows])

  if (error && !d) return <div role="alert" className="rounded-xl border border-bad-line bg-bad-soft px-4 py-3 text-sm font-medium text-bad">{error}</div>
  if (!d) {
    return (
      <section className="space-y-4" aria-busy="true">
        <div className="grid grid-cols-1 gap-5 md:grid-cols-4">{[0, 1, 2, 3].map((i) => <Skeleton key={i} className="h-28 rounded-2xl" />)}</div>
        <Skeleton className="h-48 w-full rounded-2xl" />
      </section>
    )
  }
  if (!rows.length) return null // nothing has ever been held back for this charity

  const pageCount = Math.max(1, Math.ceil(rows.length / pageSize))
  const pageNo = Math.min(page, pageCount)
  const shown = rows.slice((pageNo - 1) * pageSize, pageNo * pageSize)
  const toggle = (id: string) => setSelected((s) => { const n = new Set(s); if (n.has(id)) n.delete(id); else n.add(id); return n })

  const openPreview = async () => {
    setBusy(true)
    setError(null)
    try {
      setPreview(await apiJson<Preview>('/api/recovery/send', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ charityId, rowIds: [...selected], preview: true }) }))
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusy(false)
    }
  }
  const send = async () => {
    setBusy(true)
    try {
      const r = await apiJson<SendResult>('/api/recovery/send', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ charityId, rowIds: [...selected] }) })
      setPreview(null)
      setSelected(new Set())
      setResult(r)
      await load()
    } catch (e) {
      setPreview(null)
      setError((e as Error).message)
    } finally {
      setBusy(false)
    }
  }

  const t = d.tiles
  const note = (r: RecoveryRowDto) =>
    r.status === 'sent' ? `Sent ${fmtDay(r.emailSentAt)}`
    : r.status === 'reminded' ? `Reminder ${fmtDay(r.reminderAt)}`
    : r.status === 'recovered' ? (r.fixedBy === 'donor' ? `Replied ${fmtDay(r.repliedAt)}` : 'Fixed by hand')
    : r.status === 'no_response' ? 'Fix by hand if you can'
    : r.email ? '' : 'No email. Fix by hand'

  return (
    <section className="space-y-5" aria-label="Gift Aid recovery">
      <div>
        <h2 className="font-display text-2xl font-semibold text-ink">Held back for review</h2>
        <p className="mt-1 text-muted">Email donors whose details are missing. Their reply fixes the row and it joins your next claim. Ready donations are never held back while you wait.</p>
      </div>

      <div className="grid grid-cols-1 gap-5 sm:grid-cols-2 xl:grid-cols-4">
        <StatTile tone="bad" label="Held back for review" value={t.heldBack.count.toLocaleString()} caption={`${gbp(t.heldBack.giftAid)} Gift Aid`} />
        <StatTile tone="warn" label="Waiting for the donor" value={t.waiting.count.toLocaleString()} caption={`${gbp(t.waiting.giftAid)} Gift Aid`} />
        <StatTile tone="good" label="Recovered, ready to claim" value={gbp(t.recovered.giftAid)} caption={`${t.recovered.count.toLocaleString()} donation${t.recovered.count === 1 ? '' : 's'} fixed`} />
        <StatTile label="No response" value={t.noResponse.count.toLocaleString()} caption="after 2 emails" />
      </div>

      {t.recovered.count > 0 && onCreateClaim && (
        <div className="flex flex-wrap items-center justify-between gap-4 rounded-2xl border border-good-line bg-good-soft px-6 py-5 text-good">
          <div>
            <div className="font-display text-2xl font-bold tabular-nums">{gbp(t.recovered.giftAid)} recovered and waiting</div>
            <div>{t.recovered.count} donation{t.recovered.count === 1 ? '' : 's'} fixed. Add {t.recovered.count === 1 ? 'it' : 'them'} to your next claim?</div>
          </div>
          <Button variant="success" onClick={onCreateClaim}>Create claim</Button>
        </div>
      )}

      {!d.emailConfigured && (
        <div className="rounded-xl bg-warn-soft px-4 py-3 text-sm text-warn">
          Email sending isn’t connected yet, so nothing will actually leave BoostAid. You can still try the whole flow: after you press Email donors you’ll get the donors’ links to open yourself.
        </div>
      )}
      {error && <div role="alert" className="rounded-xl border border-bad-line bg-bad-soft px-4 py-3 text-sm font-medium text-bad">{error}</div>}

      <Card className="overflow-hidden">
        <div className="flex flex-wrap items-center gap-3 border-b border-line px-5 py-3">
          <Button size="sm" variant="outline" disabled={!canSend.length} onClick={() => setSelected(new Set(canSend.map((r) => r.id)))}>
            Select all with email ({canSend.length})
          </Button>
          <Button size="sm" variant="dark" disabled={!selected.size || busy} onClick={openPreview}>
            {busy && !preview && <Loader2 className="h-3.5 w-3.5 animate-spin" />} Email donors{selected.size ? ` (${selected.size})` : ''}
          </Button>
          {selected.size > 0 && <button type="button" className="cursor-pointer text-sm text-muted underline" onClick={() => setSelected(new Set())}>Clear selection</button>}
        </div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[860px] text-base">
            <thead>
              <tr className="border-b border-line">
                <th className={`${th} w-10`}><span className="sr-only">Select</span></th>
                <th className={th}>Why it’s held back</th>
                <th className={th}>Donor (as uploaded)</th>
                <th className={th}>Address</th>
                <th className={th}>Postcode</th>
                <th className={th}>Email</th>
                <th className={thr}>Amount</th>
                <th className={th}>Status</th>
                <th className={th}><span className="sr-only">Actions</span></th>
              </tr>
            </thead>
            <tbody>
              {shown.map((r) => {
                const can = r.status === 'none' && !!r.email
                const fixed = r.status === 'recovered' && r.fixed
                return (
                  <tr key={r.id} className="border-b border-line-soft align-top last:border-0">
                    <td className={td}>{can && <input type="checkbox" className="h-4 w-4 cursor-pointer" checked={selected.has(r.id)} onChange={() => toggle(r.id)} aria-label={`Select ${r.name}`} />}</td>
                    <td className={td}>
                      {fixed ? <span className="rounded-md bg-good-pill px-2 py-1 text-sm font-semibold text-good">{r.fixedBy === 'donor' ? 'Fixed by donor' : 'Fixed by hand'}</span>
                        : <span className="rounded-md bg-bad-pill px-2 py-1 text-sm font-semibold text-bad">{r.reason.length > 60 ? r.reason.slice(0, 57) + '…' : r.reason}</span>}
                      <div className="mt-1 text-xs text-muted">{ukDate(r.date)}</div>
                    </td>
                    <td className={`${td} text-ink`}>{fixed ? `${r.fixed!.title} ${r.fixed!.first} ${r.fixed!.last}`.trim() : r.name}</td>
                    <td className={`${td} text-ink`}>{fixed ? r.fixed!.house : r.house || '–'}</td>
                    <td className={`${td} whitespace-nowrap text-ink`}>{(fixed ? r.fixed!.postcode : r.postcode) || '–'}</td>
                    <td className={`${td} text-sm text-muted`}>{r.email || 'none'}</td>
                    <td className={`${tdr} text-ink`}>{gbp(r.amount)}</td>
                    <td className={td}>
                      <span className={`rounded-md px-2 py-1 text-sm font-semibold ${PILL[r.status] ?? PILL.none}`}>{STATUS_LABEL[r.status] ?? r.status}</span>
                      {note(r) && <div className="mt-1 text-xs text-muted">{note(r)}</div>}
                    </td>
                    <td className={`${td} text-right`}>
                      {r.status !== 'recovered' && <Button size="sm" variant="outline" onClick={() => setFixRow(r)}>Fix by hand</Button>}
                    </td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
        <Pagination page={pageNo} pageSize={pageSize} total={rows.length} onPage={setPage} onPageSize={(n) => { setPageSize(n); setPage(1) }} />
      </Card>

      {/* The email, before it goes */}
      <Dialog open={!!preview} onClose={() => !busy && setPreview(null)} wide>
        {preview && (
          <div className="space-y-4">
            <h2 className="font-display text-xl font-semibold text-ink">Email {preview.donors} donor{preview.donors === 1 ? '' : 's'}</h2>
            <p className="text-sm text-muted">
              One email per donor, however many of their donations are held back ({preview.donations} donation{preview.donations === 1 ? '' : 's'} in all). Each donor gets
              their own link. This is the first donor’s email:
            </p>
            <div className="overflow-hidden rounded-xl border border-line text-sm">
              <div className="space-y-0.5 border-b border-line bg-wash px-4 py-3">
                <div><b>From:</b> {preview.from} · <b>Reply-to:</b> {preview.replyTo || 'not set (add a contact email in Settings)'}</div>
                <div><b>To:</b> {preview.to}</div>
                <div><b>Subject:</b> {preview.subject}</div>
              </div>
              <div className="px-4 py-3" dangerouslySetInnerHTML={{ __html: preview.html }} />
            </div>
            <div className="rounded-xl bg-info-soft px-4 py-3 text-sm text-info">If they don’t reply, BoostAid sends one reminder after 14 days. Never more than 2 emails.</div>
            {!preview.emailConfigured && <div className="rounded-xl bg-warn-soft px-4 py-3 text-sm text-warn">Email sending isn’t connected yet: nothing will be delivered. You’ll get the links to try the donor’s form.</div>}
            <div className="flex justify-end gap-3">
              <Button variant="outline" onClick={() => setPreview(null)} disabled={busy}>Cancel</Button>
              <Button variant="dark" onClick={send} disabled={busy}>{busy && <Loader2 className="h-4 w-4 animate-spin" />} Send {preview.donors} email{preview.donors === 1 ? '' : 's'}</Button>
            </div>
          </div>
        )}
      </Dialog>

      {/* After sending */}
      <Dialog open={!!result} onClose={() => setResult(null)} wide>
        {result && (
          <div className="space-y-4">
            <h2 className="font-display text-xl font-semibold text-ink">
              {result.emailConfigured ? `${result.sent} donation${result.sent === 1 ? '' : 's'} emailed` : 'Email isn’t connected yet'}
            </h2>
            {result.failed > 0 && <div className="rounded-xl bg-bad-soft px-4 py-3 text-sm text-bad">{result.failed} could not be sent and are back in Needs review. Please try again.</div>}
            {result.emailConfigured ? (
              <p className="text-muted">Each row now shows Email sent with today’s date. When a donor replies, their row becomes Ready – recovered.</p>
            ) : (
              <>
                <p className="text-muted">Nothing was delivered, but the rows are marked Email sent so you can try the whole flow. Open a link to fill in the form as the donor:</p>
                <ul className="max-h-64 space-y-1 overflow-auto text-sm">
                  {result.testLinks.map((l) => (
                    <li key={l.url} className="flex flex-wrap items-center justify-between gap-2 rounded-lg border border-line px-3 py-2">
                      <span>{l.name || l.to} <span className="text-muted">({l.to})</span></span>
                      <a className="font-medium text-brand underline" href={l.url} target="_blank" rel="noreferrer">Open the donor’s form</a>
                    </li>
                  ))}
                </ul>
              </>
            )}
            <div className="flex justify-end"><Button variant="dark" onClick={() => setResult(null)}>Done</Button></div>
          </div>
        )}
      </Dialog>

      <FixDialog charityId={charityId} row={fixRow} onClose={() => setFixRow(null)} onDone={load} />
    </section>
  )
}

/** The charity fixes a row itself (it runs through the same checks as a donor's answer). There is no fee on these. */
function FixDialog({ charityId, row, onClose, onDone }: { charityId: string; row: RecoveryRowDto | null; onClose: () => void; onDone: () => Promise<void> }) {
  return (
    <Dialog open={!!row} onClose={onClose}>
      {row && <FixForm key={row.id} charityId={charityId} row={row} onClose={onClose} onDone={onDone} />}
    </Dialog>
  )
}

function FixForm({ charityId, row, onClose, onDone }: { charityId: string; row: RecoveryRowDto; onClose: () => void; onDone: () => Promise<void> }) {
  const [first, ...rest] = row.name.split(' ')
  const [f, setF] = useState({ title: '', first: first ?? '', last: rest.join(' '), house: row.house, postcode: row.postcode })
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const submit = async (e: React.FormEvent) => {
    e.preventDefault()
    setBusy(true)
    setError(null)
    try {
      await apiJson('/api/recovery/fix', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ charityId, rowId: row.id, ...f }) })
      await onDone()
      onClose()
    } catch (err) {
      setError((err as Error).message)
    } finally {
      setBusy(false)
    }
  }
  const set = (k: keyof typeof f) => (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => setF({ ...f, [k]: e.target.value })
  return (
    <form onSubmit={submit} className="space-y-4">
      <h2 className="font-display text-xl font-semibold text-ink">Fix by hand</h2>
      <p className="text-sm text-muted">{gbp(row.amount)} on {ukDate(row.date)}. Held back because: {row.reason}. Type the correct details. There’s no BoostAid fee on rows you fix yourself.</p>
      <div className="grid grid-cols-[110px_1fr] gap-3">
        <div><label className={labelClass} htmlFor="fx-title">Title</label>
          <select id="fx-title" className={inputClass} value={f.title} onChange={set('title')}>{['', 'Mr', 'Mrs', 'Ms', 'Miss', 'Dr'].map((t) => <option key={t}>{t}</option>)}</select></div>
        <div><label className={labelClass} htmlFor="fx-first">First name</label><input id="fx-first" className={inputClass} value={f.first} onChange={set('first')} /></div>
      </div>
      <div><label className={labelClass} htmlFor="fx-last">Last name</label><input id="fx-last" className={inputClass} value={f.last} onChange={set('last')} /></div>
      <div className="grid grid-cols-2 gap-3">
        <div><label className={labelClass} htmlFor="fx-house">House name or number</label><input id="fx-house" className={inputClass} value={f.house} onChange={set('house')} /></div>
        <div><label className={labelClass} htmlFor="fx-pc">Postcode</label><input id="fx-pc" className={inputClass} value={f.postcode} onChange={set('postcode')} /></div>
      </div>
      {error && <div role="alert" className="rounded-xl border border-bad-line bg-bad-soft px-4 py-3 text-sm font-medium text-bad">{error}</div>}
      <div className="flex justify-end gap-3">
        <Button type="button" variant="outline" onClick={onClose} disabled={busy}>Cancel</Button>
        <Button type="submit" variant="dark" disabled={busy}>{busy && <Loader2 className="h-4 w-4 animate-spin" />} Save</Button>
      </div>
    </form>
  )
}
