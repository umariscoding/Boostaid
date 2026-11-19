'use client'

import { useCallback, useEffect, useMemo, useState } from 'react'
import { Download, Loader2 } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Card, PageHeader, StatTile, th, thr, td, tdr } from '@/components/ui/primitives'
import { Dialog } from '@/components/ui/dialog'
import { Skeleton } from '@/components/ui/skeleton'
import { apiJson, fmtDay, gbp } from '@/components/claims/shared'

interface Row {
  claimId: string
  charity: string
  ref: string
  claimStatus: 'pending' | 'submitted' | 'received'
  recoveredGiftAid: number
  recoveredPaid: number
  short: boolean
  fee: number
  stage: 'received' | 'sent' | 'overdue' | 'pending' | 'paid'
  stageAt: string | null
  stuck: boolean
  invoice: { id: string; number: string; status: string; sentAt: string; dueAt: string; paidAt: string | null; delivered: boolean; sentTo: string } | null
}
interface Overview {
  emailConfigured: boolean
  tiles: { earned: number; owed: number; overdue: number; waiting: number }
  rows: Row[]
  readyToInvoice: number
  byCharity: { charity: string; earned: number; owed: number }[]
  archive: { charity: string; number: string; ref: string; sentAt: string; paidAt: string | null; fee: number }[]
  notifications: { id: string; type: string; text: string; at: string }[]
}
interface Preview {
  number: string
  sentAt: string
  dueAt: string
  charity: string
  claimRef: string
  recoveredPaid: number
  fee: number
  received: number | null
  to: string
  canSend: boolean
  donations: unknown[]
}

const STAGE: Record<Row['stage'], { label: string; cls: string }> = {
  received: { label: 'Received', cls: 'bg-good-pill text-good' },
  sent: { label: 'Invoice sent', cls: 'bg-info-pill text-info' },
  overdue: { label: 'Invoice overdue', cls: 'bg-bad-pill text-bad' },
  pending: { label: 'Pending', cls: 'bg-warn-pill text-warn' },
  paid: { label: 'Paid', cls: 'bg-good-pill text-good' },
}
const monthOf = (iso: string | null) => (iso ? iso.slice(0, 7) : '')
const monthName = (ym: string) => new Date(ym + '-15T12:00:00Z').toLocaleDateString('en-GB', { month: 'long', year: 'numeric', timeZone: 'UTC' })

/** BoostAid's own page: what every charity owes, invoices, and what happened. Only the admin login can open it. */
export default function AdminPage() {
  const [data, setData] = useState<Overview | null>(null)
  const [forbidden, setForbidden] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [preview, setPreview] = useState<Preview | null>(null)
  const [busy, setBusy] = useState<string | null>(null)
  const [toast, setToast] = useState<string | null>(null)
  const [month, setMonth] = useState('all')

  const load = useCallback(async () => {
    try {
      setData(await apiJson<Overview>('/api/admin/overview'))
      setError(null)
    } catch (e) {
      const m = (e as Error).message
      if (/only the boostaid admin/i.test(m)) setForbidden(true)
      else setError(m)
    }
  }, [])
  useEffect(() => { load() }, [load])

  const say = (m: string) => { setToast(m); setTimeout(() => setToast(null), 3500) }
  const post = async (url: string, body: unknown, method = 'POST') =>
    apiJson<Record<string, unknown>>(url, { method, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })

  const openInvoice = async (claimId: string) => {
    setBusy(claimId)
    setError(null)
    try { setPreview(await post('/api/admin/invoices', { claimId, preview: true }) as unknown as Preview) }
    catch (e) { setError((e as Error).message) }
    finally { setBusy(null) }
  }
  const sendInvoice = async () => {
    if (!preview) return
    setBusy('send')
    try {
      const claimId = data?.rows.find((r) => r.ref === preview.claimRef)?.claimId
      const r = (await post('/api/admin/invoices', { claimId })) as { invoice: { number: string; delivered: boolean } }
      setPreview(null)
      say(r.invoice.delivered ? `Invoice ${r.invoice.number} sent to ${preview.charity}` : `Invoice ${r.invoice.number} recorded, but NOT emailed: no email service is connected`)
      await load()
    } catch (e) { setPreview(null); setError((e as Error).message) }
    finally { setBusy(null) }
  }
  const markPaid = async (id: string, num: string) => {
    setBusy(id)
    try { await post(`/api/admin/invoices/${id}`, { action: 'paid' }, 'PATCH'); say(`${num} marked paid`); await load() }
    catch (e) { setError((e as Error).message) }
    finally { setBusy(null) }
  }
  const resend = async (id: string, num: string) => {
    setBusy(id)
    try { const r = (await post(`/api/admin/invoices/${id}`, { action: 'resend' }, 'PATCH')) as { delivered: boolean }; say(r.delivered ? `${num} sent again` : `${num} NOT emailed: no email service is connected`); await load() }
    catch (e) { setError((e as Error).message) }
    finally { setBusy(null) }
  }
  const sendAll = async () => {
    setBusy('all')
    setError(null)
    try {
      const r = (await post('/api/admin/invoices/send-all', {})) as { sent: number; failed: { ref: string; charity: string; error: string }[] }
      if (r.failed.length) setError(r.failed.map((f) => `${f.ref} (${f.charity}): ${f.error}`).join(' · '))
      say(`${r.sent} invoice${r.sent === 1 ? '' : 's'} sent${r.failed.length ? `, ${r.failed.length} could not be sent` : ''}`)
      await load()
    } catch (e) { setError((e as Error).message) }
    finally { setBusy(null) }
  }

  const months = useMemo(() => [...new Set((data?.archive ?? []).map((a) => monthOf(a.paidAt)).filter(Boolean))], [data])
  const shown = useMemo(() => (data?.archive ?? []).filter((a) => month === 'all' || monthOf(a.paidAt) === month), [data, month])

  const header = <PageHeader title="BoostAid admin" subtitle="What you see as the owner of BoostAid. Charities never see this page." />
  if (forbidden) return <>{header}<Card className="max-w-xl p-8"><div className="font-display text-xl font-semibold text-ink">This page is for the BoostAid admin only</div><p className="mt-2 text-muted">Your login doesn’t have access to it.</p></Card></>
  if (!data) {
    return (
      <>
        {header}
        {error ? <div role="alert" className="rounded-xl border border-bad-line bg-bad-soft px-4 py-3 text-sm font-medium text-bad">{error}</div> : (
          <div className="space-y-5" aria-busy="true"><div className="grid grid-cols-1 gap-5 sm:grid-cols-2 xl:grid-cols-4">{[0, 1, 2, 3].map((i) => <Skeleton key={i} className="h-28 rounded-2xl" />)}</div><Skeleton className="h-64 w-full rounded-2xl" /></div>
        )}
      </>
    )
  }

  const t = data.tiles
  const totals = data.byCharity.reduce((a, c) => ({ earned: a.earned + c.earned, owed: a.owed + c.owed }), { earned: 0, owed: 0 })
  const archiveTotal = (data.archive ?? []).reduce((s, a) => s + a.fee, 0)

  return (
    <div className="space-y-6">
      {header}
      {!data.emailConfigured && (
        <div className="rounded-xl bg-warn-soft px-4 py-3 text-sm text-warn">Email sending isn’t connected yet. Invoices are recorded and the PDF can be opened, but nothing is emailed to charities or to you until it is.</div>
      )}
      {error && <div role="alert" className="rounded-xl border border-bad-line bg-bad-soft px-4 py-3 text-sm font-medium text-bad">{error}</div>}

      <div className="grid grid-cols-1 gap-5 sm:grid-cols-2 xl:grid-cols-4">
        <StatTile label="Fees earned so far" value={gbp(t.earned)} caption="claims HMRC has paid" />
        <StatTile tone="good" label="Owed to you" value={gbp(t.owed)} caption={`${data.readyToInvoice} invoice${data.readyToInvoice === 1 ? '' : 's'} to send, plus sent and unpaid`} />
        <StatTile tone="bad" label="Overdue" value={gbp(t.overdue)} caption="invoices unpaid after 14 days" />
        <StatTile tone="warn" label="Fees waiting for HMRC" value={gbp(t.waiting)} caption="can’t be invoiced until received" />
      </div>

      <Card className="overflow-hidden">
        <div className="flex flex-wrap items-center gap-3 border-b border-line px-5 py-4">
          <h2 className="font-display text-lg font-semibold text-ink">Recovered Gift Aid by claim</h2>
          <span className="text-sm text-muted">Pending and received claims. Paid ones move to the Paid archive below.</span>
          <span className="flex-1" />
          <Button size="sm" variant="dark" disabled={!data.readyToInvoice || busy === 'all'} onClick={sendAll}>
            {busy === 'all' && <Loader2 className="h-3.5 w-3.5 animate-spin" />} Send all ready invoices ({data.readyToInvoice})
          </Button>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[900px] text-base">
            <thead>
              <tr className="border-b border-line">
                <th className={th}>Charity</th><th className={th}>Claim</th><th className={thr}>Recovered GA</th><th className={thr}>Fee 5%</th><th className={th}>Stage</th><th className={th}>Invoice</th>
              </tr>
            </thead>
            <tbody>
              {data.rows.map((r) => {
                const st = STAGE[r.stage]
                const inv = r.invoice
                return (
                  <tr key={r.claimId} className="border-b border-line-soft align-top last:border-0">
                    <td className={`${td} font-semibold text-ink`}>{r.charity}</td>
                    <td className={`${td} text-ink`}>{r.ref}</td>
                    <td className={`${tdr} text-ink`}>{gbp(r.recoveredPaid)}{r.short && <div className="text-xs text-muted">of {gbp(r.recoveredGiftAid)}, HMRC paid less</div>}</td>
                    <td className={`${tdr} font-bold text-ink`}>{gbp(r.fee)}</td>
                    <td className={td}>
                      <span className={`rounded-md px-2 py-1 text-sm font-semibold ${st.cls}`}>{r.stage === 'pending' && r.claimStatus === 'pending' ? 'Not submitted yet' : st.label}</span>
                      {r.stageAt && <div className="mt-1 text-xs text-muted">{r.stage === 'pending' ? 'submitted ' : r.stage === 'overdue' || r.stage === 'sent' ? 'sent ' : ''}{fmtDay(r.stageAt)}</div>}
                      {r.stuck && <div className="mt-1 text-xs font-semibold text-bad">With HMRC over 6 weeks: chase them</div>}
                    </td>
                    <td className={td}>
                      {r.stage === 'received' && <Button size="sm" variant="dark" disabled={busy === r.claimId} onClick={() => openInvoice(r.claimId)}>Send invoice now</Button>}
                      {inv && (r.stage === 'sent' || r.stage === 'overdue') && (
                        <div className="flex flex-wrap items-center gap-2">
                          <Button size="sm" variant="success" disabled={busy === inv.id} onClick={() => markPaid(inv.id, inv.number)}>Mark paid</Button>
                          <Button size="sm" variant="outline" disabled={busy === inv.id} onClick={() => resend(inv.id, inv.number)}>Resend</Button>
                          <a className="text-sm text-brand underline" href={`/api/admin/invoices/${inv.id}/pdf`} target="_blank" rel="noreferrer">PDF</a>
                          {!inv.delivered && <span className="text-xs font-semibold text-warn">Not emailed</span>}
                        </div>
                      )}
                      {r.stage === 'pending' && <Button size="sm" variant="outline" disabled title="You can invoice once the claim is marked received">Waiting for HMRC</Button>}
                    </td>
                  </tr>
                )
              })}
              {!data.rows.length && <tr><td colSpan={6} className="px-5 py-10 text-center text-muted">No claims with recovered Gift Aid yet.</td></tr>}
            </tbody>
          </table>
        </div>
      </Card>

      <section className="space-y-3">
        <h2 className="font-display text-xl font-semibold text-ink">Earned by charity</h2>
        <Card className="overflow-x-auto">
          <table className="w-full min-w-[480px] text-base">
            <thead><tr className="border-b border-line"><th className={th}>Charity</th><th className={thr}>Fees earned</th><th className={thr}>Owed to you</th></tr></thead>
            <tbody>
              {data.byCharity.map((c) => (
                <tr key={c.charity} className="border-b border-line-soft last:border-0">
                  <td className={`${td} text-ink`}>{c.charity}</td><td className={`${tdr} font-bold text-ink`}>{gbp(c.earned)}</td><td className={`${tdr} text-ink`}>{c.owed ? gbp(c.owed) : '–'}</td>
                </tr>
              ))}
              <tr className="bg-wash font-bold"><td className={td}>Total</td><td className={tdr}>{gbp(totals.earned)}</td><td className={tdr}>{gbp(totals.owed)}</td></tr>
            </tbody>
          </table>
        </Card>
      </section>

      <details className="rounded-2xl border border-line bg-white">
        <summary className="cursor-pointer list-none px-5 py-4 font-display text-lg font-semibold text-ink">
          Paid archive ({data.archive.length} invoice{data.archive.length === 1 ? '' : 's'}) · {gbp(archiveTotal)}
        </summary>
        <div className="space-y-3 px-5 pb-5">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <label className="flex items-center gap-2 text-sm text-muted">Show
              <select className="h-9 rounded-full border border-line-strong bg-white px-3 text-ink" value={month} onChange={(e) => setMonth(e.target.value)}>
                <option value="all">All paid invoices</option>
                {months.map((m) => <option key={m} value={m}>{monthName(m)}</option>)}
              </select>
            </label>
            <a href={`/api/admin/export${month === 'all' ? '' : `?month=${month}`}`}>
              <Button size="sm" variant="outline"><Download className="h-4 w-4" /> Export to Excel</Button>
            </a>
          </div>
          <div className="overflow-x-auto rounded-xl border border-line">
            <table className="w-full min-w-[720px] text-sm">
              <thead><tr className="border-b border-line"><th className={th}>Charity</th><th className={th}>Invoice</th><th className={th}>Claim</th><th className={th}>Sent</th><th className={th}>Paid</th><th className={thr}>Fee</th></tr></thead>
              <tbody>
                {shown.map((a) => (
                  <tr key={a.number} className="border-b border-line-soft last:border-0">
                    <td className={td}>{a.charity}</td><td className={td}>{a.number}</td><td className={td}>{a.ref}</td><td className={td}>{fmtDay(a.sentAt)}</td><td className={td}>{fmtDay(a.paidAt)}</td><td className={tdr}>{gbp(a.fee)}</td>
                  </tr>
                ))}
                {!shown.length && <tr><td colSpan={6} className="px-5 py-8 text-center text-muted">No paid invoices {month === 'all' ? 'yet' : 'in this month'}.</td></tr>}
                <tr className="bg-wash font-bold"><td className={td} colSpan={5}>Total{month === 'all' ? '' : ` for ${monthName(month)}`}</td><td className={tdr}>{gbp(shown.reduce((s, a) => s + a.fee, 0))}</td></tr>
              </tbody>
            </table>
          </div>
          <p className="text-xs text-muted">Paid invoices are kept for at least 6 years.</p>
        </div>
      </details>

      <section className="space-y-3">
        <h2 className="font-display text-xl font-semibold text-ink">Notifications</h2>
        <Card className="divide-y divide-line-soft">
          {data.notifications.map((n) => (
            <div key={n.id} className="flex gap-3 px-5 py-3.5">
              <span aria-hidden className="mt-2 h-2 w-2 shrink-0 rounded-full bg-brand" />
              <div><div className="text-ink">{n.text}</div><div className="text-sm text-muted">{fmtDay(n.at)}{data.emailConfigured ? ' · also sent to you by email' : ''}</div></div>
            </div>
          ))}
          {!data.notifications.length && <div className="px-5 py-8 text-center text-muted">Nothing yet.</div>}
        </Card>
      </section>

      <Dialog open={!!preview} onClose={() => busy !== 'send' && setPreview(null)} wide>
        {preview && (
          <div className="space-y-4">
            <h2 className="font-display text-xl font-semibold text-ink">Send invoice {preview.number}</h2>
            <section className="space-y-3 rounded-2xl border border-line p-5">
              <div><div className="font-display text-lg font-semibold text-ink">Invoice {preview.number}</div><div className="text-sm text-muted">Issued {fmtDay(preview.sentAt)} · payment due {fmtDay(preview.dueAt)}</div></div>
              <div className="grid gap-3 text-sm sm:grid-cols-3">
                <div><div className="font-semibold text-ink">From</div>BoostAid</div>
                <div><div className="font-semibold text-ink">To</div>{preview.charity}<br />{preview.to || <span className="text-bad">no contact email set</span>}</div>
                <div><div className="font-semibold text-ink">For</div>Gift Aid recovered by BoostAid in claim {preview.claimRef}</div>
              </div>
              <div className="space-y-1 border-t border-line pt-3 text-base">
                <div className="flex justify-between"><span>Recovered Gift Aid HMRC paid in {preview.claimRef}</span><b className="tabular-nums">{gbp(preview.recoveredPaid)}</b></div>
                <div className="flex justify-between font-bold text-brand"><span>Fee at 5%: total due</span><span className="tabular-nums">{gbp(preview.fee)}</span></div>
              </div>
              <div className="text-sm text-muted">Pay by bank transfer within 14 days, quoting {preview.number}. The list of recovered donations ({preview.donations.length}) is attached.</div>
            </section>
            {!preview.to && <div className="rounded-xl bg-bad-soft px-4 py-3 text-sm text-bad">Add a contact email for {preview.charity} in Settings first. The invoice is emailed there.</div>}
            {!data.emailConfigured && <div className="rounded-xl bg-warn-soft px-4 py-3 text-sm text-warn">Email sending isn’t connected: the invoice will be recorded but not emailed.</div>}
            <div className="flex justify-end gap-3">
              <Button variant="outline" onClick={() => setPreview(null)} disabled={busy === 'send'}>Cancel</Button>
              <Button variant="dark" onClick={sendInvoice} disabled={busy === 'send' || !preview.to || !preview.canSend}>
                {busy === 'send' && <Loader2 className="h-4 w-4 animate-spin" />} Email invoice to {preview.charity}
              </Button>
            </div>
          </div>
        )}
      </Dialog>

      {toast && <div role="status" className="fixed bottom-6 left-1/2 z-50 -translate-x-1/2 rounded-full bg-side px-5 py-3 text-sm font-semibold text-white shadow-lg">{toast}</div>}
    </div>
  )
}
