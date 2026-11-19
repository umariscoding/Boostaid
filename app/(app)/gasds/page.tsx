'use client'

import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import Link from 'next/link'
import { Loader2 } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Card, PageHeader, inputClass, th, thr, td, tdr } from '@/components/ui/primitives'
import { Skeleton } from '@/components/ui/skeleton'
import { Pagination } from '@/components/ui/pagination'
import NoCharityNotice from '@/components/charity/no-charity-notice'
import ReceivedDialog from '@/components/claims/received-dialog'
import UploadDialog from '@/components/gasds/upload-dialog'
import CountsDialog from '@/components/gasds/counts-dialog'
import { apiJson, fmtDay, gbp } from '@/components/claims/shared'
import { useCharity } from '@/lib/store/charity'
import {
  analyse,
  daysLeft,
  deadlineOf,
  hasCount,
  longDate,
  showNo,
  taxYearLabel,
  topUpOf,
  ukDate,
  who,
  type GReceipt,
} from '@/lib/gasds'
import type { GasdsClaimDto } from '@/lib/gasds-db'

interface PageData {
  today: string
  years: number[]
  receipts: GReceipt[]
  claims: GasdsClaimDto[]
  claimedByYear: Record<number, number>
  giftAid: Record<number, { boostaid: number; elsewhere: number }>
}

const ICON = { in: '✅', wait: '⏳', out: '❌' } as const
const COLOUR = { in: 'text-good', wait: 'text-warn', out: 'text-muted' } as const

const stepCls = (s: string) => (s === 'now' ? 'border-2 border-brand' : s === 'later' ? 'border-line opacity-60' : 'border-line')
const badgeCls = (s: string) => (s === 'now' ? 'bg-brand text-white' : s === 'done' ? 'bg-good text-white' : 'bg-wash text-muted')
function Step({ n, state, title, children }: { n: number; state: string; title: string; children: React.ReactNode }) {
  return (
  <section className={`grid grid-cols-[40px_1fr] gap-x-4 rounded-2xl border bg-white p-5 ${stepCls(state)}`}>
    <span aria-hidden className={`row-span-2 grid h-10 w-10 place-items-center rounded-full font-display font-bold ${badgeCls(state)}`}>{state === 'done' ? '✓' : n}</span>
    <h3 className="pt-1.5 font-display text-lg font-semibold text-ink">{state === 'done' && <span className="sr-only">Done: </span>}{title}</h3>
    <div className="mt-2 space-y-3 text-base text-ink">{children}</div>
  </section>
  )
}

export default function SmallDonationsPage() {
  const { loaded, charities, selectedId } = useCharity()
  const charityId = selectedId !== 'all' ? selectedId : charities.length === 1 ? charities[0].id : ''

  const [data, setData] = useState<{ charityId: string; value: PageData } | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [year, setYear] = useState<number | null>(null)
  const [skipCounts, setSkipCounts] = useState(false)
  const [uploadOpen, setUploadOpen] = useState(false)
  const [countsOpen, setCountsOpen] = useState(false)
  const [receivedFor, setReceivedFor] = useState<GasdsClaimDto | null>(null)
  const [toast, setToast] = useState<string | null>(null)
  const [hmrcRef, setHmrcRef] = useState('')
  const [saving, setSaving] = useState(false)
  const [formError, setFormError] = useState<string | null>(null)
  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = useState(50)
  const [elsewhere, setElsewhere] = useState('')
  const receiptsRef = useRef<HTMLDetailsElement>(null)

  const load = useCallback(async (id: string) => {
    try {
      const value = await apiJson<PageData>(`/api/gasds?charity=${id}`)
      setData({ charityId: id, value })
      setError(null)
    } catch (e) {
      setError((e as Error).message)
    }
  }, [])

  useEffect(() => {
    if (loaded && charityId) {
      setSkipCounts(false)
      setHmrcRef('')
      setFormError(null)
      load(charityId)
    }
  }, [loaded, charityId, load])

  const say = (m: string) => {
    setToast(m)
    setTimeout(() => setToast(null), 3500)
  }

  const d = data && data.charityId === charityId ? data.value : null
  const y = d ? (year && d.years.includes(year) ? year : d.years[0]) : null

  const view = useMemo(() => {
    if (!d || y == null) return null
    const receipts = d.receipts.filter((r) => r.year === y)
    const claims = d.claims.filter((c) => c.taxYear === y)
    const ga = d.giftAid[y] ?? { boostaid: 0, elsewhere: 0 }
    const giftAid = Math.round((ga.boostaid + ga.elsewhere) * 100) / 100
    return { receipts, claims, ga, a: analyse(receipts, claims.map((c) => c.amount), giftAid, y) }
  }, [d, y])

  // The box for "Gift Aid claimed outside BoostAid" follows the chosen year
  useEffect(() => {
    if (view) setElsewhere(view.ga.elsewhere ? String(view.ga.elsewhere) : '')
  }, [view?.ga.elsewhere, y]) // eslint-disable-line react-hooks/exhaustive-deps

  if (loaded && charities.length === 0) {
    return (
      <>
        <PageHeader title="Small donations" />
        <NoCharityNotice />
      </>
    )
  }

  const header = (
    <div>
      <PageHeader title="Small donations" subtitle="HMRC adds 25% to cash and card donations of £30 or less. This is called GASDS." />
    </div>
  )

  if (loaded && !charityId) {
    return (
      <>
        {header}
        <Card className="p-8 max-w-2xl">
          <div className="font-display text-xl font-semibold text-ink">Choose a charity first</div>
          <p className="mt-2 text-muted">
            Small donations are claimed for one charity at a time. Pick the charity in the sidebar and this page will show its receipts and claims.
          </p>
        </Card>
      </>
    )
  }

  if (error && !d) {
    return (
      <>
        {header}
        <div role="alert" className="rounded-xl border border-bad-line bg-bad-soft px-4 py-3 text-sm font-medium text-bad">{error}</div>
      </>
    )
  }

  if (!d || !view || y == null) {
    return (
      <>
        {header}
        <div className="space-y-5" aria-busy="true">
          <Skeleton className="h-10 w-72 rounded-full" />
          <Skeleton className="h-36 w-full rounded-2xl" />
          <Skeleton className="h-28 w-full rounded-2xl" />
          <Skeleton className="h-28 w-full rounded-2xl" />
        </div>
      </>
    )
  }

  const { a, claims, receipts, ga } = view
  const full = a.room <= 0 && claims.length > 0
  const hasWork = a.amount > 0
  const cols = a.collections
  const counted = cols.filter((x) => hasCount(x.r)).length
  const s1 = hasWork || full ? 'done' : 'now'
  const s2 = !hasWork ? 'later' : counted === cols.length ? 'done' : 'now'
  const s3 = !hasWork ? 'later' : s2 === 'done' ? 'now' : 'later'
  const step2 = skipCounts && s2 === 'now' ? 'done' : s2
  const step3 = skipCounts && s3 === 'later' && hasWork ? 'now' : s3

  const inN = [...a.statuses.values()].filter((s) => s.group === 'in').length
  const waitN = [...a.statuses.values()].filter((s) => s.group === 'wait').length
  const outN = receipts.length - inN - waitN
  const left = daysLeft(y, d.today)

  const notes: { tone: 'warn' | 'bad' | 'info'; text: string }[] = []
  if (a.extra > 0) {
    notes.push({ tone: 'warn', text: `${gbp(a.waiting.reduce((s, r) => s + r.amount, 0))} is waiting to be banked. Once it is banked and uploaded again, HMRC can pay ${gbp(topUpOf(a.extra))} more.` })
  }
  if (a.matchingBites) {
    notes.push({ tone: 'bad', text: `This charity has only claimed Gift Aid on ${gbp(a.giftAid)} of donations for ${taxYearLabel(y)}. HMRC only pays on small donations up to 10 times that, so the amount is lower. Claim more Gift Aid to raise it.` })
  }
  if (hasWork && left <= 180) {
    notes.push({ tone: 'bad', text: `Claim by ${longDate(deadlineOf(y))} (${left} days left), or this money is lost.` })
  }
  if (a.maybeGiftAid.length) {
    const n = a.maybeGiftAid.length
    notes.push({
      tone: 'info',
      text: `${n} donation${n > 1 ? 's' : ''} over £30 ${n > 1 ? 'are' : 'is'} from people who gave Gift Aid on another receipt. Check if ${n > 1 ? 'they' : 'it'} can go in your Gift Aid claim instead: ${a.maybeGiftAid.slice(0, 6).map((r) => `${r.donor} (${gbp(r.amount)})`).join(', ')}${n > 6 ? ` and ${n - 6} more` : ''}.`,
    })
  }

  const sorted = [...receipts].sort((p1, p2) => (p1.date < p2.date ? 1 : p1.date > p2.date ? -1 : p1.no < p2.no ? 1 : -1))
  const pageCount = Math.max(1, Math.ceil(sorted.length / pageSize))
  const pageNo = Math.min(page, pageCount)
  const shown = sorted.slice((pageNo - 1) * pageSize, pageNo * pageSize)

  const copy = async () => {
    try {
      await navigator.clipboard.writeText(a.amount.toFixed(2))
      say(`Copied ${gbp(a.amount)}`)
    } catch {
      say(`Amount: ${gbp(a.amount)}`)
    }
  }

  const saveClaim = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!hmrcRef.trim()) {
      setFormError('Enter the reference HMRC gave you first.')
      return
    }
    setSaving(true)
    setFormError(null)
    try {
      await apiJson('/api/gasds/claims', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ charityId, taxYear: y, hmrcRef: hmrcRef.trim(), amount: a.amount }),
      })
      setHmrcRef('')
      setSkipCounts(false)
      await load(charityId)
      say(`Claim saved. ${gbp(a.topUp)} is on its way from HMRC.`)
    } catch (err) {
      setFormError((err as Error).message)
      load(charityId)
    } finally {
      setSaving(false)
    }
  }

  const saveElsewhere = async () => {
    const value = Number(elsewhere || 0)
    if (!Number.isFinite(value) || value < 0 || value === ga.elsewhere) return
    try {
      await apiJson('/api/gasds/years', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ charityId, taxYear: y, giftAidElsewhere: value }),
      })
      await load(charityId)
    } catch (err) {
      say((err as Error).message)
    }
  }

  const noteCls = { warn: 'bg-warn-soft text-warn', bad: 'bg-bad-soft text-bad', info: 'bg-info-soft text-info' }

  return (
    <>
      {header}
      <div className="mt-6 space-y-5">
      <div className="flex flex-wrap items-center gap-2" role="group" aria-label="Tax year">
        <span className="text-sm text-muted">Tax year</span>
        {d.years.map((yr) => (
          <button
            key={yr}
            type="button"
            aria-pressed={yr === y}
            onClick={() => { setYear(yr); setPage(1) }}
            className={`h-9 cursor-pointer rounded-full border px-4 text-sm font-semibold ${yr === y ? 'border-side bg-side text-white' : 'border-line-strong bg-white text-ink hover:bg-wash'}`}
          >
            {taxYearLabel(yr)}
          </button>
        ))}
      </div>

      {/* The headline: one sentence */}
      <section aria-live="polite" className={`rounded-2xl border px-6 py-5 ${full || hasWork ? 'border-good-line bg-good-soft text-good' : 'border-line bg-white text-ink'}`}>
        {full ? (
          <>
            <div>✓ Limit reached for {taxYearLabel(y)}</div>
            <div className="font-display text-4xl font-bold tabular-nums">{gbp(topUpOf(a.claimed))} claimed</div>
            <div>That is the most HMRC pays for one year. Anything else from this year can’t be added.</div>
          </>
        ) : hasWork ? (
          <>
            <div>HMRC can pay you</div>
            <div className="font-display text-4xl font-bold tabular-nums">{gbp(a.topUp)}</div>
            <div>{claims.length ? `on top of the ${gbp(topUpOf(a.claimed))} already claimed this year.` : 'Follow the 3 steps below to claim it.'}</div>
          </>
        ) : claims.length ? (
          <>
            <div>All claimed so far for {taxYearLabel(y)}</div>
            <div className="font-display text-4xl font-bold tabular-nums">{gbp(topUpOf(a.claimed))} claimed</div>
            <div>Upload more receipts to claim the rest. Up to {gbp(topUpOf(a.room))} more is possible this year.</div>
          </>
        ) : (
          <>
            <div>Nothing to claim yet for {taxYearLabel(y)}</div>
            <div>Start with step 1: upload your receipts.</div>
          </>
        )}
      </section>

      {!full && (
        <div className="space-y-3">
          <Step n={1} state={s1} title="Upload your receipts">
            {receipts.length ? (
              <>
                <div className="flex flex-wrap gap-2 text-sm font-semibold">
                  <span className="rounded-full bg-good-soft px-3 py-1.5 text-good">✅ {inN} included</span>
                  <span className="rounded-full bg-warn-soft px-3 py-1.5 text-warn">⏳ {waitN} waiting for the bank</span>
                  <span className="rounded-full bg-wash px-3 py-1.5 text-muted">❌ {outN} not included</span>
                </div>
                <div className="flex flex-wrap items-center gap-4">
                  <Button size="sm" variant="outline" onClick={() => setUploadOpen(true)}>Upload more</Button>
                  <button type="button" className="cursor-pointer text-sm font-medium text-brand underline" onClick={() => { if (receiptsRef.current) { receiptsRef.current.open = true; receiptsRef.current.scrollIntoView({ block: 'start' }) } }}>See why</button>
                </div>
              </>
            ) : (
              <>
                <span>Upload your receipt book export, or the BoostAid template.</span>
                <div><Button variant="dark" onClick={() => setUploadOpen(true)}>Upload receipts</Button></div>
              </>
            )}
          </Step>

          <Step n={2} state={step2} title="Add the counts">
            {!hasWork ? (
              <span className="text-sm text-muted">After step 1.</span>
            ) : cols.length ? (
              <>
                <span>
                  {counted === cols.length
                    ? `All ${cols.length} collections in this claim have their counts.`
                    : `${cols.length} collection${cols.length > 1 ? 's are' : ' is'} in this claim. Type in the notes and coins counted for each one. This is your proof for HMRC.`}
                </span>
                <div className="flex flex-wrap items-center gap-4">
                  <Button size="sm" variant={counted === cols.length ? 'outline' : 'dark'} onClick={() => setCountsOpen(true)}>
                    {counted === cols.length ? 'Check the counts' : `Add the counts (${counted} of ${cols.length} done)`}
                  </Button>
                  {counted < cols.length && (
                    <button type="button" className="cursor-pointer text-sm font-medium text-brand underline" onClick={() => setSkipCounts(true)}>Skip for now</button>
                  )}
                </div>
              </>
            ) : (
              <span>No collections in this claim, so there is nothing to count.</span>
            )}
          </Step>

          <Step n={3} state={step3} title="Send to HMRC">
            {!hasWork ? (
              <span className="text-sm text-muted">After step 2.</span>
            ) : (
              <>
                <span>On HMRC Charities Online, type this into the small donations (GASDS) box:</span>
                <div className="flex flex-wrap items-center gap-3 rounded-xl bg-wash px-4 py-3">
                  <b className="font-display text-2xl tabular-nums">{gbp(a.amount)}</b>
                  <Button size="sm" variant="outline" onClick={copy}>Copy</Button>
                  <span className="text-sm text-muted">HMRC will pay you {gbp(a.topUp)}. Deadline {longDate(deadlineOf(y))}.</span>
                </div>
                <form onSubmit={saveClaim} className="flex flex-wrap items-end gap-3">
                  <label className="flex flex-col gap-1.5 text-sm font-semibold text-ink">
                    Done? Enter the reference HMRC gave you
                    <input className={`${inputClass} min-w-[240px]`} value={hmrcRef} onChange={(e) => setHmrcRef(e.target.value)} maxLength={100} autoComplete="off" required />
                  </label>
                  <Button type="submit" variant="success" disabled={saving || !hmrcRef.trim()}>
                    {saving && <Loader2 className="h-4 w-4 animate-spin" />} Save claim
                  </Button>
                </form>
                {formError && <div role="alert" className="rounded-xl border border-bad-line bg-bad-soft px-4 py-3 text-sm font-medium text-bad">{formError}</div>}
              </>
            )}
          </Step>
        </div>
      )}

      {full && (
        <div className="flex flex-wrap items-center gap-3">
          <Button size="sm" variant="outline" onClick={() => setUploadOpen(true)}>Upload more receipts</Button>
          <span className="text-sm text-muted">They’ll be saved as records. Receipts from another tax year can still be claimed in that year.</span>
        </div>
      )}

      {notes.length > 0 && (
        <div className="space-y-2">{notes.map((n, i) => <div key={i} className={`rounded-xl px-4 py-3 text-sm ${noteCls[n.tone]}`}>{n.text}</div>)}</div>
      )}

      {/* Claims sent to HMRC */}
      {claims.length > 0 && (
        <section className="space-y-3">
          <h2 className="font-display text-xl font-semibold text-ink">Claims sent to HMRC</h2>
          <Card className="overflow-x-auto">
            <table className="w-full min-w-[720px] text-base">
              <thead>
                <tr className="border-b border-line">
                  <th className={th}>Sent</th><th className={th}>HMRC ref</th><th className={thr}>Small donations</th><th className={thr}>HMRC pays</th><th className={th}>Status</th><th className={th}><span className="sr-only">Actions</span></th>
                </tr>
              </thead>
              <tbody>
                {claims.map((c) => (
                  <tr key={c.id} className="border-b border-line-soft last:border-0">
                    <td className={`${td} whitespace-nowrap`}>{fmtDay(c.sentAt)}{c.sentBy && <div className="text-xs text-muted">{c.sentBy}</div>}</td>
                    <td className={`${td} font-semibold`}>{c.hmrcRef}</td>
                    <td className={tdr}>{gbp(c.amount)}</td>
                    <td className={`${tdr} font-bold`}>{gbp(topUpOf(c.amount))}</td>
                    <td className={td}>
                      {c.status === 'received' ? (
                        <span>
                          <span className="font-semibold text-good">Received</span>
                          {c.receivedAmount != null && (
                            <div className={`text-xs ${c.receivedAmount === topUpOf(c.amount) ? 'text-muted' : 'text-warn'}`}>
                              {gbp(c.receivedAmount)}{c.receivedAmount !== topUpOf(c.amount) && ` (expected ${gbp(topUpOf(c.amount))})`}
                            </div>
                          )}
                        </span>
                      ) : (
                        <span className="font-semibold text-warn">Waiting for HMRC</span>
                      )}
                    </td>
                    <td className={`${td} text-right`}>
                      <div className="flex items-center justify-end gap-2">
                        {c.status === 'submitted' && <Button size="sm" variant="success" onClick={() => setReceivedFor(c)}>Mark received</Button>}
                        <Link href={`/gasds/claims/${c.id}`} className="inline-flex h-9 items-center rounded-full border border-line-strong bg-white px-4 text-sm font-medium text-ink hover:bg-wash">View evidence</Link>
                      </div>
                    </td>
                  </tr>
                ))}
                {claims.length > 1 && (
                  <tr className="bg-wash font-bold">
                    <td className={td} colSpan={2}>Total</td>
                    <td className={tdr}>{gbp(a.claimed)}</td>
                    <td className={tdr}>{gbp(topUpOf(a.claimed))}</td>
                    <td className={td} colSpan={2} />
                  </tr>
                )}
              </tbody>
            </table>
          </Card>
        </section>
      )}

      {/* All receipts, folded away */}
      <details ref={receiptsRef} className="rounded-2xl border border-line bg-white">
        <summary className="cursor-pointer list-none px-5 py-4 font-display text-lg font-semibold text-ink">All receipts for {taxYearLabel(y)} ({receipts.length})</summary>
        <div className="space-y-3 px-5 pb-5">
          <div className="overflow-x-auto rounded-xl border border-line">
            <table className="w-full min-w-[640px] text-sm">
              <thead><tr className="border-b border-line text-left text-muted"><th className={th}>Date</th><th className={th}>Receipt</th><th className={th}>From</th><th className={thr}>Amount</th><th className={th}>Included?</th></tr></thead>
              <tbody>
                {shown.map((r) => {
                  const s = a.statuses.get(r.no)!
                  return (
                    <tr key={r.no} className="border-b border-line-soft last:border-0">
                      <td className={`${td} whitespace-nowrap`}>{ukDate(r.date)}</td>
                      <td className={td}>{showNo(r.no)}</td>
                      <td className={td}>{who(r)}{!r.gift && <div className="text-xs text-muted">collection</div>}</td>
                      <td className={tdr}>{gbp(r.amount)}</td>
                      <td className={`${td} font-semibold ${COLOUR[s.group]}`}>{ICON[s.group]} {s.label}{s.reason && <div className="text-xs font-normal text-muted">{s.reason}</div>}</td>
                    </tr>
                  )
                })}
                {!receipts.length && <tr><td colSpan={5} className="px-5 py-6 text-center text-muted">No receipts uploaded for {taxYearLabel(y)} yet.</td></tr>}
              </tbody>
            </table>
          </div>
          <Pagination page={pageNo} pageSize={pageSize} total={sorted.length} onPage={setPage} onPageSize={(n) => { setPageSize(n); setPage(1) }} sizes={[50, 100]} />
        </div>
      </details>

      {/* How it is worked out, folded away */}
      <details className="rounded-2xl border border-line bg-white">
        <summary className="cursor-pointer list-none px-5 py-4 font-display text-lg font-semibold text-ink">How it is worked out</summary>
        <div className="space-y-4 px-5 pb-5 text-sm">
          <dl className="grid grid-cols-[1fr_auto] items-center gap-x-6 gap-y-1.5">
            <dt>Receipts that count, not yet claimed</dt><dd className="text-right font-semibold tabular-nums">{gbp(a.openAmount)}</dd>
            <dt>Already claimed this year</dt><dd className="text-right font-semibold tabular-nums">{gbp(a.claimed)}</dd>
            <dt>Most HMRC pays on in a year</dt><dd className="text-right font-semibold tabular-nums">£8,000.00</dd>
            <dt>Gift Aid limit: 10 × the Gift Aid donations for this year</dt><dd className="text-right font-semibold tabular-nums">{gbp(a.matchLimit)}</dd>
            <dt><b>Type into HMRC</b></dt><dd className="text-right font-bold tabular-nums">{gbp(a.amount)}</dd>
            <dt><b>HMRC pays 25%</b></dt><dd className="text-right font-bold tabular-nums">{gbp(a.topUp)}</dd>
          </dl>
          <p className="text-muted">Included: cash or card, banked, no Gift Aid, and £30 or less from one person (a collection can be any size). £50 notes are taken off. The oldest receipts are used first.</p>
          <dl className="grid grid-cols-[1fr_auto] items-center gap-x-6 gap-y-2">
            <dt>Gift Aid donations claimed in BoostAid for {taxYearLabel(y)}</dt><dd className="text-right font-semibold tabular-nums">{gbp(ga.boostaid)}</dd>
            <dt><label htmlFor="elsewhere">Gift Aid claimed outside BoostAid (e.g. by an agent)</label></dt>
            <dd>
              <input id="elsewhere" type="number" min="0" step="0.01" inputMode="decimal" placeholder="0" value={elsewhere} onChange={(e) => setElsewhere(e.target.value)} onBlur={saveElsewhere} className={`${inputClass} w-36 text-right`} />
            </dd>
          </dl>
        </div>
      </details>

      </div>

      <UploadDialog
        open={uploadOpen}
        ctx={{
          charityId,
          today: d.today,
          stored: d.receipts,
          claimedByYear: d.claimedByYear,
          giftAidByYear: Object.fromEntries(Object.entries(d.giftAid).map(([k, v]) => [k, v.boostaid + v.elsewhere])),
        }}
        onClose={() => setUploadOpen(false)}
        onDone={(m, best) => {
          if (best) { setYear(best); setPage(1) }
          setSkipCounts(false)
          load(charityId)
          say(m)
        }}
      />
      <CountsDialog open={countsOpen} receipts={cols.map((x) => x.r)} onClose={() => setCountsOpen(false)} onChanged={() => load(charityId)} />
      <ReceivedDialog
        claim={receivedFor ? { id: receivedFor.id, ref: receivedFor.hmrcRef, giftAidAmount: topUpOf(receivedFor.amount) } : null}
        endpoint={(id) => `/api/gasds/claims/${id}`}
        onClose={() => setReceivedFor(null)}
        onDone={() => load(charityId)}
      />
      {toast && <div role="status" className="fixed bottom-6 left-1/2 z-50 -translate-x-1/2 rounded-full bg-side px-5 py-3 text-sm font-semibold text-white shadow-lg">{toast}</div>}
    </>
  )
}
