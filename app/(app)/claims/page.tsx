'use client'

import { useCallback, useEffect, useRef, useState } from 'react'
import Link from 'next/link'
import { Loader2, Undo2 } from 'lucide-react'
import { TableSkeleton } from '@/components/ui/skeleton'
import { Button } from '@/components/ui/button'
import { Card, PageHeader, Tabs, th, thr, td, tdr } from '@/components/ui/primitives'
import { Pagination } from '@/components/ui/pagination'
import SummaryStrip from '@/components/claims/summary-strip'
import FeeReport from '@/components/claims/fee-report'
import RecoveredBanner from '@/components/recovery/recovered-banner'
import SubmitDialog from '@/components/claims/submit-dialog'
import ReceivedDialog from '@/components/claims/received-dialog'
import NoCharityNotice from '@/components/charity/no-charity-notice'
import { useCharity } from '@/lib/store/charity'
import type { ClaimDto } from '@/lib/claims'
import { giftAidOn } from '@/lib/claims'
import {
  NEXT_STATUS,
  PREV_STATUS,
  STATUS_LABEL,
  StatusBadge,
  apiJson,
  fmtDay,
  fmtPlainDate,
  gbp,
  type Status,
} from '@/components/claims/shared'

type View = 'active' | 'archive' | 'all'

const TABS: { id: View; label: string }[] = [
  { id: 'active', label: 'Active' },
  { id: 'archive', label: 'Received (archive)' },
  { id: 'all', label: 'All' },
]

export default function ClaimsPage() {
  const { loaded, charities, selected, selectedId } = useCharity()
  const showCharity = selectedId === 'all'
  const [view, setView] = useState<View>('active')
  const [claims, setClaims] = useState<ClaimDto[] | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [busyId, setBusyId] = useState<string | null>(null)
  const [tick, setTick] = useState(0)
  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = useState(25)
  const [total, setTotal] = useState(0)
  const [fetching, setFetching] = useState(false)
  const [submitFor, setSubmitFor] = useState<ClaimDto | null>(null)
  const [receivedFor, setReceivedFor] = useState<ClaimDto | null>(null)

  const load = useCallback(async (v: View, p: number, size: number, charity: string) => {
    setFetching(true)
    try {
      const data = await apiJson<{ claims: ClaimDto[]; total: number }>(`/api/claims?view=${v}&page=${p}&pageSize=${size}&charity=${charity}`)
      // The last claim on this page may have just moved to another tab: step back a page
      if (data.claims.length === 0 && p > 1) {
        setPage(p - 1)
        return
      }
      setClaims(data.claims)
      setTotal(data.total)
      setError(null)
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setFetching(false)
    }
  }, [])

  const scopeRef = useRef(selectedId)
  useEffect(() => {
    if (!loaded || charities.length === 0) return
    // Switching charity starts again from page 1
    if (scopeRef.current !== selectedId) {
      scopeRef.current = selectedId
      setClaims(null)
      if (page !== 1) {
        setPage(1)
        return
      }
    }
    load(view, page, pageSize, selectedId)
  }, [loaded, charities.length, selectedId, view, page, pageSize, load])

  const changeView = (v: View) => {
    setClaims(null)
    setPage(1)
    setView(v)
  }

  const changeStatus = async (id: string, status: Status) => {
    setBusyId(id)
    try {
      await apiJson(`/api/claims/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status }),
      })
      await load(view, page, pageSize, selectedId)
      setTick((t) => t + 1)
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusyId(null)
    }
  }

  return (
    <>
      <PageHeader
        title="Claims"
        subtitle={`Every claim, from ready → submitted to HMRC → money received. ${selected ? `Showing ${selected.name}.` : 'Showing all charities.'}`}
      />

      {loaded && charities.length === 0 ? (
        <NoCharityNotice />
      ) : (
      <>
      <SummaryStrip refreshKey={tick} linked={false} />
      {selectedId !== 'all' && <RecoveredBanner charityId={selectedId} refreshKey={tick} />}

      <div className="mt-8 flex flex-wrap items-end justify-between gap-3">
        <div className="min-w-[260px] flex-1">
          <Tabs tabs={TABS.map((t) => ({ id: t.id, label: t.label }))} value={view} onChange={changeView} />
        </div>
        <a href={`/api/export?view=${view}&charity=${selectedId}`}>
          <Button variant="outline">Export list</Button>
        </a>
      </div>

      {error && <div className="mt-4 rounded-xl border border-bad/20 bg-bad-soft px-4 py-3 text-sm text-bad">{error}</div>}

      <Card className="mt-4 overflow-hidden">
        {claims === null && !error ? (
          <TableSkeleton
            columns={[
              { label: 'Status' },
              { label: 'Claim' },
              ...(showCharity ? [{ label: 'Charity' }] : []),
              { label: 'Date' },
              { label: 'Donations', right: true },
              { label: 'Donated', right: true },
              { label: 'Gift Aid (25%)', right: true },
              { label: 'Received' },
              { label: '', skeleton: true },
            ]}
            rows={Math.min(pageSize, 6)}
          />
        ) : claims && claims.length === 0 ? (
          <div className="py-16 text-center text-muted">
            {view === 'archive' ? 'No received claims yet' : 'No claims yet'}
            {selected ? ` for ${selected.name}` : ''}.{' '}
            {view !== 'archive' && (
              <Link href="/" className="text-brand font-medium hover:underline">
                Clean a spreadsheet
              </Link>
            )}
          </div>
        ) : (
          claims && (
            <div className={`overflow-x-auto transition-opacity ${fetching ? 'opacity-50 pointer-events-none' : ''}`} aria-busy={fetching}>
              <table className="w-full text-base">
                <thead>
                  <tr className="border-b border-line">
                    <th className={th}>Status</th>
                    <th className={th}>Claim</th>
                    {showCharity && <th className={th}>Charity</th>}
                    <th className={th} title="Earliest donation in this claim">Date</th>
                    <th className={thr}>Donations</th>
                    <th className={thr}>Donated</th>
                    <th className={thr}>Gift Aid (25%)</th>
                    <th className={th}>Received</th>
                    <th className={th} />
                  </tr>
                </thead>
                <tbody>
                  {claims.map((c) => {
                    const next = NEXT_STATUS[c.status]
                    const prev = PREV_STATUS[c.status]
                    const busy = busyId === c.id
                    return (
                      <tr key={c.id} className="border-b border-line-soft last:border-0 hover:bg-wash/60">
                        <td className={td}>
                          <StatusBadge status={c.status} />
                        </td>
                        <td className={td}>
                          <Link href={`/claims/${c.id}`} className="text-ink hover:text-brand whitespace-nowrap">
                            {c.ref}
                          </Link>
                          {c.hmrcRef && <div className="text-xs text-muted">HMRC: {c.hmrcRef}</div>}
                        </td>
                        {showCharity && <td className={`${td} text-ink`}>{c.charityName}</td>}
                        <td className={`${td} text-ink whitespace-nowrap`}>{fmtPlainDate(c.periodFrom ?? null)}</td>
                        <td className={`${tdr} text-ink`}>{c.donationCount.toLocaleString()}</td>
                        <td className={`${tdr} text-muted`}>{gbp(c.totalAmount)}</td>
                        <td className={`${tdr} text-ink`}>{gbp(c.giftAidAmount)}</td>
                        <td className={`${td} text-ink whitespace-nowrap`}>
                          {c.status === 'received' ? fmtDay(c.receivedAt) : '—'}
                          {c.status === 'received' && c.receivedAmount !== null && (
                            <div className={`text-xs ${c.receivedAmount === c.giftAidAmount ? 'text-muted' : 'text-warn'}`}>
                              {gbp(c.receivedAmount)}{c.receivedAmount !== c.giftAidAmount ? ` (${c.receivedAmount < c.giftAidAmount ? '-' : '+'}${gbp(Math.abs(c.receivedAmount - c.giftAidAmount))})` : ''}
                            </div>
                          )}
                          {c.status === 'submitted' && <div className="text-xs text-muted">sent {fmtDay(c.submittedAt)}</div>}
                        </td>
                        <td className={td}>
                          <div className="flex items-center justify-end gap-2">
                            {prev && (
                              <button
                                title={`Move back to ${STATUS_LABEL[prev].toLowerCase()}`}
                                disabled={busy}
                                onClick={() => {
                                  if (confirm(`Move ${c.ref} back to ${STATUS_LABEL[prev].toLowerCase()}?`)) changeStatus(c.id, prev)
                                }}
                                className="p-2 rounded-full text-muted hover:text-ink hover:bg-wash disabled:opacity-50 cursor-pointer"
                              >
                                <Undo2 className="w-4 h-4" />
                              </button>
                            )}
                            {next && (
                              <Button
                                size="sm"
                                variant={next.to === 'received' ? 'success' : 'dark'}
                                disabled={busy}
                                onClick={() => (next.to === 'submitted' ? setSubmitFor(c) : setReceivedFor(c))}
                              >
                                {busy && <Loader2 className="w-3.5 h-3.5 animate-spin" />}
                                {next.label}
                              </Button>
                            )}
                            <Link
                              href={`/claims/${c.id}`}
                              className="inline-flex h-9 items-center rounded-full border border-line-strong bg-white px-4 text-sm font-medium text-ink hover:bg-wash"
                            >
                              Open
                            </Link>
                          </div>
                        </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          )
        )}
        <Pagination
          page={page}
          pageSize={pageSize}
          total={total}
          onPage={setPage}
          onPageSize={(n) => {
            setPageSize(n)
            setPage(1)
          }}
        />
      </Card>

      {selectedId !== 'all' && <FeeReport charityId={selectedId} refreshKey={tick} />}

      <ReceivedDialog
        claim={receivedFor ? { id: receivedFor.id, ref: receivedFor.ref, giftAidAmount: receivedFor.giftAidAmount, recoveredGiftAid: giftAidOn(receivedFor.recoveredAmount) } : null}
        onClose={() => setReceivedFor(null)}
        onDone={() => {
          load(view, page, pageSize, selectedId)
          setTick((x) => x + 1)
        }}
      />
      <SubmitDialog
        claim={submitFor}
        onClose={() => setSubmitFor(null)}
        onDone={() => {
          load(view, page, pageSize, selectedId)
          setTick((t) => t + 1)
        }}
      />
      </>
      )}
    </>
  )
}
