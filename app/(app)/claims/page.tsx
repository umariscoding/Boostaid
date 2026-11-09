'use client'

import { useCallback, useEffect, useState } from 'react'
import Link from 'next/link'
import { Loader2, Undo2 } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Card, PageHeader, th, td } from '@/components/ui/primitives'
import { Pagination } from '@/components/ui/pagination'
import SummaryStrip from '@/components/claims/summary-strip'
import type { ClaimDto } from '@/lib/claims'
import {
  NEXT_STATUS,
  PREV_STATUS,
  StatusBadge,
  apiJson,
  fmtDay,
  fmtPeriod,
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
  const [view, setView] = useState<View>('active')
  const [claims, setClaims] = useState<ClaimDto[] | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [busyId, setBusyId] = useState<string | null>(null)
  const [tick, setTick] = useState(0)
  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = useState(25)
  const [total, setTotal] = useState(0)

  const load = useCallback(async (v: View, p: number, size: number) => {
    try {
      const data = await apiJson<{ claims: ClaimDto[]; total: number }>(`/api/claims?view=${v}&page=${p}&pageSize=${size}`)
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
    }
  }, [])

  useEffect(() => {
    load(view, page, pageSize)
  }, [view, page, pageSize, load])

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
      await load(view, page, pageSize)
      setTick((t) => t + 1)
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusyId(null)
    }
  }

  return (
    <>
      <PageHeader title="Claims" subtitle="Track each HMRC submission from pending to received." />

      <SummaryStrip refreshKey={tick} linked={false} />

      <div className="mt-8 flex flex-wrap items-center justify-between gap-3">
        <div className="flex gap-1 rounded-full bg-wash border border-line p-1">
          {TABS.map((t) => (
            <button
              key={t.id}
              onClick={() => changeView(t.id)}
              className={`px-4 h-9 rounded-full text-sm font-medium transition-colors cursor-pointer ${
                view === t.id ? 'bg-white text-ink shadow-sm border border-line' : 'text-muted hover:text-ink border border-transparent'
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>
        <a href={`/api/export?view=${view}`}>
          <Button variant="outline">Export list</Button>
        </a>
      </div>

      {error && <div className="mt-4 rounded-xl border border-bad/20 bg-bad-soft px-4 py-3 text-sm text-bad">{error}</div>}

      <Card className="mt-4 overflow-hidden">
        {claims === null && !error ? (
          <div className="flex items-center justify-center gap-3 py-16 text-muted">
            <Loader2 className="w-5 h-5 animate-spin" /> Loading claims...
          </div>
        ) : claims && claims.length === 0 ? (
          <div className="py-16 text-center text-muted">
            {view === 'archive' ? 'No received claims yet.' : 'No claims yet.'}{' '}
            {view !== 'archive' && (
              <Link href="/" className="text-brand font-medium hover:underline">
                Clean a spreadsheet
              </Link>
            )}
          </div>
        ) : (
          claims && (
            <div className="overflow-x-auto">
              <table className="w-full text-[15px]">
                <thead>
                  <tr className="border-b border-line">
                    <th className={th}>Status</th>
                    <th className={th}>Claim</th>
                    <th className={th}>Period</th>
                    <th className={`${th} text-right`}>Donations</th>
                    <th className={`${th} text-right`}>Donated</th>
                    <th className={`${th} text-right whitespace-nowrap`}>Gift Aid (25%)</th>
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
                      <tr key={c.id} className="border-b border-line last:border-0 hover:bg-wash/60">
                        <td className={td}>
                          <StatusBadge status={c.status} />
                        </td>
                        <td className={td}>
                          <Link href={`/claims/${c.id}`} className="text-ink hover:text-brand whitespace-nowrap">
                            {c.ref}
                          </Link>
                          {c.hmrcRef && <div className="text-xs text-muted">HMRC: {c.hmrcRef}</div>}
                        </td>
                        <td className={`${td} text-ink whitespace-nowrap`}>{fmtPeriod(c.periodFrom, c.periodTo)}</td>
                        <td className={`${td} text-right text-ink`}>{c.donationCount.toLocaleString()}</td>
                        <td className={`${td} text-right text-muted whitespace-nowrap`}>{gbp(c.totalAmount)}</td>
                        <td className={`${td} text-right text-ink whitespace-nowrap`}>{gbp(c.giftAidAmount)}</td>
                        <td className={`${td} text-ink whitespace-nowrap`}>
                          {c.status === 'received' ? fmtDay(c.receivedAt) : '—'}
                          {c.status === 'submitted' && <div className="text-xs text-muted">sent {fmtDay(c.submittedAt)}</div>}
                        </td>
                        <td className={td}>
                          <div className="flex items-center justify-end gap-2">
                            {prev && (
                              <button
                                title={`Move back to ${prev}`}
                                disabled={busy}
                                onClick={() => {
                                  if (confirm(`Move ${c.ref} back to ${prev}?`)) changeStatus(c.id, prev)
                                }}
                                className="p-2 rounded-full text-muted hover:text-ink hover:bg-wash disabled:opacity-50 cursor-pointer"
                              >
                                <Undo2 className="w-4 h-4" />
                              </button>
                            )}
                            {next && (
                              <Button
                                size="sm"
                                variant={next.to === 'received' ? 'success' : 'default'}
                                disabled={busy}
                                onClick={() => changeStatus(c.id, next.to)}
                              >
                                {busy && <Loader2 className="w-3.5 h-3.5 animate-spin" />}
                                {next.label}
                              </Button>
                            )}
                            <Link
                              href={`/claims/${c.id}`}
                              className="inline-flex h-9 items-center rounded-full border border-line bg-white px-4 text-sm font-medium text-ink hover:bg-wash"
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
    </>
  )
}
