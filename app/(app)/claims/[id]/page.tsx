'use client'

import { use, useCallback, useEffect, useMemo, useState } from 'react'
import Link from 'next/link'
import { useRouter } from 'next/navigation'
import { ArrowLeft, Download, Loader2, Trash2, Undo2 } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Card, inputClass, labelClass, th, thr, td, tdr } from '@/components/ui/primitives'
import { Pagination } from '@/components/ui/pagination'
import type { ClaimDto } from '@/lib/claims'
import {
  NEXT_STATUS,
  PREV_STATUS,
  StatusBadge,
  apiJson,
  fmtDateTime,
  fmtPeriod,
  fmtPlainDate,
  gbp,
  type Status,
} from '@/components/claims/shared'

interface DonationRow {
  id: string
  title: string
  firstName: string
  lastName: string
  address: string
  postcode: string
  donationDate: string | null
  amount: number
}
interface EventRow {
  id: string
  status: Status
  at: string
  note: string | null
}

export default function ClaimDetailPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params)
  const router = useRouter()
  const [claim, setClaim] = useState<ClaimDto | null>(null)
  const [donations, setDonations] = useState<DonationRow[]>([])
  const [events, setEvents] = useState<EventRow[]>([])
  const [error, setError] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)
  const [hmrcRef, setHmrcRef] = useState('')
  const [notes, setNotes] = useState('')
  const [saved, setSaved] = useState(false)
  const [filter, setFilter] = useState('')
  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = useState(25)

  const load = useCallback(async () => {
    try {
      const data = await apiJson<{ claim: ClaimDto; donations: DonationRow[]; events: EventRow[] }>(`/api/claims/${id}`)
      setClaim(data.claim)
      setDonations(data.donations)
      setEvents(data.events)
      setHmrcRef(data.claim.hmrcRef ?? '')
      setNotes(data.claim.notes ?? '')
      setError(null)
    } catch (e) {
      setError((e as Error).message)
    }
  }, [id])

  useEffect(() => {
    load()
  }, [load])

  const patch = async (body: Record<string, unknown>) => {
    setBusy(true)
    try {
      await apiJson(`/api/claims/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      })
      await load()
      return true
    } catch (e) {
      setError((e as Error).message)
      return false
    } finally {
      setBusy(false)
    }
  }

  const saveDetails = async () => {
    if (await patch({ hmrcRef, notes })) {
      setSaved(true)
      setTimeout(() => setSaved(false), 2000)
    }
  }

  const remove = async () => {
    if (!claim || !confirm(`Delete ${claim.ref} and its ${claim.donationCount} donations? This cannot be undone.`)) return
    setBusy(true)
    try {
      await apiJson(`/api/claims/${id}`, { method: 'DELETE' })
      router.push('/claims')
    } catch (e) {
      setError((e as Error).message)
      setBusy(false)
    }
  }

  const shown = useMemo(() => {
    const q = filter.trim().toLowerCase()
    if (!q) return donations
    return donations.filter((d) => `${d.firstName} ${d.lastName} ${d.address} ${d.postcode}`.toLowerCase().includes(q))
  }, [donations, filter])

  const next = claim ? NEXT_STATUS[claim.status] : null
  const prev = claim ? PREV_STATUS[claim.status] : null

  return (
    <>
      <Link href="/claims" className="inline-flex items-center gap-2 text-sm font-medium text-muted hover:text-ink mb-5">
        <ArrowLeft className="w-4 h-4" /> All claims
      </Link>

      {error && <div className="mb-4 rounded-xl border border-bad/20 bg-bad-soft px-4 py-3 text-sm text-bad">{error}</div>}
      {!claim && !error && (
        <div className="flex items-center gap-3 py-16 justify-center text-muted">
          <Loader2 className="w-5 h-5 animate-spin" /> Loading...
        </div>
      )}

      {claim && (
        <div className="space-y-6">
          <div className="flex flex-wrap items-start justify-between gap-4">
            <div>
              <div className="flex items-center gap-4">
                <h1 className="font-display text-5xl font-bold tracking-tight text-ink">{claim.ref}</h1>
                <StatusBadge status={claim.status} />
              </div>
              <p className="text-muted mt-2">
                {claim.donationCount.toLocaleString()} donations · {fmtPeriod(claim.periodFrom, claim.periodTo)} · from{' '}
                {claim.sourceFile} · created {fmtDateTime(claim.createdAt)}
              </p>
            </div>
            <div className="flex gap-8 text-right">
              <div>
                <div className="text-sm text-muted">Donated</div>
                <div className="font-display text-3xl font-bold text-ink">{gbp(claim.totalAmount)}</div>
              </div>
              <div>
                <div className="text-sm text-muted">Gift Aid (25%)</div>
                <div className="font-display text-3xl font-bold text-good">{gbp(claim.giftAidAmount)}</div>
              </div>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {next && (
              <Button variant={next.to === 'received' ? 'success' : 'default'} disabled={busy} onClick={() => patch({ status: next.to })}>
                {next.label}
              </Button>
            )}
            {prev && (
              <Button variant="outline" disabled={busy} onClick={() => confirm(`Move ${claim.ref} back to ${prev}?`) && patch({ status: prev })}>
                <Undo2 className="w-4 h-4" /> Move back to {prev}
              </Button>
            )}
            <a href={`/api/claims/${id}/export`}>
              <Button variant="outline">
                <Download className="w-4 h-4" /> Download audit file
              </Button>
            </a>
            {claim.status === 'pending' && (
              <Button variant="danger" disabled={busy} onClick={remove} className="ml-auto">
                <Trash2 className="w-4 h-4" /> Delete
              </Button>
            )}
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <Card className="p-6 lg:col-span-2 space-y-4">
              <h3 className="font-display text-xl font-semibold text-ink">Details</h3>
              <div>
                <label className={labelClass} htmlFor="hmrc">HMRC reference</label>
                <input id="hmrc" className={inputClass} value={hmrcRef} onChange={(e) => setHmrcRef(e.target.value)} placeholder="Optional - the reference HMRC gives on submission" maxLength={100} />
              </div>
              <div>
                <label className={labelClass} htmlFor="notes">Notes</label>
                <textarea id="notes" rows={3} className={`${inputClass} h-auto py-3`} value={notes} onChange={(e) => setNotes(e.target.value)} maxLength={2000} />
              </div>
              <Button size="sm" disabled={busy} onClick={saveDetails}>
                {saved ? 'Saved' : 'Save details'}
              </Button>
            </Card>

            <Card className="p-6">
              <h3 className="font-display text-xl font-semibold text-ink mb-4">History</h3>
              <ol className="space-y-4">
                {[...events].reverse().map((e) => (
                  <li key={e.id} className="flex gap-3">
                    <span className="mt-2 h-2 w-2 flex-shrink-0 rounded-full bg-brand" />
                    <div>
                      <StatusBadge status={e.status} />
                      <div className="text-sm text-muted mt-1">{fmtDateTime(e.at)}</div>
                      {e.note && <div className="text-sm text-muted">{e.note}</div>}
                    </div>
                  </li>
                ))}
              </ol>
            </Card>
          </div>

          <Card className="overflow-hidden">
            <div className="flex flex-wrap items-center justify-between gap-3 p-5 border-b border-line">
              <h3 className="font-display text-xl font-semibold text-ink">Donations in this claim</h3>
              <input className={`${inputClass} max-w-xs`} placeholder="Filter by name, address, postcode" value={filter} onChange={(e) => { setFilter(e.target.value); setPage(1) }} />
            </div>
            <div className="overflow-auto">
              <table className="w-full text-[15px]">
                <thead className="sticky top-0 bg-white">
                  <tr className="border-b border-line">
                    <th className={th}>Donor</th>
                    <th className={th}>House</th>
                    <th className={th}>Postcode</th>
                    <th className={th}>Date</th>
                    <th className={thr}>Amount</th>
                  </tr>
                </thead>
                <tbody>
                  {shown.slice((page - 1) * pageSize, page * pageSize).map((d) => (
                    <tr key={d.id} className="border-b border-line last:border-0">
                      <td className={`${td} py-3 text-ink`}>{[d.title, d.firstName, d.lastName].filter(Boolean).join(' ')}</td>
                      <td className={`${td} py-3 text-ink`}>{d.address}</td>
                      <td className={`${td} py-3 text-ink`}>{d.postcode}</td>
                      <td className={`${td} py-3 text-ink whitespace-nowrap`}>{fmtPlainDate(d.donationDate)}</td>
                      <td className={`${tdr} py-3 text-ink`}>{gbp(d.amount)}</td>
                    </tr>
                  ))}
                  {shown.length === 0 && (
                    <tr>
                      <td colSpan={5} className="px-4 py-10 text-center text-muted">No matching donations</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            <Pagination
              page={page}
              pageSize={pageSize}
              total={shown.length}
              onPage={setPage}
              onPageSize={(n) => {
                setPageSize(n)
                setPage(1)
              }}
            />
          </Card>
        </div>
      )}
    </>
  )
}
