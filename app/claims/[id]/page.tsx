'use client'

import { use, useCallback, useEffect, useMemo, useState } from 'react'
import Link from 'next/link'
import { useRouter } from 'next/navigation'
import { ArrowLeft, Download, Loader2, Trash2, Undo2 } from 'lucide-react'
import AppHeader from '@/components/app-header'
import { Button } from '@/components/ui/button'
import type { ClaimDto } from '@/lib/claims'
import {
  NEXT_STATUS,
  PREV_STATUS,
  StatusBadge,
  apiJson,
  cardClass,
  fmtDateTime,
  fmtPlainDate,
  gbp,
  inputClass,
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
    return donations.filter((d) =>
      `${d.firstName} ${d.lastName} ${d.address} ${d.postcode}`.toLowerCase().includes(q)
    )
  }, [donations, filter])

  const next = claim ? NEXT_STATUS[claim.status] : null
  const prev = claim ? PREV_STATUS[claim.status] : null

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900">
      <AppHeader />
      <main className="max-w-7xl mx-auto px-6 py-6 space-y-6">
        <Link href="/claims" className="inline-flex items-center gap-2 text-sm text-slate-400 hover:text-white">
          <ArrowLeft className="w-4 h-4" /> All claims
        </Link>

        {error && <div className="rounded-lg border border-red-500/30 bg-red-500/10 px-4 py-3 text-sm text-red-200">{error}</div>}
        {!claim && !error && (
          <div className="flex items-center gap-3 py-16 justify-center text-slate-400">
            <Loader2 className="w-5 h-5 animate-spin" /> Loading...
          </div>
        )}

        {claim && (
          <>
            <div className={`${cardClass} p-6`}>
              <div className="flex flex-wrap items-start justify-between gap-4">
                <div>
                  <div className="flex items-center gap-3">
                    <h2 className="text-2xl font-semibold font-mono text-white">{claim.ref}</h2>
                    <StatusBadge status={claim.status} />
                  </div>
                  <p className="text-sm text-slate-400 mt-1">
                    {claim.donationCount.toLocaleString()} donations · from {claim.sourceFile} · created {fmtDateTime(claim.createdAt)}
                  </p>
                </div>
                <div className="text-right">
                  <div className="text-xs uppercase tracking-wide text-slate-400">Gift Aid amount</div>
                  <div className="text-3xl font-bold text-white">{gbp(claim.totalAmount)}</div>
                </div>
              </div>

              <div className="mt-5 flex flex-wrap items-center gap-2">
                {next && (
                  <Button
                    disabled={busy}
                    onClick={() => patch({ status: next.to })}
                    className={
                      next.to === 'received'
                        ? 'bg-emerald-600 hover:bg-emerald-700 text-white'
                        : 'bg-gradient-to-r from-violet-500 to-purple-600 hover:from-violet-600 hover:to-purple-700 text-white'
                    }
                  >
                    {next.label}
                  </Button>
                )}
                {prev && (
                  <Button
                    variant="outline"
                    disabled={busy}
                    onClick={() => confirm(`Move ${claim.ref} back to ${prev}?`) && patch({ status: prev })}
                    className="border-slate-600 bg-slate-700/50 text-slate-200 hover:bg-slate-600 hover:text-white"
                  >
                    <Undo2 className="w-4 h-4" /> Move back to {prev}
                  </Button>
                )}
                <a href={`/api/claims/${id}/export`}>
                  <Button variant="outline" className="border-slate-600 bg-slate-700/50 text-slate-200 hover:bg-slate-600 hover:text-white">
                    <Download className="w-4 h-4" /> Download audit file
                  </Button>
                </a>
                {claim.status === 'pending' && (
                  <Button variant="outline" disabled={busy} onClick={remove} className="ml-auto border-red-500/40 bg-red-500/10 text-red-200 hover:bg-red-500/20">
                    <Trash2 className="w-4 h-4" /> Delete
                  </Button>
                )}
              </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              <div className={`${cardClass} p-5 lg:col-span-2 space-y-4`}>
                <h3 className="font-semibold text-white">Details</h3>
                <div>
                  <label className="block text-xs uppercase tracking-wide text-slate-400 mb-1.5" htmlFor="hmrc">
                    HMRC reference
                  </label>
                  <input id="hmrc" className={inputClass} value={hmrcRef} onChange={(e) => setHmrcRef(e.target.value)} placeholder="Optional - e.g. the reference HMRC gives on submission" maxLength={100} />
                </div>
                <div>
                  <label className="block text-xs uppercase tracking-wide text-slate-400 mb-1.5" htmlFor="notes">
                    Notes
                  </label>
                  <textarea id="notes" rows={3} className={`${inputClass} h-auto py-2`} value={notes} onChange={(e) => setNotes(e.target.value)} maxLength={2000} />
                </div>
                <Button size="sm" disabled={busy} onClick={saveDetails} className="bg-purple-600 hover:bg-purple-700 text-white">
                  {saved ? 'Saved' : 'Save details'}
                </Button>
              </div>

              <div className={`${cardClass} p-5`}>
                <h3 className="font-semibold text-white mb-4">History</h3>
                <ol className="space-y-4">
                  {[...events].reverse().map((e) => (
                    <li key={e.id} className="flex gap-3">
                      <span className="mt-1.5 h-2 w-2 flex-shrink-0 rounded-full bg-purple-400" />
                      <div>
                        <StatusBadge status={e.status} />
                        <div className="text-xs text-slate-400 mt-1">{fmtDateTime(e.at)}</div>
                        {e.note && <div className="text-xs text-slate-500">{e.note}</div>}
                      </div>
                    </li>
                  ))}
                </ol>
              </div>
            </div>

            <div className={`${cardClass} overflow-hidden`}>
              <div className="flex flex-wrap items-center justify-between gap-3 p-5 border-b border-slate-700/60">
                <h3 className="font-semibold text-white">Donations in this claim</h3>
                <input className={`${inputClass} max-w-xs`} placeholder="Filter by name, address, postcode" value={filter} onChange={(e) => setFilter(e.target.value)} />
              </div>
              <div className="overflow-auto max-h-[560px]">
                <table className="w-full text-sm">
                  <thead className="sticky top-0 bg-slate-800">
                    <tr className="text-left text-xs uppercase tracking-wide text-slate-400 border-b border-slate-700/60">
                      <th className="px-4 py-3 font-medium">Donor</th>
                      <th className="px-4 py-3 font-medium">Address</th>
                      <th className="px-4 py-3 font-medium">Postcode</th>
                      <th className="px-4 py-3 font-medium">Date</th>
                      <th className="px-4 py-3 font-medium text-right">Amount</th>
                    </tr>
                  </thead>
                  <tbody>
                    {shown.map((d) => (
                      <tr key={d.id} className="border-b border-slate-700/40 last:border-0">
                        <td className="px-4 py-2.5 text-slate-200">{[d.title, d.firstName, d.lastName].filter(Boolean).join(' ')}</td>
                        <td className="px-4 py-2.5 text-slate-400">{d.address}</td>
                        <td className="px-4 py-2.5 font-mono text-slate-300">{d.postcode}</td>
                        <td className="px-4 py-2.5 text-slate-300 whitespace-nowrap">{fmtPlainDate(d.donationDate)}</td>
                        <td className="px-4 py-2.5 text-right text-white">{gbp(d.amount)}</td>
                      </tr>
                    ))}
                    {shown.length === 0 && (
                      <tr>
                        <td colSpan={5} className="px-4 py-8 text-center text-slate-500">No matching donations</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}
      </main>
    </div>
  )
}
