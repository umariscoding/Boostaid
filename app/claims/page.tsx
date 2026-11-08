'use client'

import { useCallback, useEffect, useState } from 'react'
import Link from 'next/link'
import { Download, Loader2, Undo2 } from 'lucide-react'
import AppHeader from '@/components/app-header'
import { Button } from '@/components/ui/button'
import type { ClaimDto } from '@/lib/claims'
import {
  NEXT_STATUS,
  PREV_STATUS,
  StatusBadge,
  apiJson,
  cardClass,
  fmtDay,
  gbp,
  type Status,
} from '@/components/claims/shared'

type View = 'active' | 'archive' | 'all'
type Summary = Record<Status, { count: number; amount: number }>

const TABS: { id: View; label: string }[] = [
  { id: 'active', label: 'Active' },
  { id: 'archive', label: 'Archive (received)' },
  { id: 'all', label: 'All' },
]

export default function ClaimsPage() {
  const [view, setView] = useState<View>('active')
  const [claims, setClaims] = useState<ClaimDto[] | null>(null)
  const [summary, setSummary] = useState<Summary | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [busyId, setBusyId] = useState<string | null>(null)

  const load = useCallback(async (v: View) => {
    try {
      const data = await apiJson<{ claims: ClaimDto[]; summary: Summary }>(`/api/claims?view=${v}`)
      setClaims(data.claims)
      setSummary(data.summary)
      setError(null)
    } catch (e) {
      setError((e as Error).message)
    }
  }, [])

  useEffect(() => {
    setClaims(null)
    load(view)
  }, [view, load])

  const changeStatus = async (id: string, status: Status) => {
    setBusyId(id)
    try {
      await apiJson(`/api/claims/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status }),
      })
      await load(view)
    } catch (e) {
      setError((e as Error).message)
    } finally {
      setBusyId(null)
    }
  }

  const tiles: { status: Status; label: string; tone: string }[] = [
    { status: 'pending', label: 'Pending', tone: 'text-amber-200' },
    { status: 'submitted', label: 'Submitted - awaiting payment', tone: 'text-sky-200' },
    { status: 'received', label: 'Received', tone: 'text-emerald-200' },
  ]

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900">
      <AppHeader />
      <main className="max-w-7xl mx-auto px-6 py-6 space-y-6">
        <div className="flex flex-wrap items-end justify-between gap-3">
          <div>
            <h2 className="text-2xl font-semibold text-white">Claims</h2>
            <p className="text-sm text-slate-400">Track each HMRC submission from pending to received.</p>
          </div>
          <a href={`/api/export?view=${view}`}>
            <Button size="sm" variant="outline" className="border-slate-600 bg-slate-700/50 text-slate-200 hover:bg-slate-600 hover:text-white">
              <Download className="w-4 h-4 mr-2" />
              Download audit file
            </Button>
          </a>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          {tiles.map((t) => (
            <div key={t.status} className={`${cardClass} p-4`}>
              <div className="text-xs uppercase tracking-wide text-slate-400 mb-1">{t.label}</div>
              <div className={`text-2xl font-bold ${t.tone}`}>{summary ? gbp(summary[t.status].amount) : '-'}</div>
              <div className="text-xs text-slate-500 mt-0.5">
                {summary ? `${summary[t.status].count} claim${summary[t.status].count === 1 ? '' : 's'}` : ' '}
              </div>
            </div>
          ))}
        </div>

        <div className="flex gap-1 border-b border-slate-700/60">
          {TABS.map((t) => (
            <button
              key={t.id}
              onClick={() => setView(t.id)}
              className={`px-4 py-2 text-sm font-medium -mb-px border-b-2 transition-colors ${
                view === t.id ? 'border-purple-400 text-white' : 'border-transparent text-slate-400 hover:text-slate-200'
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>

        {error && <div className="rounded-lg border border-red-500/30 bg-red-500/10 px-4 py-3 text-sm text-red-200">{error}</div>}

        <div className={`${cardClass} overflow-hidden`}>
          {claims === null && !error ? (
            <div className="flex items-center justify-center gap-3 py-16 text-slate-400">
              <Loader2 className="w-5 h-5 animate-spin" /> Loading claims...
            </div>
          ) : claims && claims.length === 0 ? (
            <div className="py-16 text-center text-slate-400">
              {view === 'archive' ? 'No received claims yet.' : 'No claims yet.'}{' '}
              {view !== 'archive' && (
                <Link href="/" className="text-purple-300 hover:text-purple-200">
                  Process a spreadsheet
                </Link>
              )}
            </div>
          ) : (
            claims && (
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-left text-xs uppercase tracking-wide text-slate-400 border-b border-slate-700/60">
                      <th className="px-4 py-3 font-medium">Claim ref</th>
                      <th className="px-4 py-3 font-medium">Created</th>
                      <th className="px-4 py-3 font-medium">Source file</th>
                      <th className="px-4 py-3 font-medium text-right">Donations</th>
                      <th className="px-4 py-3 font-medium text-right">Amount</th>
                      <th className="px-4 py-3 font-medium">Status</th>
                      <th className="px-4 py-3 font-medium text-right">Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {claims.map((c) => {
                      const next = NEXT_STATUS[c.status]
                      const prev = PREV_STATUS[c.status]
                      const busy = busyId === c.id
                      return (
                        <tr key={c.id} className="border-b border-slate-700/40 last:border-0 hover:bg-slate-700/20">
                          <td className="px-4 py-3">
                            <Link href={`/claims/${c.id}`} className="font-mono font-semibold text-purple-300 hover:text-purple-200">
                              {c.ref}
                            </Link>
                            {c.hmrcRef && <div className="text-xs text-slate-500">HMRC: {c.hmrcRef}</div>}
                          </td>
                          <td className="px-4 py-3 text-slate-300 whitespace-nowrap">{fmtDay(c.createdAt)}</td>
                          <td className="px-4 py-3 text-slate-400 max-w-[220px] truncate" title={c.sourceFile}>
                            {c.sourceFile}
                          </td>
                          <td className="px-4 py-3 text-right text-slate-300">{c.donationCount.toLocaleString()}</td>
                          <td className="px-4 py-3 text-right font-semibold text-white whitespace-nowrap">{gbp(c.totalAmount)}</td>
                          <td className="px-4 py-3">
                            <StatusBadge status={c.status} />
                            {c.status === 'received' && <div className="text-xs text-slate-500 mt-1">{fmtDay(c.receivedAt)}</div>}
                            {c.status === 'submitted' && <div className="text-xs text-slate-500 mt-1">{fmtDay(c.submittedAt)}</div>}
                          </td>
                          <td className="px-4 py-3">
                            <div className="flex items-center justify-end gap-2">
                              {prev && (
                                <button
                                  title={`Move back to ${prev}`}
                                  disabled={busy}
                                  onClick={() => {
                                    if (confirm(`Move ${c.ref} back to ${prev}?`)) changeStatus(c.id, prev)
                                  }}
                                  className="p-1.5 rounded-md text-slate-400 hover:text-white hover:bg-slate-700 disabled:opacity-50"
                                >
                                  <Undo2 className="w-4 h-4" />
                                </button>
                              )}
                              {next && (
                                <Button
                                  size="sm"
                                  disabled={busy}
                                  onClick={() => changeStatus(c.id, next.to)}
                                  className={
                                    next.to === 'received'
                                      ? 'bg-emerald-600 hover:bg-emerald-700 text-white'
                                      : 'bg-gradient-to-r from-violet-500 to-purple-600 hover:from-violet-600 hover:to-purple-700 text-white'
                                  }
                                >
                                  {busy && <Loader2 className="w-3.5 h-3.5 animate-spin" />}
                                  {next.label}
                                </Button>
                              )}
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
        </div>
      </main>
    </div>
  )
}
