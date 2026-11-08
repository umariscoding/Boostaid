'use client'

import { useState } from 'react'
import Link from 'next/link'
import { Loader2, Search } from 'lucide-react'
import AppHeader from '@/components/app-header'
import { Button } from '@/components/ui/button'
import {
  StatusBadge,
  apiJson,
  cardClass,
  fmtDay,
  fmtPlainDate,
  gbp,
  inputClass,
  type Status,
} from '@/components/claims/shared'

interface Result {
  id: string
  title: string
  firstName: string
  lastName: string
  address: string
  postcode: string
  donationDate: string | null
  amount: number
  claimId: string
  claimRef: string
  hmrcRef: string | null
  status: Status
  submittedAt: string | null
  receivedAt: string | null
}

const empty = { q: '', postcode: '', from: '', to: '', amount: '' }

export default function DonorsPage() {
  const [form, setForm] = useState(empty)
  const [results, setResults] = useState<Result[] | null>(null)
  const [truncated, setTruncated] = useState(false)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const set = (k: keyof typeof empty) => (e: React.ChangeEvent<HTMLInputElement>) => setForm({ ...form, [k]: e.target.value })

  const search = async (e: React.FormEvent) => {
    e.preventDefault()
    const params = new URLSearchParams()
    for (const [k, v] of Object.entries(form)) if (v.trim()) params.set(k, v.trim())
    if ([...params.keys()].length === 0) {
      setError('Enter a name, postcode, date or amount to search')
      return
    }
    setLoading(true)
    setError(null)
    try {
      const data = await apiJson<{ results: Result[]; truncated: boolean }>(`/api/donors?${params}`)
      setResults(data.results)
      setTruncated(data.truncated)
    } catch (err) {
      setError((err as Error).message)
      setResults(null)
    } finally {
      setLoading(false)
    }
  }

  const label = 'block text-xs uppercase tracking-wide text-slate-400 mb-1.5'

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900">
      <AppHeader />
      <main className="max-w-7xl mx-auto px-6 py-6 space-y-6">
        <div>
          <h2 className="text-2xl font-semibold text-white">Donor search</h2>
          <p className="text-sm text-slate-400">Find out whether a donor&apos;s donation has been claimed, when, and in which claim.</p>
        </div>

        <form onSubmit={search} className={`${cardClass} p-5`}>
          <div className="grid grid-cols-1 md:grid-cols-6 gap-4">
            <div className="md:col-span-2">
              <label className={label} htmlFor="q">Donor name</label>
              <input id="q" className={inputClass} value={form.q} onChange={set('q')} placeholder="e.g. Jane Smith" autoFocus />
            </div>
            <div>
              <label className={label} htmlFor="pc">Postcode</label>
              <input id="pc" className={inputClass} value={form.postcode} onChange={set('postcode')} placeholder="e.g. SW1A 1AA" />
            </div>
            <div>
              <label className={label} htmlFor="from">Donated from</label>
              <input id="from" type="date" className={inputClass} value={form.from} onChange={set('from')} />
            </div>
            <div>
              <label className={label} htmlFor="to">Donated to</label>
              <input id="to" type="date" className={inputClass} value={form.to} onChange={set('to')} />
            </div>
            <div>
              <label className={label} htmlFor="amt">Amount (£)</label>
              <input id="amt" type="number" step="0.01" min="0" className={inputClass} value={form.amount} onChange={set('amount')} placeholder="10.00" />
            </div>
          </div>
          <div className="mt-4 flex items-center gap-3">
            <Button type="submit" disabled={loading} className="bg-gradient-to-r from-violet-500 to-purple-600 hover:from-violet-600 hover:to-purple-700 text-white">
              {loading ? <Loader2 className="w-4 h-4 animate-spin" /> : <Search className="w-4 h-4" />}
              Search
            </Button>
            <button type="button" className="text-sm text-slate-400 hover:text-white" onClick={() => { setForm(empty); setResults(null); setError(null) }}>
              Clear
            </button>
          </div>
        </form>

        {error && <div className="rounded-lg border border-red-500/30 bg-red-500/10 px-4 py-3 text-sm text-red-200">{error}</div>}

        {results && (
          <div className={`${cardClass} overflow-hidden`}>
            {results.length === 0 ? (
              <div className="px-6 py-14 text-center">
                <div className="text-white font-medium">No claimed donations match</div>
                <p className="text-sm text-slate-400 mt-1">
                  Nothing matching this search has been included in a Gift Aid claim. Try fewer filters or a partial name.
                </p>
              </div>
            ) : (
              <>
                <div className="px-5 py-3 border-b border-slate-700/60 text-sm text-slate-300">
                  {results.length}{truncated ? '+' : ''} donation{results.length === 1 ? '' : 's'} found
                  {truncated && ' (showing the first 200 - narrow your search)'}
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="text-left text-xs uppercase tracking-wide text-slate-400 border-b border-slate-700/60">
                        <th className="px-4 py-3 font-medium">Donor</th>
                        <th className="px-4 py-3 font-medium">Postcode</th>
                        <th className="px-4 py-3 font-medium">Donated</th>
                        <th className="px-4 py-3 font-medium text-right">Amount</th>
                        <th className="px-4 py-3 font-medium">Claim</th>
                        <th className="px-4 py-3 font-medium">Status</th>
                        <th className="px-4 py-3 font-medium">Submitted</th>
                        <th className="px-4 py-3 font-medium">Received</th>
                      </tr>
                    </thead>
                    <tbody>
                      {results.map((r) => (
                        <tr key={r.id} className="border-b border-slate-700/40 last:border-0 hover:bg-slate-700/20">
                          <td className="px-4 py-3">
                            <div className="text-slate-100">{[r.title, r.firstName, r.lastName].filter(Boolean).join(' ')}</div>
                            <div className="text-xs text-slate-500">{r.address}</div>
                          </td>
                          <td className="px-4 py-3 font-mono text-slate-300">{r.postcode}</td>
                          <td className="px-4 py-3 text-slate-300 whitespace-nowrap">{fmtPlainDate(r.donationDate)}</td>
                          <td className="px-4 py-3 text-right font-semibold text-white">{gbp(r.amount)}</td>
                          <td className="px-4 py-3">
                            <Link href={`/claims/${r.claimId}`} className="font-mono text-purple-300 hover:text-purple-200">{r.claimRef}</Link>
                            {r.hmrcRef && <div className="text-xs text-slate-500">HMRC: {r.hmrcRef}</div>}
                          </td>
                          <td className="px-4 py-3"><StatusBadge status={r.status} /></td>
                          <td className="px-4 py-3 text-slate-300 whitespace-nowrap">{fmtDay(r.submittedAt)}</td>
                          <td className="px-4 py-3 text-slate-300 whitespace-nowrap">{fmtDay(r.receivedAt)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </div>
        )}
      </main>
    </div>
  )
}
