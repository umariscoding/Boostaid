'use client'

import { Fragment, useMemo, useState } from 'react'
import Link from 'next/link'
import { ChevronRight, Loader2, Search } from 'lucide-react'
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

interface DonorGroup {
  key: string
  name: string
  house: string
  postcode: string
  donations: Result[]
  total: number
  claims: { id: string; ref: string; status: Status }[]
}

/** One row per donor (same name + postcode), with every donation and the claims it appears in. */
function groupByDonor(results: Result[]): DonorGroup[] {
  const map = new Map<string, DonorGroup>()
  for (const r of results) {
    const key = `${r.firstName}|${r.lastName}|${r.postcode.replace(/\s+/g, '')}`.toLowerCase()
    let g = map.get(key)
    if (!g) {
      g = {
        key,
        name: [r.title, r.firstName, r.lastName].filter(Boolean).join(' '),
        house: r.address,
        postcode: r.postcode,
        donations: [],
        total: 0,
        claims: [],
      }
      map.set(key, g)
    }
    g.donations.push(r)
    g.total += r.amount
    if (!g.claims.some((c) => c.id === r.claimId)) g.claims.push({ id: r.claimId, ref: r.claimRef, status: r.status })
  }
  return [...map.values()]
}

export default function DonorsPage() {
  const [form, setForm] = useState(empty)
  const [results, setResults] = useState<Result[] | null>(null)
  const [truncated, setTruncated] = useState(false)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [open, setOpen] = useState<string | null>(null)
  const groups = useMemo(() => (results ? groupByDonor(results) : []), [results])
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
      setOpen(null)
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
            {groups.length === 0 ? (
              <div className="px-6 py-14 text-center">
                <div className="text-white font-medium">No claimed donations match</div>
                <p className="text-sm text-slate-400 mt-1">
                  Nothing matching this search has been included in a Gift Aid claim. Try fewer filters or a partial name.
                </p>
              </div>
            ) : (
              <>
                <div className="px-5 py-3 border-b border-slate-700/60 text-sm text-slate-300">
                  {groups.length} donor{groups.length === 1 ? '' : 's'} · {results.length}{truncated ? '+' : ''} donation{results.length === 1 ? '' : 's'}
                  {truncated && ' (showing the first 200 donations - narrow your search)'}
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="text-left text-xs uppercase tracking-wide text-slate-400 border-b border-slate-700/60">
                        <th className="px-4 py-3 w-8" />
                        <th className="px-4 py-3 font-medium">Donor</th>
                        <th className="px-4 py-3 font-medium">House</th>
                        <th className="px-4 py-3 font-medium">Postcode</th>
                        <th className="px-4 py-3 font-medium text-right">Donations</th>
                        <th className="px-4 py-3 font-medium text-right">Total given</th>
                        <th className="px-4 py-3 font-medium">In claims</th>
                      </tr>
                    </thead>
                    <tbody>
                      {groups.map((g) => {
                        const isOpen = open === g.key
                        return (
                          <Fragment key={g.key}>
                            <tr
                              className="border-b border-slate-700/40 hover:bg-slate-700/20 cursor-pointer"
                              onClick={() => setOpen(isOpen ? null : g.key)}
                            >
                              <td className="px-4 py-3 text-slate-500">
                                <ChevronRight className={`w-4 h-4 transition-transform ${isOpen ? 'rotate-90' : ''}`} />
                              </td>
                              <td className="px-4 py-3 text-slate-100 font-medium">{g.name}</td>
                              <td className="px-4 py-3 text-slate-300">{g.house}</td>
                              <td className="px-4 py-3 font-mono text-slate-300">{g.postcode}</td>
                              <td className="px-4 py-3 text-right text-slate-300">{g.donations.length}</td>
                              <td className="px-4 py-3 text-right font-semibold text-white">{gbp(g.total)}</td>
                              <td className="px-4 py-3">
                                <div className="flex flex-wrap gap-1.5">
                                  {g.claims.map((c) => (
                                    <Link
                                      key={c.id}
                                      href={`/claims/${c.id}`}
                                      onClick={(e) => e.stopPropagation()}
                                      className="inline-flex items-center gap-1.5 font-mono text-xs text-purple-300 hover:text-purple-200"
                                    >
                                      {c.ref}
                                      <StatusBadge status={c.status} />
                                    </Link>
                                  ))}
                                </div>
                              </td>
                            </tr>
                            {isOpen && (
                              <tr className="bg-slate-900/40 border-b border-slate-700/40">
                                <td />
                                <td colSpan={6} className="px-4 py-3">
                                  <table className="w-full text-xs">
                                    <thead>
                                      <tr className="text-left uppercase tracking-wide text-slate-500">
                                        <th className="py-1.5 font-medium">Donated</th>
                                        <th className="py-1.5 font-medium text-right">Amount</th>
                                        <th className="py-1.5 pl-6 font-medium">Claim</th>
                                        <th className="py-1.5 font-medium">Status</th>
                                        <th className="py-1.5 font-medium">Submitted</th>
                                        <th className="py-1.5 font-medium">Received</th>
                                      </tr>
                                    </thead>
                                    <tbody>
                                      {g.donations.map((d) => (
                                        <tr key={d.id} className="border-t border-slate-700/30">
                                          <td className="py-1.5 text-slate-300 whitespace-nowrap">{fmtPlainDate(d.donationDate)}</td>
                                          <td className="py-1.5 text-right text-white">{gbp(d.amount)}</td>
                                          <td className="py-1.5 pl-6">
                                            <Link href={`/claims/${d.claimId}`} className="font-mono text-purple-300 hover:text-purple-200">{d.claimRef}</Link>
                                            {d.hmrcRef && <span className="ml-2 text-slate-500">HMRC: {d.hmrcRef}</span>}
                                          </td>
                                          <td className="py-1.5"><StatusBadge status={d.status} /></td>
                                          <td className="py-1.5 text-slate-300 whitespace-nowrap">{fmtDay(d.submittedAt)}</td>
                                          <td className="py-1.5 text-slate-300 whitespace-nowrap">{fmtDay(d.receivedAt)}</td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </td>
                              </tr>
                            )}
                          </Fragment>
                        )
                      })}
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
