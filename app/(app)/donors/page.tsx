'use client'

import { Fragment, useEffect, useRef, useState } from 'react'
import Link from 'next/link'
import { ChevronRight, Loader2, SlidersHorizontal } from 'lucide-react'
import { Card, PageHeader, inputClass, labelClass, th, td } from '@/components/ui/primitives'
import { Pagination } from '@/components/ui/pagination'
import { StatusBadge, apiJson, fmtDay, fmtPlainDate, gbp, type Status } from '@/components/claims/shared'

interface DonationRow {
  id: string
  donationDate: string | null
  amount: number
  claimId: string
  claimRef: string
  hmrcRef: string | null
  status: Status
  submittedAt: string | null
  receivedAt: string | null
}

interface Donor {
  key: string
  name: string
  house: string
  postcode: string
  donationCount: number
  total: number
  claims: { id: string; ref: string; status: Status }[]
  donations: DonationRow[]
}

export default function DonorsPage() {
  const [query, setQuery] = useState('')
  const [filters, setFilters] = useState({ from: '', to: '', amount: '' })
  const [showFilters, setShowFilters] = useState(false)
  const [donors, setDonors] = useState<Donor[] | null>(null)
  const [total, setTotal] = useState(0)
  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = useState(25)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [open, setOpen] = useState<string | null>(null)
  const requestId = useRef(0)

  // Search as you type. Anything with a digit is treated as a postcode, otherwise a name.
  useEffect(() => {
    const params = new URLSearchParams()
    const q = query.trim()
    if (q) params.set(/\d/.test(q) ? 'postcode' : 'q', q)
    if (filters.from) params.set('from', filters.from)
    if (filters.to) params.set('to', filters.to)
    if (filters.amount) params.set('amount', filters.amount)
    params.set('page', String(page))
    params.set('pageSize', String(pageSize))

    const id = ++requestId.current
    const timer = setTimeout(async () => {
      setLoading(true)
      try {
        const data = await apiJson<{ donors: Donor[]; total: number }>(`/api/donors?${params}`)
        if (id !== requestId.current) return
        setDonors(data.donors)
        setTotal(data.total)
        setOpen(null)
        setError(null)
      } catch (err) {
        if (id === requestId.current) setError((err as Error).message)
      } finally {
        if (id === requestId.current) setLoading(false)
      }
    }, 250)
    return () => clearTimeout(timer)
  }, [query, filters, page, pageSize])

  const filtered = !!(query.trim() || filters.from || filters.to || filters.amount)

  return (
    <>
      <PageHeader title="Search donors" subtitle="Find out whether a donor's donation has been claimed, when, and in which claim." />

      <div className="max-w-3xl">
        <label className={labelClass} htmlFor="q">Find a donor by name or postcode</label>
        <div className="flex gap-3">
          <input
            id="q"
            autoFocus
            className="w-full h-16 px-6 rounded-full bg-white border border-line text-ink text-lg placeholder:text-muted/70 focus:outline-none focus:ring-2 focus:ring-brand/30 focus:border-brand/50"
            placeholder="e.g. Sample or E1 2AB"
            value={query}
            onChange={(e) => {
              setQuery(e.target.value)
              setPage(1)
            }}
          />
          <button
            onClick={() => setShowFilters((s) => !s)}
            className={`h-16 px-6 rounded-full border text-sm font-medium inline-flex items-center gap-2 cursor-pointer ${
              showFilters ? 'bg-brand-soft border-brand/30 text-brand' : 'bg-white border-line text-ink hover:bg-wash'
            }`}
          >
            <SlidersHorizontal className="w-4 h-4" /> Filters
          </button>
        </div>

        {showFilters && (
          <Card className="mt-3 p-5 grid grid-cols-1 sm:grid-cols-3 gap-4">
            <div>
              <label className={labelClass} htmlFor="from">Donated from</label>
              <input id="from" type="date" className={inputClass} value={filters.from} onChange={(e) => { setFilters({ ...filters, from: e.target.value }); setPage(1) }} />
            </div>
            <div>
              <label className={labelClass} htmlFor="to">Donated to</label>
              <input id="to" type="date" className={inputClass} value={filters.to} onChange={(e) => { setFilters({ ...filters, to: e.target.value }); setPage(1) }} />
            </div>
            <div>
              <label className={labelClass} htmlFor="amt">Amount (£)</label>
              <input id="amt" type="number" step="0.01" min="0" className={inputClass} value={filters.amount} onChange={(e) => { setFilters({ ...filters, amount: e.target.value }); setPage(1) }} placeholder="10.00" />
            </div>
          </Card>
        )}
      </div>

      {error && <div className="mt-6 rounded-xl border border-bad/20 bg-bad-soft px-4 py-3 text-sm text-bad">{error}</div>}

      <Card className="mt-8 overflow-hidden">
        {donors === null ? (
          <div className="flex items-center justify-center gap-3 py-16 text-muted">
            <Loader2 className="w-5 h-5 animate-spin" /> Loading donors...
          </div>
        ) : donors.length === 0 ? (
          <div className="px-6 py-16 text-center">
            <div className="font-display text-xl font-semibold text-ink">
              {filtered ? 'No claimed donations match' : 'No donors yet'}
            </div>
            <p className="text-muted mt-1">
              {filtered
                ? 'Nothing matching this search has been included in a Gift Aid claim. Try fewer filters or a partial name.'
                : 'Donors appear here once a spreadsheet has been cleaned and saved as a claim.'}
            </p>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-[15px]">
              <thead>
                <tr className="border-b border-line">
                  <th className="px-4 py-3.5 w-8" />
                  <th className={th}>Donor</th>
                  <th className={th}>House</th>
                  <th className={th}>Postcode</th>
                  <th className={`${th} text-right`}>Donations</th>
                  <th className={`${th} text-right`}>Total given</th>
                  <th className={th}>In claims</th>
                </tr>
              </thead>
              <tbody>
                {donors.map((g) => {
                  const isOpen = open === g.key
                  return (
                    <Fragment key={g.key}>
                      <tr
                        className="border-b border-line hover:bg-wash/60 cursor-pointer"
                        onClick={() => setOpen(isOpen ? null : g.key)}
                      >
                        <td className="pl-5 py-4 text-muted">
                          <ChevronRight className={`w-4 h-4 transition-transform ${isOpen ? 'rotate-90' : ''}`} />
                        </td>
                        <td className={`${td} text-ink`}>{g.name}</td>
                        <td className={`${td} text-ink`}>{g.house}</td>
                        <td className={`${td} text-ink whitespace-nowrap`}>{g.postcode}</td>
                        <td className={`${td} text-right text-ink`}>{g.donationCount}</td>
                        <td className={`${td} text-right text-ink whitespace-nowrap`}>{gbp(g.total)}</td>
                        <td className={td}>
                          <div className="flex flex-wrap gap-x-4 gap-y-1">
                            {g.claims.map((c) => (
                              <Link
                                key={c.id}
                                href={`/claims/${c.id}`}
                                onClick={(e) => e.stopPropagation()}
                                className="inline-flex items-center gap-2 text-ink hover:text-brand"
                              >
                                {c.ref}
                                <StatusBadge status={c.status} />
                              </Link>
                            ))}
                          </div>
                        </td>
                      </tr>
                      {isOpen && (
                        <tr className="bg-wash border-b border-line">
                          <td />
                          <td colSpan={6} className="px-5 py-4">
                            <table className="w-full text-sm">
                              <thead>
                                <tr className="text-left text-muted">
                                  <th className="py-1.5 font-medium">Donated</th>
                                  <th className="py-1.5 font-medium text-right">Amount</th>
                                  <th className="py-1.5 pl-8 font-medium">Claim</th>
                                  <th className="py-1.5 font-medium">Status</th>
                                  <th className="py-1.5 font-medium">Submitted</th>
                                  <th className="py-1.5 font-medium">Received</th>
                                </tr>
                              </thead>
                              <tbody>
                                {g.donations.map((d) => (
                                  <tr key={d.id} className="border-t border-line">
                                    <td className="py-2 text-ink whitespace-nowrap">{fmtPlainDate(d.donationDate)}</td>
                                    <td className="py-2 text-right text-ink">{gbp(d.amount)}</td>
                                    <td className="py-2 pl-8">
                                      <Link href={`/claims/${d.claimId}`} className="text-ink hover:text-brand">{d.claimRef}</Link>
                                      {d.hmrcRef && <span className="ml-2 text-muted">HMRC: {d.hmrcRef}</span>}
                                    </td>
                                    <td className="py-2"><StatusBadge status={d.status} /></td>
                                    <td className="py-2 text-ink whitespace-nowrap">{fmtDay(d.submittedAt)}</td>
                                    <td className="py-2 text-ink whitespace-nowrap">{fmtDay(d.receivedAt)}</td>
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
          sizes={[10, 25, 50]}
        />
      </Card>

      {donors && donors.length > 0 && (
        <p className="mt-4 text-muted">
          {total.toLocaleString()} donor{total === 1 ? '' : 's'}
          {filtered ? ' match' : ' in total'}
          {loading && ' \u00b7 updating...'}
        </p>
      )}
    </>
  )
}
