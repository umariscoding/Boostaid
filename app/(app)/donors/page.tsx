'use client'

import { Fragment, useEffect, useRef, useState } from 'react'
import Link from 'next/link'
import { CheckCircle2, ChevronRight, Clock, Send, SlidersHorizontal } from 'lucide-react'
import { TableSkeleton } from '@/components/ui/skeleton'
import { Card, PageHeader, inputClass, labelClass, th, thr, td, tdr } from '@/components/ui/primitives'
import { Pagination } from '@/components/ui/pagination'
import NoCharityNotice from '@/components/charity/no-charity-notice'
import { useCharity } from '@/lib/store/charity'
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
  charityId: string
  charityName: string
  name: string
  house: string
  postcode: string
  donationCount: number
  total: number
  claims: { id: string; ref: string; status: Status }[]
  donations: DonationRow[]
}

const SLOTS = [
  { status: 'pending', word: 'pending', Icon: Clock, tone: 'text-warn' },
  { status: 'submitted', word: 'submitted', Icon: Send, tone: 'text-info' },
  { status: 'received', word: 'received', Icon: CheckCircle2, tone: 'text-good' },
] as const

/** How many of a donor's claims are in each state. Always three fixed-width slots, so every row lines up. */
function ClaimStatusIcons({ claims }: { claims: { status: Status }[] }) {
  return (
    <div className="flex items-center">
      {SLOTS.map(({ status, word, Icon, tone }) => {
        const n = claims.filter((c) => c.status === status).length
        return (
          <span
            key={status}
            title={`${n} ${word} claim${n === 1 ? '' : 's'}`}
            aria-label={`${n} ${word} claim${n === 1 ? '' : 's'}`}
            className={`inline-flex w-14 items-center gap-1.5 tabular-nums ${n ? tone : 'text-muted/40'}`}
          >
            <Icon className="h-4 w-4" />
            <span className={n ? 'text-ink' : ''}>{n}</span>
          </span>
        )
      })}
    </div>
  )
}

export default function DonorsPage() {
  const { loaded, charities, selected, selectedId } = useCharity()
  const showCharity = selectedId === 'all'
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

  const scopeRef = useRef(selectedId)

  // Search as you type. Anything with a digit is treated as a postcode, otherwise a name.
  useEffect(() => {
    if (!loaded || charities.length === 0) return
    // Switching charity starts again from page 1
    if (scopeRef.current !== selectedId) {
      scopeRef.current = selectedId
      setDonors(null)
      if (page !== 1) {
        setPage(1)
        return
      }
    }
    const params = new URLSearchParams()
    params.set('charity', selectedId)
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
  }, [loaded, charities.length, selectedId, query, filters, page, pageSize])

  const filtered = !!(query.trim() || filters.from || filters.to || filters.amount)

  return (
    <>
      <PageHeader
        title="Search donors"
        subtitle={`Find out whether a donor's donation has been claimed, when, and in which claim. ${selected ? `Searching ${selected.name}.` : 'Searching all charities.'}`}
      />

      {loaded && charities.length === 0 ? (
        <NoCharityNotice />
      ) : (
      <>
      <div className="max-w-3xl">
        <label className={labelClass} htmlFor="q">Find a donor by name or postcode</label>
        <div className="flex gap-3">
          <input
            id="q"
            autoFocus
            className="w-full h-14 px-6 rounded-full bg-white border border-line text-ink text-base placeholder:text-muted/70 focus:outline-none focus:ring-2 focus:ring-brand/30 focus:border-brand/50"
            placeholder="e.g. Sample or E1 2AB"
            value={query}
            onChange={(e) => {
              setQuery(e.target.value)
              setPage(1)
            }}
          />
          <button
            onClick={() => setShowFilters((s) => !s)}
            className={`h-14 px-6 rounded-full border text-sm font-medium inline-flex items-center gap-2 cursor-pointer ${
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

      <div className="mt-8 mb-2 flex flex-wrap items-center justify-end gap-x-5 gap-y-1 text-sm text-muted">
        <span>Claims:</span>
        {SLOTS.map(({ status, word, Icon, tone }) => (
          <span key={status} className="inline-flex items-center gap-1.5">
            <Icon className={`h-4 w-4 ${tone}`} /> {word.charAt(0).toUpperCase() + word.slice(1)}
          </span>
        ))}
      </div>

      <Card className="overflow-hidden">
        {donors === null ? (
          <TableSkeleton
            minWidth={1000}
            columns={[
              { label: '', width: '3rem' },
              { label: 'Donor', width: '22%' },
              ...(showCharity ? [{ label: 'Charity', width: '15%' }] : []),
              { label: 'House', width: '17%' },
              { label: 'Postcode', width: '11%' },
              { label: 'Donations', right: true, width: '9%' },
              { label: 'Total given', right: true, width: '11%' },
              { label: 'Claims' },
            ]}
            rows={Math.min(pageSize, 8)}
          />
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
          <div className={`overflow-x-auto transition-opacity ${loading ? 'opacity-50 pointer-events-none' : ''}`} aria-busy={loading}>
            <table className="w-full min-w-[1000px] table-fixed text-base">
              <colgroup>
                <col className="w-12" />
                <col className="w-[22%]" />
                {showCharity && <col className="w-[15%]" />}
                <col className="w-[17%]" />
                <col className="w-[11%]" />
                <col className="w-[9%]" />
                <col className="w-[11%]" />
                <col />
              </colgroup>
              <thead>
                <tr className="border-b border-line">
                  <th className="py-3.5" />
                  <th className={th}>Donor</th>
                  {showCharity && <th className={th}>Charity</th>}
                  <th className={th}>House</th>
                  <th className={th}>Postcode</th>
                  <th className={thr}>Donations</th>
                  <th className={thr}>Total given</th>
                  <th className={th}>Claims</th>
                </tr>
              </thead>
              <tbody>
                {donors.map((g) => {
                  const isOpen = open === g.key
                  return (
                    <Fragment key={g.key}>
                      <tr
                        className={`border-b border-line hover:bg-wash/60 cursor-pointer ${isOpen ? 'bg-wash/60' : ''}`}
                        onClick={() => setOpen(isOpen ? null : g.key)}
                      >
                        <td className="pl-5 py-4 align-middle text-muted">
                          <ChevronRight className={`w-4 h-4 transition-transform ${isOpen ? 'rotate-90' : ''}`} />
                        </td>
                        <td className={`${td} truncate text-ink`} title={g.name}>{g.name}</td>
                        {showCharity && <td className={`${td} truncate text-muted`} title={g.charityName}>{g.charityName}</td>}
                        <td className={`${td} truncate text-ink`} title={g.house}>{g.house}</td>
                        <td className={`${td} truncate text-ink`}>{g.postcode}</td>
                        <td className={`${tdr} text-ink`}>{g.donationCount}</td>
                        <td className={`${tdr} text-ink`}>{gbp(g.total)}</td>
                        <td className={td}>
                          <ClaimStatusIcons claims={g.claims} />
                        </td>
                      </tr>
                      {isOpen && (
                        <tr className="bg-wash border-b border-line">
                          <td colSpan={showCharity ? 8 : 7} className="pb-5 pt-1 pr-5 pl-12">
                            <div className="bg-white border border-line rounded-xl overflow-hidden">
                              <table className="w-full table-fixed text-sm">
                                <colgroup>
                                  <col className="w-[16%]" />
                                  <col className="w-[14%]" />
                                  <col className="w-[24%]" />
                                  <col className="w-[16%]" />
                                  <col className="w-[15%]" />
                                  <col className="w-[15%]" />
                                </colgroup>
                                <thead>
                                  <tr className="border-b border-line">
                                    <th className={th}>Donated</th>
                                    <th className={thr}>Amount</th>
                                    <th className={th}>Claim</th>
                                    <th className={th}>Status</th>
                                    <th className={th}>Submitted</th>
                                    <th className={th}>Received</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {g.donations.map((d) => (
                                    <tr key={d.id} className="border-b border-line last:border-0">
                                      <td className="px-5 py-3 align-middle text-ink whitespace-nowrap">{fmtPlainDate(d.donationDate)}</td>
                                      <td className="px-5 py-3 align-middle text-right tabular-nums text-ink">{gbp(d.amount)}</td>
                                      <td className="px-5 py-3 align-middle">
                                        <Link href={`/claims/${d.claimId}`} className="text-ink hover:text-brand tabular-nums">{d.claimRef}</Link>
                                        {d.hmrcRef && <div className="text-xs text-muted">HMRC: {d.hmrcRef}</div>}
                                      </td>
                                      <td className="px-5 py-3 align-middle"><StatusBadge status={d.status} /></td>
                                      <td className="px-5 py-3 align-middle text-ink whitespace-nowrap">{fmtDay(d.submittedAt)}</td>
                                      <td className="px-5 py-3 align-middle text-ink whitespace-nowrap">{fmtDay(d.receivedAt)}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
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
        </p>
      )}
      </>
      )}
    </>
  )
}
