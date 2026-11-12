'use client'

import { useMemo, useState } from 'react'
import { ChevronRight } from 'lucide-react'
import { Pill, td, tdr, th, thr } from '@/components/ui/primitives'
import { Pagination } from '@/components/ui/pagination'
import { Select } from '@/components/ui/select'
import { gbp } from '@/components/claims/shared'
import type { ClaimResult } from '@/lib/giftaid'
import { toDDMMYY } from '@/lib/giftaid/dates'
import { changesByDonation } from '@/lib/claim-extras'

const FIELD: Record<string, string> = { Name: 'Name', 'First Name': 'First name', 'Last Name': 'Last name', House: 'House', Postcode: 'Postcode' }
const show = (v: unknown) => (v === null || v === undefined || String(v).trim() === '' ? '(empty)' : String(v))

/** Optional extras: what changed on each donation, and how HMRC's schedule will look. Not needed to submit. */
export default function OptionalChecks({ result }: { result: ClaimResult }) {
  const [open, setOpen] = useState(false)
  const [tab, setTab] = useState<'rows' | 'hmrc'>('rows')
  const [onlyChanged, setOnlyChanged] = useState(true)
  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = useState(25)
  const [fileIdx, setFileIdx] = useState(0)

  const changes = useMemo(() => changesByDonation(result), [result])
  const claimed = useMemo(() => [...result.ready, ...result.overseas], [result])
  const rows = useMemo(() => (onlyChanged ? claimed.filter((r) => r.cleaned) : claimed), [claimed, onlyChanged])
  const start = (page - 1) * pageSize
  const file = result.files[Math.min(fileIdx, result.files.length - 1)]

  return (
    <div className="overflow-hidden rounded-2xl border border-line bg-white">
      <button onClick={() => setOpen((o) => !o)} className="flex w-full cursor-pointer items-center justify-between gap-4 px-5 py-4 text-left hover:bg-wash/60">
        <div>
          <div className="flex items-center gap-3">
            <h3 className="font-display text-lg font-semibold text-ink">Check every row</h3>
            <Pill tone="neutral">Optional</Pill>
          </div>
          <p className="text-sm text-muted">Not required. See what we changed on each donation, and how HMRC&apos;s schedule will look.</p>
        </div>
        <ChevronRight className={`h-5 w-5 flex-shrink-0 text-muted transition-transform ${open ? 'rotate-90' : ''}`} />
      </button>

      {open && (
        <div className="border-t border-line">
          <div className="flex flex-wrap items-center justify-between gap-3 px-5 py-3">
            <div className="flex gap-1 rounded-full border border-line bg-wash p-1">
              {([['rows', 'Before and after'], ['hmrc', 'HMRC schedule preview']] as const).map(([id, label]) => (
                <button
                  key={id}
                  onClick={() => setTab(id)}
                  className={`h-9 cursor-pointer rounded-full border px-4 text-sm font-medium transition-colors ${tab === id ? 'border-line bg-white text-ink shadow-sm' : 'border-transparent text-muted hover:text-ink'}`}
                >
                  {label}
                </button>
              ))}
            </div>
            {tab === 'rows' ? (
              <label className="flex cursor-pointer items-center gap-2 text-sm text-ink">
                <input
                  type="checkbox"
                  checked={onlyChanged}
                  onChange={(e) => { setOnlyChanged(e.target.checked); setPage(1) }}
                  className="h-4 w-4 accent-[#4a2bb3]"
                />
                Only rows we changed
              </label>
            ) : (
              result.files.length > 1 && (
                <Select variant="compact" value={fileIdx} onChange={(e) => setFileIdx(Number(e.target.value))}>
                  {result.files.map((f, i) => <option key={f.name} value={i}>{f.name}</option>)}
                </Select>
              )
            )}
          </div>

          {tab === 'rows' ? (
            <>
              <div className="overflow-x-auto border-t border-line">
                <table className="w-full min-w-[900px] table-fixed text-base">
                  <colgroup>
                    <col className="w-[18%]" />
                    <col className="w-[16%]" />
                    <col className="w-[10%]" />
                    <col className="w-[9%]" />
                    <col className="w-[9%]" />
                    <col />
                  </colgroup>
                  <thead>
                    <tr className="border-b border-line">
                      <th className={th}>Donor</th>
                      <th className={th}>House</th>
                      <th className={th}>Postcode</th>
                      <th className={th}>Date</th>
                      <th className={thr}>Amount</th>
                      <th className={th}>What we changed</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.slice(start, start + pageSize).map((r, i) => {
                      const list = changes.get(String(r.donationId ?? '')) ?? []
                      return (
                        <tr key={start + i} className="border-b border-line align-top last:border-0">
                          <td className={`${td} truncate text-ink`} title={`${r.firstName} ${r.lastName}`}>{r.firstName} {r.lastName}</td>
                          <td className={`${td} truncate text-ink`} title={r.house}>{r.house}</td>
                          <td className={`${td} truncate text-ink`}>{r.postcode}</td>
                          <td className={`${td} whitespace-nowrap text-ink`}>{r.date ? toDDMMYY(r.date) : '-'}</td>
                          <td className={`${tdr} text-ink`}>{gbp(r.amount)}</td>
                          <td className={`${td} text-sm text-ink`}>
                            {list.length ? (
                              <div className="space-y-0.5">
                                {list.map((c, j) => (
                                  <div key={j} className="break-words">
                                    <span className="text-muted">{FIELD[c.field] ?? c.field}:</span> {show(c.original)} <span className="text-muted">→</span> {c.cleaned}
                                  </div>
                                ))}
                              </div>
                            ) : r.cleaned ? (
                              <Pill tone="info">Changed</Pill>
                            ) : (
                              <span className="text-muted">No change</span>
                            )}
                          </td>
                        </tr>
                      )
                    })}
                    {rows.length === 0 && (
                      <tr><td colSpan={6} className="px-5 py-10 text-center text-muted">{onlyChanged ? 'We did not need to change any donation.' : 'Nothing here.'}</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
              <Pagination page={page} pageSize={pageSize} total={rows.length} onPage={setPage} onPageSize={(n) => { setPageSize(n); setPage(1) }} />
            </>
          ) : (
            <div className="border-t border-line">
              <p className="px-5 pt-4 text-sm text-muted">
                The first rows of {file?.name}, as they will sit in HMRC&apos;s template from row 25. Columns H and I are left empty.
              </p>
              <div className="overflow-x-auto px-5 pb-5 pt-3">
                <table className="w-full min-w-[800px] text-sm">
                  <thead>
                    <tr className="border-b border-line text-muted">
                      {['Row', 'C Title', 'D First name', 'E Last name', 'F House name or number', 'G Postcode', 'H Aggregated', 'I Sponsored', 'J Date', 'K Amount'].map((h, i) => (
                        <th key={h} className={`px-3 py-3 font-medium ${i === 9 ? 'text-right' : 'text-left'}`}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {file?.rows.slice(0, 5).map((r, i) => (
                      <tr key={i} className="border-b border-line last:border-0">
                        <td className="px-3 py-2.5 text-muted">{r.hmrcRow}</td>
                        <td className="px-3 py-2.5 text-ink">{r.title}</td>
                        <td className="px-3 py-2.5 text-ink">{r.firstName}</td>
                        <td className="px-3 py-2.5 text-ink">{r.lastName}</td>
                        <td className="px-3 py-2.5 text-ink">{r.house}</td>
                        <td className="px-3 py-2.5 text-ink">{r.postcode}</td>
                        <td className="px-3 py-2.5 italic text-muted">empty</td>
                        <td className="px-3 py-2.5 italic text-muted">empty</td>
                        <td className="px-3 py-2.5 text-ink">{r.date ? toDDMMYY(r.date) : ''}</td>
                        <td className="px-3 py-2.5 text-right tabular-nums text-ink">{r.amount.toFixed(2)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  )
}
