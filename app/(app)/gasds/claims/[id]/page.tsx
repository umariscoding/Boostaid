'use client'

import { use, useEffect, useState } from 'react'
import Link from 'next/link'
import { ArrowLeft, Download, Printer } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Card, PageHeader, th, thr, td, tdr } from '@/components/ui/primitives'
import { Skeleton } from '@/components/ui/skeleton'
import { apiJson, gbp } from '@/components/claims/shared'
import { countText, p, showNo, taxYearLabel, ukDate, who, type GReceipt } from '@/lib/gasds'
import { RULES_TEXT, evidenceSummary } from '@/lib/gasds-summary'
import type { GasdsClaimDto } from '@/lib/gasds-db'

/** What an auditor needs for one small donations claim: the claim, the Gift Aid rule, and each receipt with its counts and slip. */
export default function GasdsEvidencePage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params)
  const [data, setData] = useState<{ claim: GasdsClaimDto; charity: string; receipts: GReceipt[] } | null>(null)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    apiJson<{ claim: GasdsClaimDto; charity: string; receipts: GReceipt[] }>(`/api/gasds/claims/${id}`).then(setData, (e) => setError((e as Error).message))
  }, [id])

  const back = (
    <Link href="/gasds" className="inline-flex items-center gap-1.5 text-sm font-medium text-brand hover:underline print:hidden">
      <ArrowLeft className="h-4 w-4" /> Small donations
    </Link>
  )
  if (error) return <>{back}<div role="alert" className="mt-4 rounded-xl border border-bad-line bg-bad-soft px-4 py-3 text-sm font-medium text-bad">{error}</div></>
  if (!data) return <>{back}<div className="mt-4 space-y-4" aria-busy="true"><Skeleton className="h-10 w-64" /><Skeleton className="h-56 w-full rounded-2xl" /></div></>

  const { claim: c, charity, receipts } = data

  return (
    <div className="space-y-6">
      {back}
      <PageHeader
        title={`Evidence · ${c.hmrcRef}`}
        subtitle={`${charity} · small donations claim for ${taxYearLabel(c.taxYear)}`}
        action={
          <div className="flex flex-wrap gap-2 print:hidden">
            <Button variant="outline" onClick={() => window.print()}><Printer className="h-4 w-4" /> Print or save as PDF</Button>
            <a href={`/api/gasds/claims/${id}/evidence`}><Button variant="dark"><Download className="h-4 w-4" /> Download (printout + Excel)</Button></a>
          </div>
        }
      />

      <Card className="max-w-2xl p-6">
        <dl className="grid grid-cols-[1fr_auto] gap-x-8 gap-y-2 text-base">
          {evidenceSummary(c, charity).map(([k, v]) => (
            <div key={k} className="contents">
              <dt className="text-muted">{k}</dt>
              <dd className={`text-right font-semibold tabular-nums ${v.endsWith('✗') ? 'text-bad' : v.endsWith('✓') ? 'text-good' : ''}`}>{v}</dd>
            </div>
          ))}
        </dl>
      </Card>

      <Card className="overflow-x-auto print:hidden">
        <table className="w-full min-w-[960px] text-sm">
          <thead>
            <tr className="border-b border-line">
              <th className={th}>Date</th><th className={th}>Receipt</th><th className={th}>From</th><th className={th}>Paid by</th>
              <th className={th}>Count</th><th className={th}>Bank slip</th><th className={thr}>Receipt</th><th className={thr}>Used</th>
            </tr>
          </thead>
          <tbody>
            {receipts.map((r) => (
              <tr key={r.id} className="border-b border-line-soft align-top">
                <td className={`${td} whitespace-nowrap`}>{ukDate(r.date)}</td>
                <td className={td}>{showNo(r.no)}</td>
                <td className={td}>{who(r)}<div className="text-xs text-muted">{r.gift ? 'individual' : 'collection'}</div></td>
                <td className={td}>{r.card ? 'Card' : 'Cash'}</td>
                <td className={`${td} text-muted`}>
                  {countText(r)}
                  {r.photo && <div className="mt-1 print:break-inside-avoid"><a href={`/api/gasds/receipts/${r.id}/photo`} target="_blank" rel="noreferrer" className="text-brand underline print:hidden">View photo</a>
                    {/* eslint-disable-next-line @next/next/no-img-element */}
                    <img src={`/api/gasds/receipts/${r.id}/photo`} alt={`Count sheet for ${showNo(r.no)}`} className="mt-1 hidden max-h-64 rounded border border-line print:block" /></div>}
                </td>
                <td className={td}>
                  {r.slip ?? (r.gift ? 'Banked' : '—')}
                  {r.count && r.count.n50 > 0 && <div className="text-xs text-muted">{gbp(r.count.n50 * 50)} in £50 notes taken off</div>}
                  {r.counted != null && p(r.counted) !== p(r.amount) && <div className="text-xs text-muted">bank counted {gbp(r.counted)}</div>}
                </td>
                <td className={tdr}>{gbp(r.amount)}</td>
                <td className={`${tdr} font-bold`}>{gbp(r.used ?? 0)}</td>
              </tr>
            ))}
            <tr className="font-bold"><td className={td} colSpan={7}>Total</td><td className={tdr}>{gbp(c.amount)}</td></tr>
          </tbody>
        </table>
      </Card>

      <p className="text-sm text-muted">{RULES_TEXT}</p>
    </div>
  )
}
