'use client'

import { useEffect, useState } from 'react'
import Link from 'next/link'
import { CheckCircle2, Clock, Send } from 'lucide-react'
import { apiJson, gbp, type Status } from './shared'

type Summary = Record<Status, { count: number; amount: number }>

const TILES: { status: Status; label: string; hint: string; icon: typeof Clock; tone: string; ring: string }[] = [
  { status: 'pending', label: 'Pending', hint: 'Created, not yet sent to HMRC', icon: Clock, tone: 'text-amber-200', ring: 'border-amber-500/30 bg-amber-500/10' },
  { status: 'submitted', label: 'Submitted', hint: 'Sent to HMRC, awaiting payment', icon: Send, tone: 'text-sky-200', ring: 'border-sky-500/30 bg-sky-500/10' },
  { status: 'received', label: 'Received', hint: 'Money banked', icon: CheckCircle2, tone: 'text-emerald-200', ring: 'border-emerald-500/30 bg-emerald-500/10' },
]

/** Claim counts and totals by status. Quietly hides itself if the database isn't reachable. */
export default function SummaryStrip({ refreshKey = 0 }: { refreshKey?: number }) {
  const [summary, setSummary] = useState<Summary | null>(null)
  const [failed, setFailed] = useState(false)

  useEffect(() => {
    let cancelled = false
    apiJson<{ summary: Summary }>('/api/claims?summary=1')
      .then((d) => !cancelled && setSummary(d.summary))
      .catch(() => !cancelled && setFailed(true))
    return () => {
      cancelled = true
    }
  }, [refreshKey])

  if (failed) return null

  return (
    <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
      {TILES.map(({ status, label, hint, icon: Icon, tone, ring }) => (
        <Link
          key={status}
          href="/claims"
          className={`rounded-2xl border p-4 transition-colors hover:brightness-125 ${ring}`}
        >
          <div className="flex items-center justify-between">
            <div className="text-xs uppercase tracking-wide text-slate-300">{label} claims</div>
            <Icon className={`w-4 h-4 ${tone}`} />
          </div>
          <div className="mt-2 flex items-baseline gap-3">
            <div className={`text-3xl font-bold ${tone}`}>{summary ? summary[status].count : '-'}</div>
            <div className="text-sm text-slate-300">{summary ? gbp(summary[status].amount) : ' '}</div>
          </div>
          <div className="mt-1 text-xs text-slate-400">{hint}</div>
        </Link>
      ))}
    </div>
  )
}
