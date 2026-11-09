'use client'

import { useEffect, useState } from 'react'
import Link from 'next/link'
import { StatTile } from '@/components/ui/primitives'
import { Skeleton } from '@/components/ui/skeleton'
import { apiJson, gbp, type Status } from './shared'
import { giftAidOn } from '@/lib/claims'

type Summary = Record<Status, { count: number; amount: number }>

const TILES: { status: Status; label: string; hint: string; tone: 'warn' | 'info' | 'good' }[] = [
  { status: 'pending', label: 'Pending claims', hint: 'not yet sent to HMRC', tone: 'warn' },
  { status: 'submitted', label: 'Submitted claims', hint: 'awaiting payment', tone: 'info' },
  { status: 'received', label: 'Received claims', hint: 'money banked', tone: 'good' },
]

/** Claim counts and values by status. Quietly hides itself if the database isn't reachable. */
export default function SummaryStrip({ refreshKey = 0, linked = true }: { refreshKey?: number; linked?: boolean }) {
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
    <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
      {TILES.map(({ status, label, hint, tone }) => {
        const s = summary?.[status]
        const tile = (
          <StatTile
            label={label}
            tone={tone}
            value={s ? s.count.toLocaleString() : <Skeleton className="h-12 w-16" />}
            caption={
              s ? (
                `${gbp(giftAidOn(s.amount))} Gift Aid · ${hint}`
              ) : (
                <span className="inline-flex items-center gap-2">
                  <Skeleton className="h-4 w-20" /> Gift Aid &middot; {hint}
                </span>
              )
            }
          />
        )
        return linked ? (
          <Link key={status} href="/claims" className="block transition-shadow hover:shadow-md rounded-2xl">
            {tile}
          </Link>
        ) : (
          <div key={status}>{tile}</div>
        )
      })}
    </div>
  )
}
