'use client'

import { useEffect, useState } from 'react'
import Link from 'next/link'
import { StatTile } from '@/components/ui/primitives'
import { Skeleton } from '@/components/ui/skeleton'
import { useCharity } from '@/lib/store/charity'
import { apiJson, gbp, type Status } from './shared'
import { giftAidOn } from '@/lib/claims'

type Summary = Record<Status, { count: number; amount: number; paid?: number }>

// Order and colours follow the design: money in (blue), waiting (amber), ready to go (green)
const TILES: { status: Status; label: string; hint: string; tone: 'warn' | 'info' | 'good' }[] = [
  { status: 'received', label: 'Received from HMRC', hint: 'banked', tone: 'info' },
  { status: 'submitted', label: 'Waiting for HMRC', hint: 'submitted, not yet paid', tone: 'warn' },
  { status: 'pending', label: 'Ready to submit', hint: 'download and upload to HMRC', tone: 'good' },
]

/** Claim counts and values by status. Quietly hides itself if the database isn't reachable. */
export default function SummaryStrip({ refreshKey = 0, linked = true }: { refreshKey?: number; linked?: boolean }) {
  const { loaded, charities, selectedId } = useCharity()
  // Remember which charity the numbers belong to, so a switch shows placeholders instead of stale totals
  const [data, setData] = useState<{ scope: string; summary: Summary } | null>(null)
  const [failed, setFailed] = useState(false)
  const hasCharities = charities.length > 0
  const summary = data && data.scope === selectedId ? data.summary : null

  useEffect(() => {
    if (!loaded || !hasCharities) return
    let cancelled = false
    apiJson<{ summary: Summary }>(`/api/claims?summary=1&charity=${selectedId}`)
      .then((d) => !cancelled && setData({ scope: selectedId, summary: d.summary }))
      .catch(() => !cancelled && setFailed(true))
    return () => {
      cancelled = true
    }
  }, [refreshKey, loaded, hasCharities, selectedId])

  if (failed || (loaded && !hasCharities)) return null

  return (
    <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
      {TILES.map(({ status, label, hint, tone }) => {
        const s = summary?.[status]
        const tile = (
          <StatTile
            label={label}
            tone={tone}
            value={s ? gbp(status === 'received' && s.paid !== undefined ? s.paid : giftAidOn(s.amount)) : <Skeleton className="h-10 w-40" />}
            caption={
              s ? (
                `${s.count.toLocaleString()} claim${s.count === 1 ? '' : 's'} · ${hint}`
              ) : (
                <span className="inline-flex items-center gap-2">
                  <Skeleton className="h-4 w-16" /> {hint}
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
