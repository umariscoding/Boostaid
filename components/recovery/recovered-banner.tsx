'use client'

import { useEffect, useState } from 'react'
import Link from 'next/link'
import { apiJson, gbp } from '@/components/claims/shared'

/** "£X recovered and waiting. Add it to your next claim?" on the Claims page. */
export default function RecoveredBanner({ charityId, refreshKey = 0 }: { charityId: string; refreshKey?: number }) {
  const [t, setT] = useState<{ count: number; giftAid: number } | null>(null)
  useEffect(() => {
    let off = false
    apiJson<{ tiles: { recovered: { count: number; giftAid: number } } }>(`/api/recovery?charity=${charityId}`)
      .then((d) => !off && setT(d.tiles.recovered))
      .catch(() => !off && setT(null))
    return () => { off = true }
  }, [charityId, refreshKey])
  if (!t || t.count === 0) return null
  return (
    <div className="mt-6 flex flex-wrap items-center justify-between gap-4 rounded-2xl border border-good-line bg-good-soft px-6 py-5 text-good">
      <div>
        <div className="font-display text-2xl font-bold tabular-nums">{gbp(t.giftAid)} recovered and waiting</div>
        <div>{t.count} donation{t.count === 1 ? '' : 's'} fixed by donors. Add {t.count === 1 ? 'it' : 'them'} to your next claim?</div>
      </div>
      <Link href="/?recovered=1" className="inline-flex h-11 items-center rounded-full bg-good px-6 text-sm font-semibold text-white hover:opacity-90">Create claim</Link>
    </div>
  )
}
