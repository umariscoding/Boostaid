'use client'

import { createContext, useCallback, useContext, useEffect, useMemo, useState } from 'react'
import { apiJson } from '@/components/claims/shared'

export interface CharityInfo {
  id: string
  name: string
  createdAt: string
  claimCount: number
  donationCount: number
  donated: number
}

interface CharityState {
  charities: CharityInfo[]
  /** True once the charity list (and the remembered selection) has been read. */
  loaded: boolean
  /** "all" or a charity id. Views (Claims, Search, tiles) follow this. */
  selectedId: string
  selected: CharityInfo | null
  select: (id: string) => void
  refresh: () => Promise<CharityInfo[]>
}

const STORAGE_KEY = 'boostaid.charity'
const Ctx = createContext<CharityState | null>(null)

export function useCharity(): CharityState {
  const ctx = useContext(Ctx)
  if (!ctx) throw new Error('useCharity must be used inside <CharityProvider>')
  return ctx
}

export default function CharityProvider({ children }: { children: React.ReactNode }) {
  const [charities, setCharities] = useState<CharityInfo[]>([])
  const [loaded, setLoaded] = useState(false)
  const [selectedId, setSelectedId] = useState('all')

  const refresh = useCallback(async () => {
    const data = await apiJson<{ charities: CharityInfo[] }>('/api/charities')
    setCharities(data.charities)
    // A remembered or just-deleted charity that no longer exists falls back to "all"
    setSelectedId((cur) => (cur !== 'all' && !data.charities.some((c) => c.id === cur) ? 'all' : cur))
    return data.charities
  }, [])

  // Initial load: fetch the list, then restore the remembered selection (only if that charity still exists)
  useEffect(() => {
    let cancelled = false
    apiJson<{ charities: CharityInfo[] }>('/api/charities')
      .then((data) => {
        if (cancelled) return
        setCharities(data.charities)
        try {
          const saved = localStorage.getItem(STORAGE_KEY)
          if (saved && data.charities.some((c) => c.id === saved)) setSelectedId(saved)
        } catch {
          // Storage can be unavailable (private window); selection just won't be remembered
        }
      })
      .catch(() => {})
      .finally(() => {
        if (!cancelled) setLoaded(true)
      })
    return () => {
      cancelled = true
    }
  }, [])

  const select = useCallback((id: string) => {
    setSelectedId(id)
    try {
      localStorage.setItem(STORAGE_KEY, id)
    } catch {
      // ignore
    }
  }, [])

  const value = useMemo<CharityState>(
    () => ({
      charities,
      loaded,
      selectedId,
      selected: charities.find((c) => c.id === selectedId) ?? null,
      select,
      refresh,
    }),
    [charities, loaded, selectedId, select, refresh]
  )

  return <Ctx.Provider value={value}>{children}</Ctx.Provider>
}
