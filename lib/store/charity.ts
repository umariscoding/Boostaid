'use client'

import { create } from 'zustand'
import { useShallow } from 'zustand/react/shallow'
import { apiJson } from '@/components/claims/shared'

export interface CharityInfo {
  id: string
  name: string
  createdAt: string
  claimCount: number
  donationCount: number
  donated: number
}

interface CharityStore {
  charities: CharityInfo[]
  /** True once the charity list (and the remembered selection) has been read. */
  loaded: boolean
  /** The sidebar selection: "all" or a charity id. Views (Claims, Search, tiles) follow it. */
  selectedId: string
  /** Only used while "all" is selected and there are several charities: the one chosen for an upload.
   *  Cleared on every sidebar change and after each finished upload, so it is always chosen deliberately. */
  allModeTargetId: string

  init: () => Promise<void>
  refresh: () => Promise<CharityInfo[]>
  select: (id: string) => void
  chooseUploadTarget: (id: string) => void
}

const STORAGE_KEY = 'boostaid.charity'
let initPromise: Promise<void> | null = null

export const useCharityStore = create<CharityStore>((set, get) => ({
  charities: [],
  loaded: false,
  selectedId: 'all',
  allModeTargetId: '',

  refresh: async () => {
    const { charities } = await apiJson<{ charities: CharityInfo[] }>('/api/charities')
    // A selection that no longer exists (e.g. the charity was just deleted) falls back to "all"
    set((s) => ({
      charities,
      selectedId: s.selectedId !== 'all' && !charities.some((c) => c.id === s.selectedId) ? 'all' : s.selectedId,
      allModeTargetId: charities.some((c) => c.id === s.allModeTargetId) ? s.allModeTargetId : '',
    }))
    return charities
  },

  /** Loads the list once and restores the remembered selection. Safe to call repeatedly. */
  init: () => {
    if (!initPromise) {
      initPromise = (async () => {
        try {
          const list = await get().refresh()
          const saved = localStorage.getItem(STORAGE_KEY)
          if (saved && list.some((c) => c.id === saved)) set({ selectedId: saved })
        } catch {
          // List failed or storage is unavailable: start on "all charities"
        } finally {
          set({ loaded: true })
        }
      })()
    }
    return initPromise
  },

  select: (id) => {
    // Switching charity always clears the "all charities" upload choice, so it must be made again
    set({ selectedId: id, allModeTargetId: '' })
    try {
      localStorage.setItem(STORAGE_KEY, id)
    } catch {
      // Selection just won't be remembered
    }
  },

  chooseUploadTarget: (id) => set({ allModeTargetId: id }),
}))

/**
 * The charity the next upload is for. Derived, never copied into other state, so it can't drift from the sidebar:
 * a specific sidebar charity wins; with "all" it's the only charity, or the one picked on the Clean page.
 */
export function resolveUploadTarget(s: Pick<CharityStore, 'charities' | 'selectedId' | 'allModeTargetId'>): string {
  if (s.selectedId !== 'all') return s.selectedId
  if (s.charities.length === 1) return s.charities[0].id
  return s.charities.some((c) => c.id === s.allModeTargetId) ? s.allModeTargetId : ''
}

export function useCharity() {
  return useCharityStore(
    useShallow((s) => ({
      charities: s.charities,
      loaded: s.loaded,
      selectedId: s.selectedId,
      selected: s.charities.find((c) => c.id === s.selectedId) ?? null,
      uploadTargetId: resolveUploadTarget(s),
      select: s.select,
      refresh: s.refresh,
      chooseUploadTarget: s.chooseUploadTarget,
    }))
  )
}
