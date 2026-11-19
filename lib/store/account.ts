'use client'

import { create } from 'zustand'
import { apiJson } from '@/components/claims/shared'

interface AccountStore {
  email: string | null
  name: string
  /** 'admin' = BoostAid's owner (sees the Admin page). */
  role: string
  init: () => Promise<void>
}

let started = false

/** The signed-in person (shown in the sidebar and under Settings -> Account). */
export const useAccountStore = create<AccountStore>((set) => ({
  email: null,
  name: '',
  role: 'user',
  init: async () => {
    if (started) return
    started = true
    try {
      const me = await apiJson<{ email: string; name: string; role: string }>('/api/auth/me')
      set({ email: me.email, name: me.name, role: me.role })
    } catch {
      // apiJson already redirects to /login when the session is no longer valid
    }
  },
}))
