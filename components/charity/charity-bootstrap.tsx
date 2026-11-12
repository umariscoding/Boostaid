'use client'

import { useEffect } from 'react'
import { useCharityStore } from '@/lib/store/charity'
import { useAccountStore } from '@/lib/store/account'

/** Loads the charity list and the signed-in account into their global stores once when the app shell mounts. */
export default function CharityBootstrap() {
  useEffect(() => {
    useCharityStore.getState().init()
    useAccountStore.getState().init()
  }, [])
  return null
}
