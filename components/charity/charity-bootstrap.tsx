'use client'

import { useEffect } from 'react'
import { useCharityStore } from '@/lib/store/charity'

/** Loads the charity list into the global store once when the app shell mounts. */
export default function CharityBootstrap() {
  useEffect(() => {
    useCharityStore.getState().init()
  }, [])
  return null
}
