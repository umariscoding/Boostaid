import { Suspense } from 'react'
import ClaimBuilder from '@/components/claim-builder'

export default function Home() {
  // useSearchParams (for /?recovered=1) needs a Suspense boundary
  return (
    <Suspense>
      <ClaimBuilder />
    </Suspense>
  )
}
