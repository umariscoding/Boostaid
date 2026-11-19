import { and, eq, inArray } from 'drizzle-orm'
import { batches, claims, donations, getDb } from '@/lib/db'

type Db = Pick<ReturnType<typeof getDb>, 'select'>

export interface DonationKey {
  platform: string
  donationId: string
}

export interface AlreadyClaimed extends DonationKey {
  claimRef: string
}

const key = (k: DonationKey) => `${k.platform}\u0000${k.donationId}`

/**
 * Donations (same platform + platform donation ID) that are already in a claim for this charity.
 * A donor's donation must only ever be claimed once.
 */
export async function findAlreadyClaimed(db: Db, charityId: string, items: DonationKey[]): Promise<AlreadyClaimed[]> {
  const wanted = new Set(items.filter((i) => i.donationId).map(key))
  if (wanted.size === 0) return []
  const ids = [...new Set(items.filter((i) => i.donationId).map((i) => i.donationId))]

  const found: AlreadyClaimed[] = []
  const seen = new Set<string>()
  for (let i = 0; i < ids.length; i += 500) {
    const rows = await db
      .select({ platform: donations.platform, donationId: donations.platformDonationId, claimRef: claims.ref })
      .from(donations)
      .innerJoin(claims, eq(claims.id, donations.claimId))
      .innerJoin(batches, eq(batches.id, claims.batchId))
      .where(and(inArray(donations.platformDonationId, ids.slice(i, i + 500)), eq(batches.charityId, charityId)))
    for (const r of rows) {
      const k = { platform: r.platform, donationId: r.donationId ?? '' }
      if (wanted.has(key(k)) && !seen.has(key(k))) {
        seen.add(key(k))
        found.push({ ...k, claimRef: r.claimRef })
      }
    }
  }
  return found
}
