import { sql } from 'drizzle-orm'
import { UUID_RE } from '@/lib/api'
import { batches, claims } from '@/lib/db'

/** `?charity=` is a charity id, or "all"/absent for every charity. Returns the id, or null for all. */
export function parseCharityParam(value: string | null): string | null {
  if (!value || value === 'all') return null
  if (!UUID_RE.test(value)) throw new RangeError('Invalid charity')
  return value
}

/** Restricts a claims query to one charity (the claims table must be in the FROM clause). */
export function claimsOfCharity(charityId: string) {
  return sql`${claims.batchId} in (select ${batches.id} from ${batches} where ${batches.charityId} = ${charityId})`
}
