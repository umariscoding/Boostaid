import type { claims } from '@/lib/db'

export type ClaimRow = typeof claims.$inferSelect

export interface ClaimDto {
  id: string
  ref: string
  hmrcRef: string | null
  status: 'pending' | 'submitted' | 'received'
  donationCount: number
  /** Sum of the donations in the claim (the schedule's "Total donations" box). */
  totalAmount: number
  /** What HMRC pays on those donations: 25%. */
  giftAidAmount: number
  notes: string | null
  createdAt: string
  submittedAt: string | null
  receivedAt: string | null
  chunkIndex: number
  sourceFile?: string
  /** Earliest / latest donation date in the claim (yyyy-MM-dd). */
  periodFrom?: string | null
  periodTo?: string | null
}

export const GIFT_AID_RATE = 0.25
export const giftAidOn = (donations: number) => Math.round(donations * GIFT_AID_RATE * 100) / 100

export function toClaimDto(row: ClaimRow, sourceFile?: string, period?: { from: string | null; to: string | null }): ClaimDto {
  return {
    id: row.id,
    ref: row.ref,
    hmrcRef: row.hmrcRef,
    status: row.status,
    donationCount: row.donationCount,
    totalAmount: Number(row.totalAmount),
    giftAidAmount: giftAidOn(Number(row.totalAmount)),
    notes: row.notes,
    createdAt: row.createdAt.toISOString(),
    submittedAt: row.submittedAt?.toISOString() ?? null,
    receivedAt: row.receivedAt?.toISOString() ?? null,
    chunkIndex: row.chunkIndex,
    sourceFile,
    periodFrom: period?.from ?? null,
    periodTo: period?.to ?? null,
  }
}

/** dd/MM/yyyy -> yyyy-MM-dd, or null if it isn't a real date. */
export function ukDateToIso(value: unknown): string | null {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(value ?? '').trim())
  if (!m) return null
  const [, d, mo, y] = m
  const dt = new Date(Date.UTC(Number(y), Number(mo) - 1, Number(d)))
  if (dt.getUTCFullYear() !== Number(y) || dt.getUTCMonth() !== Number(mo) - 1 || dt.getUTCDate() !== Number(d)) {
    return null
  }
  return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}`
}

export function normalisePostcode(value: string): string {
  return value.replace(/\s+/g, '').toUpperCase()
}

export function escapeLike(value: string): string {
  return value.replace(/[\\%_]/g, (c) => `\\${c}`)
}
