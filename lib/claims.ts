import type { claims } from '@/lib/db'

export type ClaimRow = typeof claims.$inferSelect

export interface ClaimDto {
  id: string
  ref: string
  hmrcRef: string | null
  status: 'pending' | 'submitted' | 'received'
  donationCount: number
  totalAmount: number
  notes: string | null
  createdAt: string
  submittedAt: string | null
  receivedAt: string | null
  chunkIndex: number
  sourceFile?: string
}

export function toClaimDto(row: ClaimRow, sourceFile?: string): ClaimDto {
  return {
    id: row.id,
    ref: row.ref,
    hmrcRef: row.hmrcRef,
    status: row.status,
    donationCount: row.donationCount,
    totalAmount: Number(row.totalAmount),
    notes: row.notes,
    createdAt: row.createdAt.toISOString(),
    submittedAt: row.submittedAt?.toISOString() ?? null,
    receivedAt: row.receivedAt?.toISOString() ?? null,
    chunkIndex: row.chunkIndex,
    sourceFile,
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
