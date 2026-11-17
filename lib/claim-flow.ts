import type { Table } from '@/lib/giftaid/read'
import { buildClaim, type ClaimResult } from '@/lib/giftaid'

export interface Upload {
  name: string
  data: Uint8Array
}

export const tableKey = (t: Pick<Table, 'fileName' | 'sheetName'>) => `${t.fileName}:${t.sheetName}`

const text = (v: unknown) => (v === null || v === undefined ? '' : String(v).trim())
const isEmptyRow = (r: unknown[]) => !r.some((v) => v !== null && v !== undefined && v !== '')

/** A sheet as header -> value records (what the column-mapping screen shows). */
export function tableRecords(t: Table): Record<string, unknown>[] {
  const headings = (t.rows[0] ?? []).map(text)
  return t.rows.slice(1).filter((r) => !isEmptyRow(r)).map((r) => Object.fromEntries(headings.map((h, i) => [h || `Column ${i + 1}`, r[i] ?? null])))
}

/**
 * Takes donations out of a claim without moving anything: the row is emptied, so every other row keeps its
 * original row number (which the review workbook quotes as the "source").
 * `sources` look like "file.xlsx:Sheet1!57".
 */
export function blankSources(tables: Table[], sources: string[]): Table[] {
  const byTable = new Map<string, Set<number>>()
  for (const src of sources) {
    const bang = src.lastIndexOf('!')
    const key = src.slice(0, bang)
    const set = byTable.get(key) ?? new Set<number>()
    set.add(Number(src.slice(bang + 1)) - 1) // "Sheet!57" is the 57th spreadsheet row = rows[56]
    byTable.set(key, set)
  }
  return tables.map((t) => {
    const drop = byTable.get(tableKey(t))
    if (!drop) return t
    return { ...t, rows: t.rows.map((r, i) => (drop.has(i) ? r.map(() => null) : r)) }
  })
}

export const CASH_BOOK_HEADINGS = ['Title', 'First name', 'Last name', 'House name or number', 'Postcode', 'Donation date', 'Amount']

export interface MappingLike {
  [target: string]: string | string[] | null
}

/**
 * A sheet the engine did not recognise, re-laid-out with HMRC's own cash-book headings so the normal cleaning
 * rules apply to it. Only rows whose Gift Aid value was ticked are kept; positions are preserved.
 */
export function mapToCashBook(t: Table, mapping: MappingLike, giftAid: { column: string | null; valuesToKeep: string[] }): Table {
  const headings = (t.rows[0] ?? []).map(text)
  const idx = (name: string | string[] | null | undefined): number[] =>
    (Array.isArray(name) ? name : name ? [name] : []).map((n) => headings.indexOf(n)).filter((i) => i >= 0)
  const one = (key: string) => idx(mapping[key])[0] ?? -1
  const iTitle = one('Title'), iFirst = one('First Name'), iLast = one('Last Name'), iPost = one('Postcode'), iDate = one('Date'), iAmt = one('Amount')
  const iAddr = idx(mapping['Address'])
  const iGift = giftAid.column ? headings.indexOf(giftAid.column) : -1
  const keep = new Set(giftAid.valuesToKeep)

  const rows: unknown[][] = [CASH_BOOK_HEADINGS]
  for (const r of t.rows.slice(1)) {
    const keepRow = !isEmptyRow(r) && (iGift < 0 || keep.has(text(r[iGift])))
    if (!keepRow) {
      rows.push(CASH_BOOK_HEADINGS.map(() => null))
      continue
    }
    const get = (i: number) => (i >= 0 ? r[i] ?? null : null)
    rows.push([
      get(iTitle),
      get(iFirst),
      get(iLast),
      iAddr.map((i) => text(r[i])).filter(Boolean).join(', ') || null,
      get(iPost),
      get(iDate),
      get(iAmt),
    ])
  }
  return { fileName: t.fileName, sheetName: `${t.sheetName} (mapped)`, rows }
}

export interface AlreadyClaimed {
  platform: string
  donationId: string
  claimRef: string
}

/** Asks the server which of these (platform, donation ID) pairs are already in a claim for the charity. */
export type ClaimedLookup = (items: { platform: string; donationId: string }[]) => Promise<AlreadyClaimed[]>

/**
 * Cleans the tables and takes out every donation that is already claimed.
 *
 * Taking a row out can bring back a copy of it: some exports (e.g. LaunchGood's "Sheet1") repeat rows, and the
 * engine skips a repeat only while the original is still there. So this keeps checking and taking out until the
 * claim holds nothing already claimed. Each donation ID is asked about once.
 */
export async function withoutAlreadyClaimed(
  tables: Table[],
  lookup: ClaimedLookup,
  maxPasses = 10
): Promise<{ result: ClaimResult; blocked: (AlreadyClaimed & { name: string; amount: number })[] }> {
  const key = (platform: string, id: unknown) => `${platform}\u0000${String(id)}`
  const hasId = (r: { donationId: unknown }) => r.donationId !== null && r.donationId !== undefined && String(r.donationId) !== ''
  const asked = new Set<string>()
  const claimedRef = new Map<string, string>()
  const blocked = new Map<string, AlreadyClaimed & { name: string; amount: number }>()

  let current = tables
  let result = buildClaim(current)
  for (let pass = 0; pass < maxPasses; pass++) {
    const withIds = [...result.ready, ...result.overseas].filter(hasId)
    const toAsk = withIds.filter((r) => !asked.has(key(r.platform, r.donationId)))
    for (let i = 0; i < toAsk.length; i += 20000) {
      const chunk = toAsk.slice(i, i + 20000).map((r) => ({ platform: r.platform, donationId: String(r.donationId) }))
      chunk.forEach((c) => asked.add(key(c.platform, c.donationId)))
      for (const c of await lookup(chunk)) claimedRef.set(key(c.platform, c.donationId), c.claimRef)
    }
    const hit = withIds.filter((r) => claimedRef.has(key(r.platform, r.donationId)))
    if (!hit.length) break
    for (const r of hit) {
      const k = key(r.platform, r.donationId)
      if (!blocked.has(k)) {
        blocked.set(k, { name: `${r.firstName} ${r.lastName}`.trim(), amount: r.amount, platform: r.platform, donationId: String(r.donationId), claimRef: claimedRef.get(k) ?? '' })
      }
    }
    current = blankSources(current, hit.map((r) => r.source))
    result = buildClaim(current)
  }
  return { result, blocked: [...blocked.values()] }
}
