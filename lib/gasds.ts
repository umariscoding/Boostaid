/**
 * GASDS (Gift Aid Small Donations Scheme): the rules, with no database and no UI.
 * Source: HMRC "Charities: detailed guidance notes", chapter 8, as set out in INSTRUCTIONS_GASDS.txt.
 *
 *  - A receipt counts if it is cash or card, has no Gift Aid declaration, is banked, is not void,
 *    and is £30 or less when it is one person's gift (a collection can be any size).
 *  - Its amount is capped at what the bank counted, less £50 x the £50 notes (HMRC doesn't allow them).
 *  - Tax year 6 April - 5 April, claimable until 5 April two years after it ends.
 *  - Limit per charity per tax year = lowest of £8,000 and 10 x the Gift Aid donations of that year,
 *    less what is already claimed. HMRC pays 25%. Oldest receipts first; the last may be part used.
 *
 * All money is worked out in whole pence so there are no rounding surprises.
 */

export const GASDS_LIMIT = 8000
export const GASDS_RATE = 0.25
export const GASDS_MATCH = 10
export const GIFT_MAX = 30

export interface Count { n50: number; n20: number; n10: number; n5: number; coins: number }

/** One receipt (or collection) as stored. */
export interface GReceipt {
  id?: string // database id (stored receipts only)
  no: string
  slip?: string | null
  date: string // yyyy-MM-dd
  year: number // tax year it falls in (the year it starts, 2026 = 2026/27)
  amount: number
  card: boolean
  notCash: boolean // cheque, bank transfer, online ...
  notes: string
  donor: string // '' = anonymous
  gift: boolean // one person's gift (true) or a collection (false)
  giftAid: boolean
  banked: boolean
  isVoid: boolean
  counted?: number | null // what the bank counted in the envelope
  count?: Count | null // notes and coins counted by the charity
  photo?: string | null // file name of an attached count sheet
  claimRef?: string | null // set once the receipt is in a saved claim
  used?: number | null // the amount used in that claim
}

export type Group = 'in' | 'wait' | 'out'
export interface Status { group: Group; label: string; reason: string }

// ---------------------------------------------------------------- small helpers
export const p = (n: number) => Math.round(n * 100)
export const gbp = (n: number) =>
  '£' + (Number(n) || 0).toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
export const topUpOf = (amount: number) => Math.round(p(amount) * GASDS_RATE) / 100
export const ukDate = (iso: string) => (iso ? iso.split('-').reverse().join('/') : '')

export function taxYearOf(iso: string): number {
  const [y, m, d] = iso.split('-').map(Number)
  return m > 4 || (m === 4 && d >= 6) ? y : y - 1
}
export const taxYearLabel = (y: number) => `${y}/${String((y + 1) % 100).padStart(2, '0')}`
/** Claims for a tax year close on 5 April, two years after it ends. */
export const deadlineOf = (y: number) => `${y + 3}-04-05`
export const longDate = (iso: string) =>
  new Date(iso + 'T12:00:00Z').toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' })

/** The tax years that can still be claimed on `today` (newest first). */
export function openYears(today: string): number[] {
  const cur = taxYearOf(today)
  return [cur, cur - 1, cur - 2].filter((y) => deadlineOf(y) >= today)
}
export const daysLeft = (y: number, today: string) =>
  Math.round((Date.parse(deadlineOf(y) + 'T00:00:00Z') - Date.parse(today + 'T00:00:00Z')) / 86400000)

export const taxYearRange = (y: number) => ({ from: `${y}-04-06`, to: `${y + 1}-04-05` })

// ---------------------------------------------------------------- one receipt
/** Not included for a reason that no upload or banking can change. '' = eligible so far. */
export function notIncludedWhy(r: GReceipt): string {
  if (r.isVoid) return 'cancelled receipt'
  if (r.notCash) return 'paid by cheque or bank transfer'
  if (r.giftAid) return 'has Gift Aid, so it goes in the Gift Aid claim'
  if (r.gift && r.amount > GIFT_MAX) return 'over £30 from one person'
  return ''
}

export function countTotal(c: Count): number {
  return (p(c.n50 * 50) + p(c.n20 * 20) + p(c.n10 * 10) + p(c.n5 * 5) + p(c.coins)) / 100
}

/** How much of one receipt counts towards a claim (before the limit). 0 if it can't count. */
export function counts(r: GReceipt): number {
  if (notIncludedWhy(r) || !r.banked) return 0
  const base = r.counted != null ? Math.min(p(r.amount), p(r.counted)) : p(r.amount)
  return r.count && !r.gift ? Math.max(0, base - p(r.count.n50 * 50)) / 100 : base / 100
}

/** Cash collections need a count as proof. Single gifts and card collections don't. */
export const needsCount = (r: GReceipt) => !r.gift && !r.card
export const hasCount = (r: GReceipt) => !needsCount(r) || !!(r.count || r.photo)
/** Receipts with no number are given a stable one from their date, amount and notes. */
export const showNo = (no: string) => (/^T \d{4}-/.test(no) ? '—' : no)
export const who = (r: GReceipt) => (r.gift ? r.donor || 'Anonymous donor' : r.notes || 'Collection')

// ---------------------------------------------------------------- the claim for one tax year
export interface Analysis {
  year: number
  claimed: number // already claimed this year (saved claims)
  open: GReceipt[] // receipts that count and are not in a claim yet
  openAmount: number
  giftAid: number // Gift Aid donations for the year (claimed in BoostAid + outside)
  matchLimit: number // 10 x giftAid
  cap: number // lowest of £8,000 and the matching limit
  room: number // cap less claimed (>= 0)
  amount: number // what to type into HMRC now
  topUp: number // what HMRC pays on it
  overBy: number // eligible money that doesn't fit under the limit
  matchingBites: boolean // the Gift Aid rule is what is holding the amount down
  plan: { r: GReceipt; used: number }[] // receipts making up this claim, oldest first
  collections: { r: GReceipt; used: number }[] // those that need a count
  statuses: Map<string, Status>
  waiting: GReceipt[] // not banked yet (and otherwise fine)
  extra: number // money that banking the waiting receipts would add to the claim
  maybeGiftAid: GReceipt[] // over £30 from a donor who gave Gift Aid on another receipt (any payment method)
}

const bareName = (n: string) => n.toLowerCase().replace(/^(mr|mrs|miss|ms|dr)\.?\s+/, '').replace(/\s+/g, ' ').trim()
const byDate = (a: GReceipt, b: GReceipt) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.no < b.no ? -1 : a.no > b.no ? 1 : 0)

/**
 * `receipts` = every stored receipt of the charity for that tax year.
 * `claimedAmounts` = the amounts of the claims already saved for the year.
 * `giftAid` = Gift Aid donations for the year (claimed in BoostAid + claimed outside BoostAid).
 */
export function analyse(receipts: GReceipt[], claimedAmounts: number[], giftAid: number, year: number): Analysis {
  const claimedP = claimedAmounts.reduce((s, a) => s + p(a), 0)
  const open = receipts.filter((r) => !r.claimRef && counts(r) > 0)
  const openP = open.reduce((s, r) => s + p(counts(r)), 0)
  const matchP = p(giftAid) * GASDS_MATCH
  const capP = Math.min(p(GASDS_LIMIT), matchP)
  const roomP = Math.max(0, capP - claimedP)
  const amountP = Math.min(openP, roomP)

  // Oldest first, up to the amount. The last receipt can be part used.
  let left = amountP
  const plan: { r: GReceipt; used: number }[] = []
  for (const r of [...open].sort(byDate)) {
    if (left <= 0) break
    const use = Math.min(p(counts(r)), left)
    plan.push({ r, used: use / 100 })
    left -= use
  }
  const inPlan = new Set(plan.map((x) => x.r.no))

  // Why the year is full: the £8,000 cap, or the 10x Gift Aid rule when that is the lower one
  const limitWhy = matchP < p(GASDS_LIMIT) ? 'not enough Gift Aid claimed for the year (HMRC pays on up to 10 times it)' : 'the £8,000 limit for the year is reached'
  const statuses = new Map<string, Status>()
  for (const r of receipts) {
    let s: Status
    const why = notIncludedWhy(r)
    if (r.claimRef) {
      s = { group: 'in', label: 'Claimed', reason: `HMRC ref ${r.claimRef}${p(r.used ?? 0) < p(counts({ ...r, claimRef: null })) ? ', part of it' : ''}` }
    } else if (inPlan.has(r.no)) s = { group: 'in', label: 'In this claim', reason: '' }
    else if (why) s = { group: 'out', label: 'Not included', reason: why }
    // Unbanked receipts keep waiting until claims have actually filled the year (then they can never be added)
    else if (!r.banked) {
      s = roomP <= 0
        ? { group: 'out', label: 'Not included', reason: limitWhy }
        : { group: 'wait', label: 'Waiting', reason: 'not banked yet' }
    } else if (roomP - amountP <= 0) s = { group: 'out', label: 'Not included', reason: limitWhy }
    else s = { group: 'in', label: 'Included', reason: '' }
    statuses.set(r.no, s)
  }

  const waiting = receipts.filter((r) => !r.claimRef && !r.banked && !notIncludedWhy(r))
  const waitingP = waiting.reduce((s, r) => s + p(r.amount), 0)
  const extraP = Math.min(waitingP, Math.max(0, roomP - amountP))

  const gaNames = new Set(receipts.filter((r) => r.giftAid && r.donor).map((r) => bareName(r.donor)))
  const maybeGiftAid = receipts.filter(
    (r) => r.gift && r.amount > GIFT_MAX && !r.giftAid && !r.isVoid && r.donor && gaNames.has(bareName(r.donor))
  )

  return {
    year,
    claimed: claimedP / 100,
    open,
    openAmount: openP / 100,
    giftAid,
    matchLimit: matchP / 100,
    cap: capP / 100,
    room: roomP / 100,
    amount: amountP / 100,
    topUp: topUpOf(amountP / 100),
    overBy: Math.max(0, openP - roomP) / 100,
    matchingBites: matchP < p(GASDS_LIMIT) && openP + claimedP > matchP,
    plan,
    collections: plan.filter((x) => needsCount(x.r)),
    statuses,
    waiting,
    extra: extraP / 100,
    maybeGiftAid,
  }
}

// ---------------------------------------------------------------- uploads
export interface BankingFile { slip: string; items: { no: string; amount: number; counted: number }[] }

/** What happens to one uploaded receipt, given what is stored and what the banking files say. */
export interface UploadOutcome {
  receipt: GReceipt // the receipt as it would be stored (banking applied)
  action: 'new' | 'update' | 'same' | 'claimed' | 'too-old'
  status: Status
  used: number // pence-exact amount that would be added to the claim from this upload
}

/** Room left per open year before an upload: the limit less claimed and less already waiting to be claimed. */
export function roomByYear(
  stored: GReceipt[],
  claimedByYear: Record<number, number>,
  giftAidByYear: Record<number, number>,
  today: string
): Record<number, number> {
  const room: Record<number, number> = {}
  for (const y of openYears(today)) {
    let used = p(claimedByYear[y] ?? 0)
    used += stored.filter((r) => r.year === y && !r.claimRef).reduce((s, r) => s + p(counts(r)), 0)
    room[y] = Math.max(0, Math.min(p(GASDS_LIMIT), p(giftAidByYear[y] ?? 0) * GASDS_MATCH) - used)
  }
  return room
}

export const changed = (prev: GReceipt, next: GReceipt) =>
  prev.banked !== next.banked ||
  prev.giftAid !== next.giftAid ||
  p(prev.amount) !== p(next.amount) ||
  (!!next.slip && prev.slip !== next.slip) ||
  (next.counted != null && p(prev.counted ?? -1) !== p(next.counted))

/** Applies the banking files to the uploaded receipts: banked, with the slip and the counted amount. */
export function applyBanking(uploaded: GReceipt[], banking: BankingFile[]): GReceipt[] {
  const by = new Map<string, { slip: string; counted: number }>()
  for (const b of banking) for (const it of b.items) by.set(it.no, { slip: b.slip, counted: it.counted })
  return uploaded.map((r) => {
    const b = by.get(r.no)
    return b ? { ...r, banked: true, slip: b.slip, counted: b.counted } : r
  })
}

/** Previews an upload: what each receipt would become, and how much it adds to the claim. */
export function previewUpload(
  stored: GReceipt[],
  uploaded: GReceipt[],
  banking: BankingFile[],
  claimedByYear: Record<number, number>,
  giftAidByYear: Record<number, number>,
  today: string
): UploadOutcome[] {
  const have = new Map(stored.map((r) => [r.no, r]))
  const room = roomByYear(stored, claimedByYear, giftAidByYear, today)
  const open = new Set(openYears(today))
  return applyBanking(uploaded, banking).map((r): UploadOutcome => {
    const prev = have.get(r.no)
    if (!open.has(r.year)) return { receipt: r, action: 'too-old', status: { group: 'out', label: 'Not included', reason: 'too old to claim' }, used: 0 }
    if (prev?.claimRef) return { receipt: r, action: 'claimed', status: { group: 'out', label: 'Already claimed', reason: `HMRC ref ${prev.claimRef}` }, used: 0 }
    if (prev && !changed(prev, r)) return { receipt: r, action: 'same', status: { group: 'out', label: 'Already uploaded', reason: '' }, used: 0 }
    const action = prev ? 'update' : 'new'
    const why = notIncludedWhy(r)
    if (why) return { receipt: r, action, status: { group: 'out', label: 'Not included', reason: why }, used: 0 }
    if (!r.banked) return { receipt: r, action, status: { group: 'wait', label: 'Waiting', reason: 'not banked yet' }, used: 0 }
    // An update of a receipt already waiting in the plan replaces what it counted before
    const before = prev && !prev.claimRef ? p(counts(prev)) : 0
    if (before) room[r.year] = (room[r.year] ?? 0) + before
    const c = p(counts(r))
    const fit = Math.min(c, room[r.year] ?? 0)
    room[r.year] = (room[r.year] ?? 0) - fit
    if (fit <= 0) {
      const lowGiftAid = p(giftAidByYear[r.year] ?? 0) * GASDS_MATCH < p(GASDS_LIMIT)
      return {
        receipt: r,
        action,
        status: { group: 'out', label: 'Not included', reason: lowGiftAid ? `not enough Gift Aid claimed for ${taxYearLabel(r.year)}` : `${taxYearLabel(r.year)} has reached its limit` },
        used: 0,
      }
    }
    return { receipt: r, action, status: { group: 'in', label: 'Included', reason: fit < c ? 'part of it (the limit is reached)' : '' }, used: fit / 100 }
  })
}

/** How a receipt's count is shown in the evidence (page, Excel and PDF). */
export function countText(r: GReceipt): string {
  if (r.gift) return 'Single gift, the receipt is the record'
  if (r.card) return 'Card payments, no notes'
  if (!r.count) return r.photo ? `Photo of count sheet: ${r.photo}` : 'Not recorded'
  const c = r.count
  const parts: string[] = []
  for (const [k, v] of [['n50', 50], ['n20', 20], ['n10', 10], ['n5', 5]] as const) if (c[k]) parts.push(`${c[k]} × £${v}`)
  if (c.coins) parts.push(`${gbp(c.coins)} coins`)
  return `${parts.join(', ')} = ${gbp(countTotal(c))}${r.photo ? ' (+ photo)' : ''}`
}
