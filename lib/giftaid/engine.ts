/**
 * The cleaning engine: uploaded tables -> a Gift Aid claim.
 *
 *   tables ──► detect platform per sheet ──► common records ──► eligibility ──► HMRC cleaning
 *          ──► Ready (UK / overseas) · Needs Review (with reasons) · Not Eligible (with reasons)
 *          ──► plan HMRC files (max 1,000 rows each; every row knows its file + row number)
 *
 * Pure function: no files, no network, no React. Safe to run in the browser or on a server.
 * Port of boostaid_clean.py - results must match the expected outputs in the handover pack.
 */
import { s, toAscii, nonLatin, cap, fmtG, formatPostcode, UK_POSTCODE } from './text'
import { houseFrom } from './address'
import { detectDateOrder, parseDate, DateFormat } from './dates'
import { detectProfile, isJustGiving, PROFILES, SourceRecord } from './profiles'
import type { Table } from './read'

const JUNK = new Set(['anonymous', 'anon', 'anonymously', 'guest', 'user', 'test', 'x', 'xx', 'n/a', 'na', 'none', 'unknown'])
const TITLES = new Set(['mr', 'mrs', 'miss', 'ms', 'dr', 'prof'])
const UK_WORDS = new Set(['gb', 'uk', 'united kingdom', 'england', 'scotland', 'wales', 'northern ireland', 'great britain', ''])
export const MAX_ROWS_PER_HMRC_FILE = 1000

/** One donation that will go to HMRC. */
export interface ClaimRow {
  hmrcFile?: string          // e.g. "UK_schedule_3_of_8.ods"
  hmrcRow?: number           // row in that file (data starts at row 25)
  title: string
  firstName: string
  lastName: string
  house: string
  postcode: string           // "X" for overseas
  date: string | null        // YYYY-MM-DD
  amount: number             // GBP
  platform: string
  donationId: unknown
  source: string             // "file.xlsx:Sheet!row" - where it came from
  cleaned: boolean           // did BoostAid correct anything on this row?
}

export interface ReviewRow extends ClaimRow {
  reason: string             // "No house number or name; Missing first or last name"
  suggestedName: string      // billing name hint, when one exists - must be checked by a person
  originalAmount: string     // "10 USD"
  originalName: string
  originalAddress: string
  originalPostcode: string
  country: string
}

export interface IneligibleRow {
  reason: string
  name: string
  amount: number
  donationId: unknown
  source: string
}

export interface Change {
  donationId: unknown
  field: 'Name' | 'First Name' | 'Last Name' | 'House' | 'Postcode'
  original: unknown
  cleaned: string
}

export interface SheetNote {
  table: string              // "file.xlsx:Sheet"
  platform: string           // "LaunchGood" or "Unknown layout - use the column-mapping screen"
  rowsUsed: number
  copiesSkipped: number
  dateOrder: 'US (month first)' | 'UK (day first)' | ''
}

export interface HmrcFilePlan {
  name: string
  kind: 'UK' | 'Overseas'
  rows: ClaimRow[]
  total: number
  firstDate: string
  lastDate: string
}

export interface ClaimResult {
  platforms: string[]
  sheets: SheetNote[]
  linesRead: number
  ready: ClaimRow[]            // UK, sorted by date - same order as the HMRC files
  overseas: ClaimRow[]
  review: ReviewRow[]
  ineligible: IneligibleRow[]
  changes: Change[]
  files: HmrcFilePlan[]        // UK files first, then overseas
  box1Date: string | null      // earliest donation in the claim (HMRC Box 1)
  totals: {
    ukDonations: number; overseasDonations: number; reviewDonations: number
    giftAid: number            // 25% of UK + overseas
    reviewedDonations: number  // every real donation in the files, with or without Gift Aid
    cleanedRows: number
    compliancePct: number
  }
  reviewReasons: [string, number][]
  ineligibleReasons: [string, number][]
}

export class JustGivingError extends Error {
  constructor() { super('JustGiving export detected: JustGiving claims Gift Aid for the charity itself, so these donations must NOT go in a BoostAid claim.') }
}

type Rec = SourceRecord & { where: string; platform: string; claims: boolean; dateFmt: DateFormat }

export interface BuildOptions {
  /** Only if the charity confirms in writing that a "claims itself" platform does NOT claim for them. */
  platformDoesNotClaim?: boolean
}

export function buildClaim(tables: Table[], opts: BuildOptions = {}): ClaimResult {
  // ---------------------------------------------------------------- 1. read + detect platform per sheet
  const raw: Rec[] = []
  const seen = new Set<string>()
  const sheets: SheetNote[] = []
  const platforms: string[] = []
  for (const t of tables) {
    if (!t.rows.length) continue
    const title = `${t.fileName}:${t.sheetName}`
    const hdr = t.rows[0].map((h) => s(h))
    if (isJustGiving(hdr)) throw new JustGivingError()
    const prof = detectProfile(hdr)
    if (!prof) { sheets.push({ table: title, platform: 'Unknown layout - use the column-mapping screen', rowsUsed: 0, copiesSkipped: 0, dateOrder: '' }); continue }
    if (!platforms.includes(prof.name)) platforms.push(prof.name)
    let added = 0, skipped = 0
    const table: Rec[] = []
    t.rows.slice(1).forEach((r, idx) => {
      if (!r.some((v) => v !== null && v !== undefined && v !== '')) return
      const key = JSON.stringify(r)
      if (seen.has(key)) { skipped++; return }         // exact copy of a row already read (e.g. LaunchGood "Sheet1")
      seen.add(key)
      const d: Record<string, unknown> = {}
      hdr.forEach((h, i) => { d[h] = r[i] ?? null })
      table.push({ ...prof.map(d), where: `${title}!${idx + 2}`, platform: prof.name, claims: prof.claimsGiftAid, dateFmt: 'DMY4' })
      added++
    })
    const fmt = detectDateOrder(table.filter((x) => !(x.dateRaw && typeof x.dateRaw === 'object')).map((x) => x.dateRaw)) // per file
    table.forEach((x) => { x.dateFmt = fmt })
    raw.push(...table)
    sheets.push({ table: title, platform: prof.name, rowsUsed: added, copiesSkipped: skipped, dateOrder: fmt.startsWith('M') ? 'US (month first)' : 'UK (day first)' })
  }
  const refunded = new Set(raw.filter((r) => r.refundOf !== undefined && r.refundOf !== null && r.refundOf !== '').map((r) => s(r.refundOf)))

  // ---------------------------------------------------------------- 2. eligibility + cleaning
  const ready: ClaimRow[] = [], overseas: ClaimRow[] = [], review: ReviewRow[] = [], ineligible: IneligibleRow[] = [], changes: Change[] = []
  const reviewedLines: Rec[] = []
  for (const r of raw) {
    const name0 = `${r.first} ${r.last}`.trim()
    const changesBefore = changes.length
    if (!r.kindOk) reviewedLines.push(r)
    const base = { donationId: r.id, source: r.where, name: name0, amount: r.amount }
    if (r.kindOk) { ineligible.push({ ...base, reason: r.kindOk }); continue }
    if (!r.giftAid) { ineligible.push({ ...base, reason: 'No Gift Aid declaration' }); continue }
    if (s(r.id) && refunded.has(s(r.id))) { ineligible.push({ ...base, reason: 'Refunded' }); continue }
    if (r.amount <= 0) { ineligible.push({ ...base, reason: 'Amount is zero or negative' }); continue }
    if (r.claims && !opts.platformDoesNotClaim) {
      ineligible.push({ ...base, reason: `${r.platform} claims Gift Aid for the charity - claiming again would be a double claim` }); continue
    }

    const why: string[] = []
    const foreign = r.currency.toUpperCase() !== 'GBP'
    if (foreign) why.push(`Not in GBP (${fmtG(r.amount)} ${r.currency}) - check the sterling amount received`) // never auto-convert

    // --- names
    let fn0 = r.first, ln0 = r.last
    const bill = r.billing.replace(/\s+/g, ' ')
    const bw = bill.split(' ')
    let suggested = ''
    const placeholder = ['guest user', 'anonymous', 'anon'].includes(name0.toLowerCase()) || name0.includes('@') || /\d/.test(name0)
    let single = Boolean(fn0) !== Boolean(ln0) ? (fn0 || ln0) : ''
    if (single && /^[A-Za-z]+[ .]+[A-Za-z' \-]+$/.test(single)) {              // "O.Ahmed" / "Mohammed El-Jarad" in one column
      const m = single.match(/^(.*?)[ .]+(.*)$/)!
      fn0 = m[1]; ln0 = m[2]; single = ''
      changes.push({ donationId: r.id, field: 'Name', original: name0, cleaned: `${fn0} ${ln0} (split into first and last name)` })
    }
    const goodBill = bw.length >= 2 && /^[A-Za-z' .\-]+$/.test(bill) && bill.toLowerCase() !== 'guest user'
    if (single && goodBill && single.toLowerCase() === bw[0].toLowerCase()) {
      fn0 = bw[0]; ln0 = bw.slice(1).join(' ')
      changes.push({ donationId: r.id, field: 'Name', original: name0, cleaned: `${fn0} ${ln0} (surname from billing name)` })
    } else if (single && goodBill && single.toLowerCase() === bw[bw.length - 1].toLowerCase()) {
      fn0 = bw.slice(0, -1).join(' '); ln0 = bw[bw.length - 1]
      changes.push({ donationId: r.id, field: 'Name', original: name0, cleaned: `${fn0} ${ln0} (first name from billing name)` })
    } else if (goodBill && (placeholder || !fn0 || !ln0 || nonLatin(fn0 + ln0))) {
      suggested = bill                                                           // a HINT only - a person must check it
    }
    if (nonLatin(fn0 + ln0)) why.push('Name not in English letters')
    let fn = toAscii(fn0), ln = toAscii(ln0)
    if ((fn + ln).includes('@')) why.push('Email address in name')
    if (/\d/.test(fn + ln)) why.push('Digits in name')
    if (/[&+]|\band\b|,/i.test(fn + ' ' + ln)) why.push('Two people named - Gift Aid needs one taxpayer: confirm with the donor before claiming')
    const tidyName = (x: string) => x.replace(/[^A-Za-z' \-.]/g, ' ').replace(/\./g, ' ').replace(/\s+/g, ' ').replace(/^[ \-']+|[ \-']+$/g, '')
    fn = tidyName(fn); ln = tidyName(ln)
    const t0 = s(r.title).replace(/^\.+|\.+$/g, '')
    let title = TITLES.has(t0.toLowerCase()) ? t0[0].toUpperCase() + t0.slice(1).toLowerCase() : ''
    const w = fn.split(' ')
    if (w.length > 1 && TITLES.has(w[0].toLowerCase())) { title = w[0][0].toUpperCase() + w[0].slice(1).toLowerCase(); fn = w.slice(1).join(' ') }
    if (fn.includes(' ')) fn = fn.split(' ')[0]                                  // HMRC: first name with no spaces
    fn = cap(fn); ln = cap(ln)
    if (!fn || !ln) why.push('Missing first or last name')
    if (JUNK.has(fn.toLowerCase()) || JUNK.has(ln.toLowerCase())) why.push('Anonymous / placeholder name')
    if (fn.length > 35 || ln.length > 35) why.push('Name over 35 characters')

    // --- address
    const country = toAscii(s(r.country)).toLowerCase()
    const isUk = UK_WORDS.has(country)
    let house: string | null, pc: string
    if (isUk) {
      const h = houseFrom(r.lines)
      house = h.house
      if (h.reason) why.push(h.reason)
      pc = formatPostcode(r.postcode)
      if (!pc) why.push('No postcode')
      else if (!UK_POSTCODE.test(pc)) why.push(!r.country ? 'Invalid postcode - UK or overseas?' : 'Invalid or partial postcode')
    } else {                                                                      // overseas: full address in the house box, postcode X
      const lines = r.lines.filter((x) => s(x)).map((x) => toAscii(s(x)))
      house = [...lines, s(r.country)].join(', '); pc = 'X'
      if (!lines.length) why.push('No address')
      if (house.length > 40) {
        house = [...lines.slice(0, 1), s(r.country)].join(', ')
        if (house.length > 40) why.push('Overseas address over 40 characters')
      }
    }
    const date = parseDate(r.dateRaw, r.dateFmt)
    if (!date) why.push('Date not readable')

    const out: ClaimRow = {
      title, firstName: fn, lastName: ln, house: house ?? '', postcode: pc, date,
      amount: foreign ? round2(r.gbpEst ?? 0) : round2(r.amount),
      platform: r.platform, donationId: r.id, source: r.where, cleaned: false,
    }
    const origAddr = r.lines.filter((x) => s(x)).map((x) => s(x)).join(' | ')
    const fieldPairs: [Change['field'], unknown, string][] = [['First Name', r.first, out.firstName], ['Last Name', r.last, out.lastName], ['House', origAddr, out.house], ['Postcode', s(r.postcode), out.postcode]]
    for (const [field, o, c] of fieldPairs) if (s(o) !== s(c)) changes.push({ donationId: r.id, field, original: o, cleaned: c })
    out.cleaned = changes.length > changesBefore

    if (why.length) {
      review.push({
        ...out, reason: [...new Set(why)].join('; '), suggestedName: suggested,
        originalAmount: `${fmtG(r.amount)} ${r.currency}`, originalName: name0, originalAddress: origAddr,
        originalPostcode: s(r.postcode), country: s(r.country) || 'GB',
      })
    } else if (isUk) ready.push(out)
    else overseas.push(out)
  }

  // ---------------------------------------------------------------- 3. plan the HMRC files
  const files = [...plan(ready, 'UK_schedule', 'UK'), ...plan(overseas, 'Overseas_schedule', 'Overseas')]
  const claimed = [...ready, ...overseas]
  const ukTot = sum(ready), osTot = sum(overseas), rvTot = sum(review)
  const reviewedTotal = reviewedLines.reduce((a, r) => a + (r.currency.toUpperCase() === 'GBP' ? r.amount : (r.gbpEst ?? 0)), 0)
  return {
    platforms, sheets, linesRead: raw.length, ready, overseas, review, ineligible, changes, files,
    box1Date: claimed.length ? claimed.map((r) => r.date!).sort()[0] : null,
    totals: {
      ukDonations: ukTot, overseasDonations: osTot, reviewDonations: rvTot, giftAid: (ukTot + osTot) * 0.25,
      reviewedDonations: reviewedTotal, cleanedRows: claimed.filter((r) => r.cleaned).length,
      compliancePct: claimed.length + review.length ? (claimed.length / (claimed.length + review.length)) * 100 : 0,
    },
    reviewReasons: counter(review.flatMap((r) => r.reason.split('; '))),
    ineligibleReasons: counter(ineligible.map((r) => r.reason)),
  }
}

/** Sort by date and split into files of at most 1,000 rows. Each row learns its file name and row number. */
function plan(rows: ClaimRow[], prefix: string, kind: 'UK' | 'Overseas'): HmrcFilePlan[] {
  rows.sort((a, b) => cmp(a.date ?? '', b.date ?? '') || cmp(a.platform, b.platform) || cmp(s(a.donationId), s(b.donationId)))
  const n = Math.ceil(rows.length / MAX_ROWS_PER_HMRC_FILE)
  const out: HmrcFilePlan[] = []
  for (let i = 0; i < n; i++) {
    const chunk = rows.slice(i * MAX_ROWS_PER_HMRC_FILE, (i + 1) * MAX_ROWS_PER_HMRC_FILE)
    const name = `${prefix}_${i + 1}_of_${n}.ods`
    chunk.forEach((r, j) => { r.hmrcFile = name; r.hmrcRow = 25 + j })
    out.push({ name, kind, rows: chunk, total: round2(sum(chunk)), firstDate: chunk[0].date!, lastDate: chunk[chunk.length - 1].date! })
  }
  return out
}

const cmp = (a: string, b: string) => (a < b ? -1 : a > b ? 1 : 0)
const sum = (rows: { amount: number }[]) => rows.reduce((a, r) => a + r.amount, 0)
export const round2 = (x: number) => Math.round((x + Number.EPSILON) * 100) / 100

/** Like Python's Counter.most_common(): highest count first, ties in first-seen order. */
function counter(items: string[]): [string, number][] {
  const m = new Map<string, number>()
  for (const x of items) m.set(x, (m.get(x) ?? 0) + 1)
  return [...m.entries()].sort((a, b) => b[1] - a[1])
}

export { PROFILES }
