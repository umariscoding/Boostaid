/**
 * Gift Aid recovery, the rules (no database, no UI): reading emails out of an export, the secret link codes,
 * checking a donor's answers with the cleaning engine, and what a new upload does to rows that were already held back.
 */
import { buildClaim, type ReviewRow } from '@/lib/giftaid'
import type { Table } from '@/lib/giftaid/read'

export const REMINDER_AFTER_DAYS = 14
export const NO_RESPONSE_AFTER_DAYS = 14 // after the reminder
export const LINK_VALID_DAYS = 60

export const DECLARATION_TEXT =
  'I want to Gift Aid my donation to {charity}. I am a UK taxpayer and understand that if I pay less Income Tax and/or Capital Gains Tax than the amount of Gift Aid claimed on all my donations in that tax year it is my responsibility to pay any difference.'

// ------------------------------------------------------------------------------ emails in an export
const text = (v: unknown) => (v === null || v === undefined ? '' : String(v).trim())
const EMAIL_RE = /^[^\s@,;<>()]+@[^\s@,;<>()]+\.[^\s@,;<>()]{2,}$/

/** The column holding the donor's email: "Email", "Email Address", then any other e-mail column (never "Reward Email"). */
export function emailColumn(headers: unknown[]): number {
  const h = headers.map((x) => text(x).toLowerCase())
  const exact = ['email', 'e-mail', 'email address', 'donor email']
  for (const name of exact) { const i = h.indexOf(name); if (i >= 0) return i }
  return h.findIndex((x) => /e-?mail/.test(x) && !/reward|gift|recipient|friend/.test(x))
}

/** The donor's email for one engine row, found through its source ("file.xlsx:Sheet!57"). '' if there isn't a usable one. */
export function emailForSource(tables: Table[], source: string): string {
  const bang = source.lastIndexOf('!')
  const key = source.slice(0, bang)
  const t = tables.find((x) => `${x.fileName}:${x.sheetName}` === key)
  if (!t) return ''
  const col = emailColumn(t.rows[0] ?? [])
  if (col < 0) return ''
  const row = t.rows[Number(source.slice(bang + 1)) - 1]
  let v = text(row?.[col]).toLowerCase()
  // Some exports fall back to a second email column when the first is empty
  if (!EMAIL_RE.test(v)) {
    const others = (t.rows[0] ?? []).map((x, i) => (i !== col && /^e-?mail( address)?$/i.test(text(x)) ? i : -1)).filter((i) => i >= 0)
    for (const i of others) { const o = text(row?.[i]).toLowerCase(); if (EMAIL_RE.test(o)) { v = o; break } }
  }
  return EMAIL_RE.test(v) ? v.slice(0, 200) : ''
}

export interface HeldBackRow {
  platform: string
  donationId: string
  donationDate: string | null
  amount: number
  title: string
  firstName: string
  lastName: string
  house: string
  postcode: string
  reason: string
  sourceFile: string
  email: string
}

/** The engine's Needs-review rows, with each donor's email, ready to be stored. */
export function heldBackRows(review: ReviewRow[], tables: Table[]): HeldBackRow[] {
  return review.map((r) => ({
    platform: r.platform,
    // A file with no donation ID still gets a stable key from where the row came from
    donationId: r.donationId === null || r.donationId === undefined || String(r.donationId) === '' ? r.source : String(r.donationId),
    donationDate: r.date,
    amount: r.amount,
    title: r.title,
    firstName: r.firstName,
    lastName: r.lastName,
    house: r.house,
    postcode: r.postcode,
    reason: r.reason,
    sourceFile: r.source.split(':')[0],
    email: emailForSource(tables, r.source),
  }))
}

/** Emailable donors: one email per donor address, however many donations they have held back. */
export function groupByDonor<T extends { email: string | null }>(rows: T[]): Map<string, T[]> {
  const out = new Map<string, T[]>()
  for (const r of rows) {
    const e = (r.email ?? '').trim().toLowerCase()
    if (!e) continue
    out.set(e, [...(out.get(e) ?? []), r])
  }
  return out
}

export const addDays = (d: Date, n: number) => new Date(d.getTime() + n * 86400000)

// ------------------------------------------------------------------------------ the donor's answers
export interface Answers { title: string; first: string; last: string; house: string; postcode: string }
export type Checked = { ok: true; cleaned: Answers; overseas: boolean } | { ok: false; error: string }

const CASH_BOOK = ['Title', 'First name', 'Last name', 'House name or number', 'Postcode', 'Donation date', 'Amount']

/**
 * Runs the donor's answers through the cleaning engine, exactly like an uploaded row. Only a row the engine would put
 * in the claim passes; otherwise the donor is told what to fix (in their words, not the review reason).
 */
export function checkAnswers(a: Answers, date: string, amount: number): Checked {
  const first = a.first.trim(), last = a.last.trim(), house = a.house.trim(), postcode = a.postcode.trim()
  if (!first || !last) return { ok: false, error: 'Enter your first and last name.' }
  if (/&|\band\b|\+|\//i.test(first + ' ' + last)) return { ok: false, error: 'Gift Aid is for one person. Enter the name of the person who pays the tax.' }
  if (!house) return { ok: false, error: 'Enter your house name or number.' }
  if (!postcode) return { ok: false, error: 'Enter your postcode.' }
  const table: Table = {
    fileName: 'donor-answer.csv',
    sheetName: 'CSV',
    rows: [CASH_BOOK, [a.title.trim() || null, first, last, house, postcode, date, amount]],
  }
  const res = buildClaim([table])
  const row = res.ready[0] ?? res.overseas[0]
  if (!row) {
    const why = res.review[0]?.reason ?? ''
    if (/postcode/i.test(why)) return { ok: false, error: 'Enter a full UK postcode, for example LS8 2HP.' }
    if (/house/i.test(why)) return { ok: false, error: 'Enter your house name or number, for example 12 or Rose Cottage.' }
    if (/name|anonymous|script|email/i.test(why)) return { ok: false, error: 'Please check your name. Enter your own first and last name in English letters.' }
    return { ok: false, error: 'Please check your details and try again.' }
  }
  return { ok: true, cleaned: { title: row.title, first: row.firstName, last: row.lastName, house: row.house, postcode: row.postcode }, overseas: res.overseas.length > 0 }
}

// ------------------------------------------------------------------------------ a new upload meets held-back rows
export type Held = { key: string; status: string; id?: string }

/**
 * One donation is only ever in one place (waiting for a reply, recovered, or in one claim). What a new upload does:
 *  - still missing and already stored            -> nothing changes (stays "Email sent" etc.): no new row, no second email
 *  - still missing and new                       -> added to Needs review
 *  - now clean in the file (was held back)       -> fixed from the file: it joins this claim, no fee, the email link stops working
 */
export function reconcileUpload(existing: Held[], reviewKeys: string[], readyKeys: string[]): { add: string[]; fixedFromFile: string[] } {
  const have = new Map(existing.map((e) => [e.key, e]))
  const add = reviewKeys.filter((k) => !have.has(k))
  const fixedFromFile = readyKeys.filter((k) => {
    const e = have.get(k)
    return !!e && e.status !== 'claimed'
  })
  return { add, fixedFromFile }
}

export const rowKey = (platform: string, donationId: string) => `${platform}\u0000${donationId}`

/** Status text as shown to the charity. */
export const STATUS_LABEL: Record<string, string> = {
  none: 'Needs review',
  sent: 'Email sent',
  reminded: 'Reminder sent',
  no_response: 'No response',
  recovered: 'Ready – recovered',
  claimed: 'In a claim',
}
