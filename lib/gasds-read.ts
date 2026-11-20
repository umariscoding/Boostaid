/**
 * Reads the files a charity uploads for GASDS (CSV or Excel):
 *   a) the receipt book's "Full receipt list"  b) its "Banking" export (has "Counted in envelope")
 *   c) the BoostAid GASDS template.
 * Columns are found by heading, in any order. Pure: no database, so it runs in the browser and in tests.
 */
import * as XLSX from 'xlsx'
import { taxYearOf, type BankingFile, type GReceipt } from './gasds'

const MAX_ROWS = 20000

function parseCsv(text: string): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let cell = ''
  let q = false
  text = text.replace(/^﻿/, '')
  const endRow = () => {
    row.push(cell)
    cell = ''
    if (row.some((x) => x.trim())) rows.push(row)
    row = []
  }
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (q) {
      if (ch === '"') {
        if (text[i + 1] === '"') { cell += '"'; i++ } else q = false
      } else cell += ch
      continue
    }
    if (ch === '"') q = true
    else if (ch === ',') { row.push(cell); cell = '' }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++
      endRow()
    } else cell += ch
  }
  endRow()
  return rows
}

/** yyyy-mm-dd or dd/mm/yyyy (also dd/mm/yy) -> yyyy-MM-dd, or '' if it isn't a date. */
function toIso(s: string): string {
  s = String(s).trim()
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s)
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s)
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{2})$/.exec(s)
  return m ? `20${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}` : ''
}

const money = (s: string | undefined) => Number(String(s ?? '').replace(/[£,\s]/g, ''))

/** The text of one uploaded file as CSV, whatever it was. */
export function fileToCsv(name: string, data: Uint8Array): string {
  if (/\.xlsx?$/i.test(name)) {
    const wb = XLSX.read(data, { type: 'array', cellDates: true })
    const sheet = wb.Sheets[wb.SheetNames[0]]
    if (!sheet) return ''
    return XLSX.utils.sheet_to_csv(sheet, { dateNF: 'dd/mm/yyyy', blankrows: false })
  }
  return new TextDecoder('utf-8').decode(data)
}

/** True for the headings of a receipt book export, a banking export or the BoostAid GASDS template (not Gift Aid donations). */
export function looksLikeGasds(headers: unknown[]): boolean {
  const h = headers.map((x) => String(x ?? '').trim().toLowerCase())
  const has = (n: string) => h.includes(n)
  return has('paying-in slip') || has('counted in envelope') || (has('banked') && has('gift aid') && (has('receipt no') || has('paid by') || has('type')))
}

export const isBankingCsv = (csv: string) => /counted in envelope/i.test(csv.split(/\r?\n/)[0] ?? '')

const COLS = {
  no: 'receipt no', date: 'date', amount: 'amount', pay: 'paid by', notes: 'notes', status: 'status',
  donor: 'donor', giftAid: 'gift aid', banked: 'banked', kind: 'type', slip: 'paying-in slip',
} as const

/** A receipt list or a template: one receipt per row. Voids are returned (flagged) so the caller can skip them. */
export function readReceipts(csv: string): GReceipt[] {
  const rows = parseCsv(csv)
  if (rows.length < 2) throw new Error('The file has no receipts in it.')
  if (rows.length > MAX_ROWS) throw new Error('The file is too big. Please upload it in parts.')
  const head = rows[0].map((h) => h.trim().toLowerCase())
  const idx = Object.fromEntries(Object.entries(COLS).map(([k, v]) => [k, head.indexOf(v)])) as Record<keyof typeof COLS, number>
  const missing = (['date', 'amount', 'giftAid', 'banked'] as const).filter((k) => idx[k] < 0).map((k) => COLS[k])
  if (missing.length) {
    throw new Error(`This file isn’t a receipt book export or the BoostAid template (it has no ${missing.join(', ')} column).`)
  }
  const seen: Record<string, number> = {}
  const out: GReceipt[] = []
  rows.slice(1).forEach((r, i) => {
    const get = (k: keyof typeof COLS) => (idx[k] < 0 ? '' : String(r[idx[k]] ?? '').trim())
    const date = toIso(get('date'))
    const amount = money(get('amount'))
    if (!date || !(amount > 0)) throw new Error(`Row ${i + 2}: please check the date and amount.`)
    const donor = /anonymous/i.test(get('donor')) ? '' : get('donor')
    let no = get('no')
    if (!no) {
      const base = `T ${date} ${amount.toFixed(2)} ${(get('notes') || get('donor')).slice(0, 30)}`
      seen[base] = (seen[base] ?? 0) + 1
      no = seen[base] > 1 ? `${base} #${seen[base]}` : base
    }
    const pay = get('pay')
    const slip = get('slip')
    out.push({
      no: no.slice(0, 120),
      slip: slip || undefined,
      date,
      year: taxYearOf(date),
      amount,
      card: /card|contactless|chip|pin/i.test(pay),
      notCash: /cheque|check|transfer|bacs|standing|direct debit|online|paypal|bank/i.test(pay),
      notes: get('notes').slice(0, 300),
      donor: donor.slice(0, 200),
      // A "Type" column decides; otherwise a named donor (or anything up to £30) is one person's gift
      gift: idx.kind >= 0 && get('kind') ? !/collection/i.test(get('kind')) : !!donor || amount <= 30,
      giftAid: /^y/i.test(get('giftAid')),
      banked: /^y/i.test(get('banked')) || !!slip,
      isVoid: /void/i.test(get('status')),
    })
  })
  return out
}

/** A banking export: the receipts in one paying-in envelope, then a "Paying-in slip, <number>" row. */
export function readBanking(csv: string): BankingFile {
  const rows = parseCsv(csv)
  const head = rows[0].map((h) => h.trim().toLowerCase())
  const iAmt = head.indexOf('receipt amount')
  const iCnt = head.indexOf('counted in envelope')
  let slip = ''
  const items: BankingFile['items'] = []
  for (const r of rows.slice(1)) {
    const first = String(r[0] ?? '').trim()
    if (/paying-in slip/i.test(first)) { slip = String(r[1] ?? '').trim(); continue }
    if (!first || /^total/i.test(first)) continue
    const amount = money(r[iAmt])
    const counted = money(r[iCnt])
    items.push({ no: first, amount, counted: Number.isNaN(counted) ? amount : counted })
  }
  if (!slip) throw new Error('No bank slip number found in the banking file.')
  return { slip, items }
}

export interface ReadResult { receipts: GReceipt[]; banking: BankingFile[]; errors: string[] }

/** Reads several files at once. A bad file is reported by name; the good ones still count. */
export function readGasdsFiles(files: { name: string; data: Uint8Array }[]): ReadResult {
  const receipts: GReceipt[] = []
  const banking: BankingFile[] = []
  const errors: string[] = []
  for (const f of files) {
    try {
      const csv = fileToCsv(f.name, f.data)
      if (!csv.trim()) throw new Error('The file is empty or could not be read.')
      if (isBankingCsv(csv)) banking.push(readBanking(csv))
      else receipts.push(...readReceipts(csv).filter((r) => !r.isVoid))
    } catch (e) {
      errors.push(`${f.name}: ${(e as Error).message}`)
    }
  }
  return { receipts, banking, errors }
}
