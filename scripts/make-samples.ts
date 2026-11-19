/**
 * Builds FAKE donation exports for the 7 platforms BoostAid reads, plus an answer sheet per file.
 * Every donor, address, ID and amount is invented. Column headings and value formats copy the real
 * exports so the files are detected exactly like the real thing.
 *
 *   npm run samples:make   (writes samples/)
 */
import fs from 'node:fs'
import path from 'node:path'
import * as XLSX from 'xlsx'
import HEADERS from './sample-headers.json'

type Bucket = 'ready' | 'overseas' | 'review' | 'ineligible' | 'skipped'

interface Case {
  key: string
  what: string
  first: string
  last: string
  title?: string
  a1: string
  a2?: string
  town?: string
  pc: string
  /** '' = UK by default, 'GB' = UK stated, 'FR' = France */
  country?: '' | 'GB' | 'FR'
  date: string // yyyy-mm-dd, or '' for a missing date
  amount: number
  currency?: string
  gbp?: number // platform's own GBP conversion for non-GBP rows
  giftAid?: boolean
  /** Platform-specific raw column overrides */
  raw?: Record<string, unknown>
  expect: Bucket
  reason?: string
  /** For ready rows: what the HMRC schedule should show */
  out?: { title?: string; first: string; last: string; house: string; postcode: string }
  /** Emit this row twice, identical (the second copy should be skipped) */
  twice?: boolean
}

// ------------------------------------------------------------------ the shared test cases
// Feature flags decide which cases a platform can express (e.g. no currency column = no USD case).
interface Features { country: boolean; currency: boolean; cashBook?: boolean }

function commonCases(f: Features, d: (n: number) => string): Case[] {
  const c: Case[] = [
    { key: 'clean', what: 'Clean UK donation', first: 'Amina', last: 'Khan', title: 'Ms', a1: '14 Cedar Road', town: 'London', pc: 'E7 9AB', country: f.country ? 'GB' : '', date: d(0), amount: 25, expect: 'ready', out: { title: 'Ms', first: 'Amina', last: 'Khan', house: '14', postcode: 'E7 9AB' } },
    { key: 'tidy', what: 'Messy capitals, spaces and lower-case postcode get tidied', first: '  omar ', last: 'FAROOQ', a1: '22 elm street', town: 'Birmingham', pc: 'b11 2qt', date: d(1), amount: 10, expect: 'ready', out: { first: 'Omar', last: 'Farooq', house: '22', postcode: 'B11 2QT' } },
    { key: 'title-in-name', what: 'Title typed into the first-name box', first: 'Dr Sana', last: 'Iqbal', a1: '7 Park Avenue', town: 'Leeds', pc: 'LS8 2HP', date: d(2), amount: 50, expect: 'ready', out: { title: 'Dr', first: 'Sana', last: 'Iqbal', house: '7', postcode: 'LS8 2HP' } },
    { key: 'flat', what: 'Flat number and building name', first: 'Zainab', last: 'Choudhury', a1: 'Flat 3 Oak House', a2: '18 Mill Lane', town: 'Bradford', pc: 'BD7 1HD', date: d(3), amount: 15, expect: 'ready', out: { first: 'Zainab', last: 'Choudhury', house: 'Flat 3 Oak House', postcode: 'BD7 1HD' } },
    { key: 'repeat-donor', what: 'Same donor gives again on a different day (both claimed)', first: 'Amina', last: 'Khan', title: 'Ms', a1: '14 Cedar Road', town: 'London', pc: 'E7 9AB', country: f.country ? 'GB' : '', date: d(4), amount: 30, expect: 'ready', out: { title: 'Ms', first: 'Amina', last: 'Khan', house: '14', postcode: 'E7 9AB' } },
    { key: 'no-declaration', what: 'Donor did not tick Gift Aid', first: 'Daniel', last: 'Hughes', a1: '9 Station Road', town: 'Cardiff', pc: 'CF10 1EP', date: d(5), amount: 40, giftAid: false, expect: 'ineligible', reason: 'No Gift Aid declaration' },
    { key: 'zero', what: 'Zero amount', first: 'Rehana', last: 'Malik', a1: '31 Kings Road', town: 'Bolton', pc: 'BL1 4QX', date: d(6), amount: 0, expect: 'ineligible', reason: 'Amount is zero or negative' },
    { key: 'joint', what: 'Two people in the name - one taxpayer needed', first: 'Bilal & Maryam', last: 'Hussain', a1: '5 Grove Street', town: 'Leicester', pc: 'LE4 6JT', date: d(7), amount: 100, expect: 'review', reason: 'Two people named' },
    { key: 'bad-postcode', what: 'Half a postcode', first: 'Imran', last: 'Sheikh', a1: '40 High Street', town: 'London', pc: 'E1 6A', country: f.country ? 'GB' : '', date: d(8), amount: 20, expect: 'review', reason: 'postcode' },
    { key: 'no-postcode', what: 'Postcode missing', first: 'Ruth', last: 'Evans', a1: '12 Church Lane', town: 'Swansea', pc: '', date: d(9), amount: 35, expect: 'review', reason: 'No postcode' },
    { key: 'anonymous', what: 'Anonymous placeholder name', first: 'Anonymous', last: 'Donor', a1: '1 Market Square', town: 'Luton', pc: 'LU1 2LJ', date: d(10), amount: 5, expect: 'review', reason: 'Anonymous / placeholder name' },
    { key: 'no-surname', what: 'Surname missing', first: 'Yasmin', last: '', a1: '66 Queens Road', town: 'Coventry', pc: 'CV2 3LR', date: d(11), amount: 12, expect: 'review', reason: 'Missing first or last name' },
    { key: 'exact-copy', what: 'The same line appears twice in the export (second copy skipped)', first: 'Hamza', last: 'Rashid', a1: '27 Victoria Street', town: 'Manchester', pc: 'M14 5RN', date: d(12), amount: 60, expect: 'ready', twice: true, out: { first: 'Hamza', last: 'Rashid', house: '27', postcode: 'M14 5RN' } },
  ]
  if (f.currency) {
    c.push({ key: 'usd', what: 'Paid in US dollars - never auto-converted', first: 'Sarah', last: 'Ahmed', a1: '3 Albert Road', town: 'Oldham', pc: 'OL8 1BS', date: d(13), amount: 50, currency: 'USD', gbp: 39.4, expect: 'review', reason: 'Not in GBP' })
  }
  if (f.country) {
    c.push({ key: 'overseas', what: 'Overseas donor - goes on the overseas schedule, postcode X', first: 'Karim', last: 'Benali', a1: '5 Rue Cler', town: 'Paris', pc: '75007', country: 'FR', date: d(14), amount: 20, expect: 'overseas' })
  }
  return c
}

// ------------------------------------------------------------------ helpers
const pad = (n: number, w = 2) => String(n).padStart(w, '0')
const parts = (iso: string) => iso.split('-').map(Number) as [number, number, number]
const us4 = (iso: string) => { const [y, m, d] = parts(iso); return `${pad(m)}/${pad(d)}/${y}` }
const uk4 = (iso: string) => { const [y, m, d] = parts(iso); return `${pad(d)}/${pad(m)}/${y}` }
const uk2 = (iso: string) => { const [y, m, d] = parts(iso); return `${pad(d)}/${pad(m)}/${String(y).slice(2)}` }
const usTime = (iso: string, i: number) => `${us4(iso)} ${1 + (i % 11)}:${pad((i * 7) % 60)}:${pad((i * 13) % 60)} ${i % 2 ? 'PM' : 'AM'}`
const ukTime = (iso: string, i: number) => `${uk4(iso)} ${pad(9 + (i % 10))}:${pad((i * 7) % 60)}`
const excelDate = (iso: string) => { const [y, m, d] = parts(iso); return new Date(Date.UTC(y, m - 1, d)) }
/** Dates spread over Jan-Oct 2026; several have a day over 12 so US/UK order is detectable. */
const dayOf = (start: string) => (n: number) => {
  const t = new Date(Date.UTC(...(([y, m, d]) => [y, m - 1, d] as const)(parts(start))) + n * 9 * 86400000)
  return `${t.getUTCFullYear()}-${pad(t.getUTCMonth() + 1)}-${pad(t.getUTCDate())}`
}
const countryName = (c?: string) => (c === 'FR' ? 'France' : c === 'GB' ? 'United Kingdom' : '')

interface PlatformDef {
  file: string
  sheet: string
  headers: string[]
  format: 'xlsx' | 'csv'
  cases: Case[]
  row: (c: Case, i: number) => Record<string, unknown>
  /** The export has no title column, so the schedule's title is always blank */
  noTitle?: boolean
  /** Extra sheets (LaunchGood exports a second sheet that repeats rows) */
  extraSheet?: { name: string; copyRows: number }
}

const H = (f: string) => (HEADERS as Record<string, { sheets: { headers: string[] }[] }>)[f].sheets[0].headers

// ------------------------------------------------------------------ the 7 platforms
const platforms: PlatformDef[] = []

{ // LaunchGood: US text dates, refunds point at the donation they refund, payouts/reserves in the same sheet
  const d = dayOf('2026-01-14')
  const cases = commonCases({ country: true, currency: true }, d)
  cases.push(
    { key: 'refunded', what: 'Donation later refunded in the same export', first: 'Faisal', last: 'Karim', a1: '8 Brook Street', town: 'Preston', pc: 'PR1 2XB', country: 'GB', date: d(15), amount: 75, raw: { 'Donation ID': '90000099' }, expect: 'ineligible', reason: 'Refunded' },
    { key: 'refund-line', what: 'The refund line itself', first: 'Faisal', last: 'Karim', a1: '8 Brook Street', town: 'Preston', pc: 'PR1 2XB', country: 'GB', date: d(16), amount: -75, raw: { 'Donation ID': '90000099', Type: 'refund' }, expect: 'ineligible', reason: 'Not a donation (refund)' },
    { key: 'payout', what: 'Payout line - not a donation', first: '', last: '', a1: '', pc: '', date: d(17), amount: -500, raw: { Type: 'payout', 'Donation ID': '' }, expect: 'ineligible', reason: 'Not a donation (payout)' },
    { key: 'surname-from-billing', what: 'Surname blank but the billing name has it', first: 'Layla', last: '', a1: '19 Hill Street', town: 'Dewsbury', pc: 'WF13 2AA', country: 'GB', date: d(18), amount: 22, raw: { 'Billing Name': 'Layla Mirza' }, expect: 'ready', out: { first: 'Layla', last: 'Mirza', house: '19', postcode: 'WF13 2AA' } },
  )
  platforms.push({
    file: 'LaunchGood_SAMPLE.xlsx', sheet: 'Transactions', headers: H('LG_Test_data.xlsx'), format: 'xlsx', cases, noTitle: true,
    extraSheet: { name: 'Sheet1', copyRows: 4 },
    row: (c, i) => ({
      'Donation ID': String(90000001 + i), 'First Name': c.first, 'Last Name': c.last, 'Display Name': `${c.first} ${c.last}`.trim(),
      Email: '', 'Anonymous?': 'no', Type: 'donation', 'Donation Currency': c.currency ?? 'GBP', 'Donation Amount': c.amount,
      'Amount (May be approx.) (PC)': c.gbp ?? c.amount, 'Gift Aid': c.giftAid === false ? 'no' : 'yes',
      'Created Date': us4(c.date), 'Created Time': '10:15:00', 'Billing Name': '', 'Billing Address': c.a1, 'Billing Address 2': c.a2 ?? '',
      'Billing City': c.town ?? '', 'Billing Zip': c.pc, 'Billing Country': c.country ?? '', 'Project Name': 'Sample Appeal',
    }),
  })
}

{ // Raisely: Excel date cells, lower-case titles, "Total" includes fees the donor covered
  const d = dayOf('2026-02-03')
  const cases = commonCases({ country: false, currency: true }, d)
  cases.push(
    { key: 'test-mode', what: 'Test payment made while setting up the page', first: 'Peter', last: 'Lewis', a1: '2 Rose Way', town: 'Reading', pc: 'RG1 3EH', date: d(15), amount: 10, raw: { Mode: 'TEST' }, expect: 'ineligible', reason: 'Test donation (TEST)' },
    { key: 'fees-covered', what: 'Donor covered the fees: claim Donation Total (20), not Total (21.10)', first: 'Noor', last: 'Begum', title: 'mrs', a1: '45 Oxford Road', town: 'Reading', pc: 'RG1 7LH', date: d(16), amount: 20, raw: { Total: 21.1 }, expect: 'ready', out: { title: 'Mrs', first: 'Noor', last: 'Begum', house: '45', postcode: 'RG1 7LH' } },
  )
  platforms.push({
    file: 'Raisely_SAMPLE.xlsx', sheet: 'donations-1000001-2026-11-16T10', headers: H('Raisely.xlsx'), format: 'xlsx', cases,
    row: (c, i) => ({
      'Campaign Name': 'Sample Appeal', Uuid: `00000000-0000-4000-8000-${pad(i + 1, 12)}`, Id: 7000001 + i, Title: (c.title ?? '').toLowerCase(),
      'First Name': c.first, 'Last Name': c.last, Email: '', Anonymous: 'false', Date: excelDate(c.date), Mode: 'LIVE', Status: 'OK',
      Type: 'ONLINE', Currency: c.currency ?? 'GBP', 'Donation Total': c.amount, Total: c.amount, Net: c.amount,
      'Includes Gift Aid': c.giftAid === false ? 'false' : 'true', 'Address 1': c.a1, 'Address 2': c.a2 ?? '', Postcode: c.pc,
      'Gift Aid Text': c.giftAid === false ? '' : 'I am a UK taxpayer and understand that if I pay less Income Tax and/or Capital Gains Tax than the amount of Gift Aid claimed on all my donations in that tax year it is my responsibility to pay any difference.',
      'Receipt Id': `R-${1000001 + i}`,
    }),
  })
}

{ // CAF Donate: Excel date cells, "Tax Effective" flag, DonationGrossAmount already includes the 25%
  const d = dayOf('2026-01-20')
  const cases = commonCases({ country: true, currency: false }, d)
  cases.push(
    { key: 'non-tax-effective', what: 'Marked Non Tax Effective', first: 'George', last: 'Bennett', a1: '17 Elm Grove', town: 'Bath', pc: 'BA1 2QS', country: 'GB', date: d(15), amount: 30, raw: { DonationTaxEffective: 'Non Tax Effective' }, expect: 'ineligible', reason: 'No Gift Aid declaration' },
    { key: 'cancelled', what: 'Payment cancelled', first: 'Hina', last: 'Aslam', a1: '4 Mill Road', town: 'Bath', pc: 'BA2 3RT', country: 'GB', date: d(16), amount: 25, raw: { DonationStatus: 'Cancelled' }, expect: 'ineligible', reason: 'Status Cancelled' },
  )
  platforms.push({
    file: 'CAF_SAMPLE.xlsx', sheet: 'Report_CSV_File', headers: H('CAF_1.xlsx'), format: 'xlsx', cases,
    row: (c, i) => ({
      DonorTitle: c.title ?? '', DonorFirstName: c.first, DonorInitials: c.first.trim().slice(0, 1), DonorSurname: c.last,
      DonorAddressLine1: c.a1, DonorAddressLine2: c.a2 ?? '', DonorAddressCity: c.town ?? '', DonorAddressPostCode: c.pc,
      DonorAddressCountry: c.country === 'FR' ? 'France' : 'United Kingdom', DonationPayerName: `${c.first} ${c.last}`.trim(),
      DonationReference: `${10000000001 + i}AB`, DonationDate: excelDate(c.date), DonationAmount: c.amount,
      DonationGrossAmount: c.giftAid === false ? c.amount : Math.round(c.amount * 125) / 100, DonationStatus: 'Paid',
      DonationTaxEffective: c.giftAid === false ? 'Non Tax Effective' : 'Tax Effective',
      GiftAidAmount: c.giftAid === false ? 0 : Math.round(c.amount * 25) / 100,
    }),
  })
}

{ // MuslimGiving: BOTH "Is Donation GA Eligible" and "UK Tax Payer" must be Yes
  const d = dayOf('2026-03-02')
  const cases = commonCases({ country: false, currency: false }, d)
  cases.push({ key: 'not-taxpayer', what: 'Gift Aid ticked but donor says they are not a UK taxpayer', first: 'Adil', last: 'Shah', a1: '10 Green Lane', town: 'Derby', pc: 'DE23 6RL', date: d(15), amount: 45, raw: { 'UK Tax Payer': 'No' }, expect: 'ineligible', reason: 'No Gift Aid declaration' })
  platforms.push({
    file: 'MuslimGiving_SAMPLE.xlsx', sheet: 'Report Mon Nov 16 2026 10_00_00', headers: H('MuslimGiving_1.xlsx'), format: 'xlsx', cases,
    row: (c, i) => ({
      'Date of Donation': excelDate(c.date), 'Donation Id': 500001 + i, 'Donor Title': c.title ?? '', 'Donor FirstName': c.first,
      'Donor LastName': c.last, 'Donor Address Line 1': c.a1, 'Donor Address Line 2': c.a2 ?? '', 'Donor Town': c.town ?? '',
      'Donor Postcode': c.pc, 'Appeal Name': 'Sample Appeal', 'Base Donation Amount': c.amount,
      'Is Donation GA Eligible': c.giftAid === false ? 'No' : 'Yes', 'UK Tax Payer': c.giftAid === false ? 'No' : 'Yes',
      'Gift Aid Amount': c.giftAid === false ? 0 : Math.round(c.amount * 25) / 100,
    }),
  })
}

{ // Fundraise Up: US date-times in mixed styles, failed/refunded payments, company donations
  const d = dayOf('2026-02-10')
  const cases = commonCases({ country: true, currency: true }, d)
  cases.push(
    { key: 'failed', what: 'Card payment failed', first: 'Tom', last: 'Walker', a1: '6 Bridge Street', town: 'York', pc: 'YO1 6DD', country: 'GB', date: d(15), amount: 20, raw: { 'Donation Status': 'Failed' }, expect: 'ineligible', reason: 'Status Failed' },
    { key: 'refunded', what: 'Refunded payment', first: 'Mariam', last: 'Yusuf', a1: '12 Castle Road', town: 'York', pc: 'YO10 3BU', country: 'GB', date: d(16), amount: 50, raw: { 'Donation Status': 'Refunded' }, expect: 'ineligible', reason: 'Status Refunded' },
    { key: 'company', what: 'Given on behalf of a company - Gift Aid is for individuals only', first: 'James', last: 'Clarke', a1: '1 Trade Park', town: 'Hull', pc: 'HU2 8DH', country: 'GB', date: d(17), amount: 250, raw: { 'On Behalf Of': 'Sample Trading Ltd' }, expect: 'ineligible', reason: 'on behalf of an organisation' },
  )
  platforms.push({
    file: 'FundraiseUp_SAMPLE.csv', sheet: 'CSV', headers: H('donations-11-11-2026-2-58-PM.csv'), format: 'csv', cases, noTitle: true,
    row: (c, i) => ({
      'Donation ID': 'SMPL' + String.fromCharCode(65 + (i % 26)) + String.fromCharCode(65 + ((i * 7) % 26)) + 'QX',
      'Donation Status': 'Succeeded', 'Donation Date': i % 2 ? usTime(c.date, i) : `${us4(c.date)} ${pad(8 + (i % 12))}:${pad((i * 5) % 60)}`,
      'Donation Frequency': 'One-time', 'Supporter First Name': c.first, 'Supporter Last Name': c.last, 'On Behalf Of': '',
      'Mailing Address Line 1': c.a1, 'Mailing Address Line 2': c.a2 ?? '', 'Mailing City': c.town ?? '', 'Mailing Zip/Postal': c.pc,
      'Mailing Country Code': c.country ?? '', 'Mailing Country': countryName(c.country), 'Donation Amount': c.amount,
      'Donation Currency': c.currency ?? 'GBP', 'Donation Amount Before Fees Covered Option': c.amount,
      'Converted Donation Amount Before Fees Covered Option': c.gbp ?? c.amount, 'Gift Aid': c.giftAid === false ? 'No' : 'Yes',
    }),
  })
}

{ // Donations Manager: UK date-times, no yes/no column - a Gift Aid amount above 0 means a declaration
  const d = dayOf('2026-01-08')
  const cases = commonCases({ country: true, currency: true }, d).map((c) =>
    // No yes/no column: a £0 donation has a £0 Gift Aid amount, so it reads as "no declaration" (still not claimed)
    c.key === 'zero' ? { ...c, what: c.what + ' (reads as no declaration: Gift Aid amount is also 0)', reason: 'No Gift Aid declaration' } : c)
  cases.push({ key: 'organisation', what: 'Donor Organisation filled in - not an individual', first: 'Helen', last: 'Price', a1: '50 Commerce Way', town: 'Slough', pc: 'SL1 4TE', country: 'GB', date: d(15), amount: 150, raw: { 'Donor Organisation': 'Sample Builders Ltd' }, expect: 'ineligible', reason: 'on behalf of an organisation' })
  platforms.push({
    file: 'DonationsManager_SAMPLE.csv', sheet: 'CSV', headers: H('donations-11112026145923-i86f5ad525e0vsr5h4c9.csv'), format: 'csv', cases,
    row: (c, i) => ({
      'Donation Ref': String(8000001 + i), 'Charity Donation Ref': String(40001 + i), 'Donation Date': ukTime(c.date, i),
      'Donation Currency': c.currency ?? 'GBP', 'Donation Exchange Rate': c.currency && c.currency !== 'GBP' ? '0.788' : '1',
      'Donation Amount Original': c.amount, 'Donation Amount (?)': c.gbp ?? c.amount,
      'Gift Aid Amount (?)': c.giftAid === false ? 0 : Math.round((c.gbp ?? c.amount) * 25) / 100,
      'Donation Source': 'Online', 'Appeal Name': 'Sample Appeal', 'Donor Title': c.title ?? '', 'Donor Forename': c.first,
      'Donor Surname': c.last, 'Donor Address Line 1': c.a1, 'Donor Address Line 2': c.a2 ?? '', 'Donor Town': c.town ?? '',
      'Donor Postcode': c.pc, 'Donor Country': c.country === 'FR' ? 'France' : 'United Kingdom', 'Donor Organisation': '',
    }),
  })
}

{ // Cash book in HMRC's own layout: every line is a declared Gift Aid donation, UK dd/mm/yy dates
  const d = dayOf('2026-04-06')
  const cases = commonCases({ country: false, currency: false, cashBook: true }, d)
    .filter((c) => !['no-declaration'].includes(c.key)) // a cash book only lists Gift Aid donations
    .map((c): Case => ({ ...c, a1: c.out?.house ?? c.a1.replace(/^(\d+)\s.*$/, '$1'), a2: '' })) // one "house name or number" box
  cases.push(
    { key: 'no-date', what: 'Date left blank', first: 'Kamal', last: 'Uddin', a1: '17', pc: 'BD7 1HD', date: '', amount: 40, expect: 'review', reason: 'Date not readable' },
    { key: 'same-day-twice', what: 'Same donor, same day, two DIFFERENT amounts (both claimed)', first: 'Tariq', last: 'Mahmood', title: 'Mr', a1: '45', pc: 'LS8 3AB', date: d(13), amount: 20, expect: 'ready', out: { title: 'Mr', first: 'Tariq', last: 'Mahmood', house: '45', postcode: 'LS8 3AB' } },
    { key: 'same-day-twice-2', what: '...second donation that day', first: 'Tariq', last: 'Mahmood', title: 'Mr', a1: '45', pc: 'LS8 3AB', date: d(13), amount: 15, expect: 'ready', out: { title: 'Mr', first: 'Tariq', last: 'Mahmood', house: '45', postcode: 'LS8 3AB' } },
  )
  platforms.push({
    file: 'CashBook_SAMPLE.csv', sheet: 'CSV', headers: H('gift-aid-schedule-SAMPLE.csv'), format: 'csv', cases,
    row: (c) => ({
      Title: c.title ?? '', 'First name': c.first, 'Last name': c.last, 'House name or number': c.a1, Postcode: c.pc,
      'Aggregated donations': '', 'Sponsored event': '', 'Donation date': c.date ? uk2(c.date) : '', Amount: c.amount.toFixed(2),
    }),
  })
}

// ------------------------------------------------------------------ write files + answer sheets
const outDir = process.argv[2]
if (!outDir) { console.error('usage: npm run samples:make   (writes samples/)'); process.exit(2) }
fs.mkdirSync(outDir, { recursive: true })

const csvCell = (v: unknown) => { const t = v === null || v === undefined ? '' : String(v); return /[",\n]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t }
const answers: Record<string, unknown>[] = []

for (const p of platforms) {
  const aoa: unknown[][] = [p.headers]
  for (const [i, c] of p.cases.entries()) {
    const rec = { ...p.row(c, i), ...(c.raw ?? {}) }
    for (const k of Object.keys(rec)) if (!p.headers.includes(k)) throw new Error(`${p.file}: unknown column ${k}`)
    const line = p.headers.map((h) => rec[h] ?? '')
    const copies = c.twice ? 2 : 1
    for (let n = 0; n < copies; n++) {
      aoa.push(line)
      const isCopy = n === 1
      answers.push({
        File: p.file, Row: aoa.length, Case: isCopy ? `${c.key} (copy)` : c.key, 'What it tests': isCopy ? 'Exact copy of the row above' : c.what,
        Expected: isCopy ? 'skipped' : c.expect, 'Reason contains': isCopy ? 'exact copy - not read twice' : c.reason ?? '',
        Title: isCopy || (p.noTitle && c.key !== 'title-in-name') ? '' : c.out?.title ?? '', 'First name': isCopy ? '' : c.out?.first ?? '', 'Last name': isCopy ? '' : c.out?.last ?? '',
        House: isCopy ? '' : c.out?.house ?? '', Postcode: isCopy ? '' : c.out?.postcode ?? '',
        'Donation date': isCopy || c.expect !== 'ready' ? '' : c.date, Amount: isCopy || c.expect !== 'ready' ? '' : c.amount,
      })
    }
  }
  const file = path.join(outDir, p.file)
  if (p.format === 'csv') {
    fs.writeFileSync(file, aoa.map((r) => r.map(csvCell).join(',')).join('\r\n') + '\r\n')
  } else {
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa, { cellDates: true, dateNF: 'dd/mm/yyyy' }), p.sheet)
    if (p.extraSheet) {
      // LaunchGood's export repeats some rows on a second sheet; BoostAid must not read them twice
      const copy = [aoa[0], ...aoa.slice(1, 1 + p.extraSheet.copyRows)]
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(copy, { cellDates: true }), p.extraSheet.name)
      for (let r = 2; r <= copy.length; r++) {
        answers.push({ File: p.file, Row: `${p.extraSheet.name}!${r}`, Case: 'second-sheet copy', 'What it tests': `Repeat of Transactions row ${r} on the "${p.extraSheet.name}" sheet`, Expected: 'skipped', 'Reason contains': 'exact copy - not read twice' })
      }
    }
    XLSX.writeFile(wb, file, { cellDates: true })
  }
  console.log(`${p.file}: ${aoa.length - 1} rows`)
}

// The answer sheet: one row per line in every sample file
const wb = XLSX.utils.book_new()
const ws = XLSX.utils.json_to_sheet(answers)
ws['!cols'] = [26, 10, 22, 60, 11, 34, 6, 12, 12, 20, 10, 12, 8].map((wch) => ({ wch }))
XLSX.utils.book_append_sheet(wb, ws, 'Expected results')
XLSX.writeFile(wb, path.join(outDir, 'EXPECTED_RESULTS.xlsx'))
fs.writeFileSync(path.join(outDir, 'expected.json'), JSON.stringify(answers, null, 1))
console.log(`EXPECTED_RESULTS.xlsx: ${answers.length} answers`)
