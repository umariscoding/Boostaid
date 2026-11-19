/**
 * Platform profiles.
 *
 * A profile ONLY maps one platform's columns onto a common record. Every HMRC rule is
 * shared (engine.ts). To support a new platform, add one profile here and a test file -
 * nothing else changes. The platform is detected from the column headings.
 */
import { s, yes, num } from './text'

export type Row = Record<string, unknown>

/** The common record every profile produces. */
export interface SourceRecord {
  id: unknown                 // the platform's donation ID (for tracing)
  kindOk: string | null       // null = a real completed donation; otherwise the reason it isn't
  giftAid: boolean            // did the donor make a Gift Aid declaration?
  amount: number              // gross donation (before fees, excluding any Gift Aid estimate)
  currency: string
  gbpEst?: number             // platform's own GBP conversion - for review totals only, never claimed
  first: string
  last: string
  billing: string             // billing / payer name - used to complete a missing surname
  title: string
  lines: unknown[]            // address lines in order
  postcode: unknown
  country: string
  dateRaw: unknown
  refundOf?: unknown          // LaunchGood: a refund row points at the donation it refunds
}

export interface Profile {
  name: string
  /** Every one of these headings must be present for the file to match this platform. */
  headings: string[]
  map: (d: Row) => SourceRecord
  /** True if the platform claims Gift Aid itself (e.g. JustGiving) - its donations must never be claimed again. */
  claimsGiftAid: boolean
}

/** First column whose heading starts with `start` (handles "Donation Amount (?)" / "(£)"). */
function col(d: Row, start: string): unknown {
  for (const k of Object.keys(d)) if (k.startsWith(start)) return d[k]
  return null
}

export const PROFILES: Profile[] = [
  {
    name: 'LaunchGood',
    headings: ['Donation ID', 'Gift Aid', 'Billing Zip'],
    claimsGiftAid: false,
    map: (d) => {
      const type = s(d['Type']).toLowerCase()
      return {
        id: d['Donation ID'],
        kindOk: type === 'donation' ? null : `Not a donation (${s(d['Type']) || 'blank'})`, // reserve, payout, refund, adjustment
        giftAid: yes(d['Gift Aid']),
        amount: num(d['Donation Amount']), currency: s(d['Donation Currency']) || 'GBP', gbpEst: num(d['Amount (May be approx.) (PC)']),
        first: s(d['First Name']), last: s(d['Last Name']), billing: s(d['Billing Name']), title: '',
        lines: [d['Billing Address'], d['Billing Address 2']], postcode: d['Billing Zip'],
        country: s(d['Billing Country']), dateRaw: d['Created Date'],
        refundOf: type === 'refund' ? d['Donation ID'] : undefined,
      }
    },
  },
  {
    name: 'Raisely',
    headings: ['Includes Gift Aid', 'Donation Total', 'Receipt Id'],
    claimsGiftAid: false,
    map: (d) => {
      let bad: string | null = null
      const mode = s(d['Mode']).toUpperCase(), status = s(d['Status']).toUpperCase()
      if (mode !== 'LIVE' && mode !== '') bad = `Test donation (${s(d['Mode'])})`
      else if (status !== 'OK' && status !== 'PAID' && status !== '') bad = `Status ${s(d['Status'])}`
      return {
        id: d['Receipt Id'] || d['Id'], kindOk: bad, giftAid: yes(d['Includes Gift Aid']),
        amount: num(d['Donation Total']), // NOT "Total", which includes fees the donor chose to cover
        currency: s(d['Currency']) || 'GBP',
        first: s(d['First Name']), last: s(d['Last Name']), billing: '', title: s(d['Title']),
        lines: [d['Address 1'], d['Address 2']], postcode: d['Postcode'], country: '', dateRaw: d['Date'],
      }
    },
  },
  {
    name: 'CAF Donate',
    headings: ['DonationTaxEffective', 'DonorSurname', 'GiftAidAmount'],
    claimsGiftAid: false, // GiftAidAmount is only CAF's 25% estimate - the charity claims
    map: (d) => {
      const st = s(d['DonationStatus']).toLowerCase()
      return {
        id: d['DonationReference'], kindOk: st === 'paid' || st === '' ? null : `Status ${s(d['DonationStatus'])}`,
        giftAid: s(d['DonationTaxEffective']).toLowerCase() === 'tax effective',
        amount: num(d['DonationAmount']), // NOT DonationGrossAmount, which already includes the 25%
        currency: 'GBP',
        first: s(d['DonorFirstName']) || s(d['DonorInitials']), last: s(d['DonorSurname']), billing: s(d['DonationPayerName']),
        title: s(d['DonorTitle']),
        lines: [d['DonorAddressLine1'], d['DonorAddressLine2'], d['DonorAddressLine3'], d['DonorAddressLine4']],
        postcode: d['DonorAddressPostCode'], country: s(d['DonorAddressCountry']), dateRaw: d['DonationDate'],
      }
    },
  },
  {
    name: 'MuslimGiving',
    headings: ['Is Donation GA Eligible', 'Base Donation Amount', 'UK Tax Payer'],
    claimsGiftAid: false,
    map: (d) => ({
      id: d['Donation Id'], kindOk: null,
      giftAid: yes(d['Is Donation GA Eligible']) && yes(d['UK Tax Payer']), // BOTH must be yes
      amount: num(d['Base Donation Amount']), currency: 'GBP',
      first: s(d['Donor FirstName']), last: s(d['Donor LastName']), billing: '', title: s(d['Donor Title']),
      lines: [d['Donor Address Line 1'], d['Donor Address Line 2']], postcode: d['Donor Postcode'],
      country: '', dateRaw: d['Date of Donation'],
    }),
  },
  {
    name: 'Fundraise Up',
    headings: ['Donation Status', 'Supporter First Name', 'Donation Amount Before Fees Covered Option'],
    claimsGiftAid: false,
    map: (d) => {
      const st = s(d['Donation Status'])
      let bad: string | null = st.toLowerCase() === 'succeeded' ? null : `Status ${st || 'blank'}` // Failed, Retrying, Pending, Refunded
      if (!bad && s(d['On Behalf Of'])) bad = `Given on behalf of an organisation (${s(d['On Behalf Of'])}) - Gift Aid is for individuals only`
      return {
        id: d['Donation ID'], kindOk: bad, giftAid: yes(d['Gift Aid']),
        amount: num(d['Donation Amount Before Fees Covered Option']), currency: s(d['Donation Currency']) || 'GBP',
        gbpEst: num(d['Converted Donation Amount Before Fees Covered Option']),
        first: s(d['Supporter First Name']), last: s(d['Supporter Last Name']), billing: '', title: '',
        lines: [d['Mailing Address Line 1'], d['Mailing Address Line 2']], postcode: d['Mailing Zip/Postal'],
        country: s(d['Mailing Country Code']) || s(d['Mailing Country']), dateRaw: d['Donation Date'],
      }
    },
  },
  {
    name: 'Donations Manager',
    headings: ['Charity Donation Ref', 'Donor Forename', 'Donation Amount Original'],
    claimsGiftAid: false,
    map: (d) => ({
      id: d['Charity Donation Ref'] || d['Donation Ref'],
      kindOk: s(d['Donor Organisation']) ? `Given on behalf of an organisation (${s(d['Donor Organisation'])}) - Gift Aid is for individuals only` : null,
      giftAid: num(col(d, 'Gift Aid Amount')) > 0, // no yes/no column: a Gift Aid amount means a declaration was made
      amount: num(d['Donation Amount Original']), currency: s(d['Donation Currency']) || 'GBP', gbpEst: num(col(d, 'Donation Amount (')),
      first: s(d['Donor Forename']), last: s(d['Donor Surname']), billing: '', title: s(d['Donor Title']),
      lines: [d['Donor Address Line 1'], d['Donor Address Line 2']], postcode: d['Donor Postcode'],
      country: s(d['Donor Country']), dateRaw: d['Donation Date'],
    }),
  },
  {
    name: 'Cash book (HMRC layout)',
    headings: ['House name or number', 'Donation date', 'Amount'],
    claimsGiftAid: false,
    map: (d) => ({
      id: null, kindOk: null, giftAid: true, amount: num(d['Amount']), currency: 'GBP',
      first: s(d['First name']), last: s(d['Last name']), billing: '', title: s(d['Title']),
      lines: [d['House name or number']], postcode: d['Postcode'], country: '', dateRaw: d['Donation date'],
    }),
  },
]

/** The platform whose headings are all present, or null (=> use the column-mapping screen). */
export function detectProfile(headings: string[]): Profile | null {
  const set = new Set(headings)
  return PROFILES.find((p) => p.headings.every((h) => set.has(h))) ?? null
}

/** JustGiving claims Gift Aid for the charity itself - its files must be refused. Headings only. */
export function isJustGiving(headings: string[]): boolean {
  return headings.some((h) => h.toLowerCase().includes('justgiving'))
}
