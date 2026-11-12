import * as XLSX from 'xlsx'
import type { Change, ClaimResult, HmrcFilePlan } from '@/lib/giftaid'

const ddmmyyyy = (iso: string | null) => (iso ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}` : '')

export interface SimpleSheet {
  plan: HmrcFilePlan
  data: Uint8Array
}

/** Optional: each HMRC file's clean data as a plain Excel sheet (no HMRC template), for backup or your own records. */
export function buildSimpleSheets(result: ClaimResult): SimpleSheet[] {
  return result.files.map((plan) => {
    const rows = plan.rows.map((r) => ({
      Title: r.title,
      'First name': r.firstName,
      'Last name': r.lastName,
      'House name or number': r.house,
      Postcode: r.postcode,
      'Donation date': ddmmyyyy(r.date),
      Amount: r.amount,
      Platform: r.platform,
      'Donation ID': r.donationId === null || r.donationId === undefined ? '' : String(r.donationId),
    }))
    const ws = XLSX.utils.json_to_sheet(rows)
    ws['!cols'] = [8, 16, 16, 30, 12, 14, 12, 20, 22].map((wch) => ({ wch }))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
    return { plan, data: new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' })) }
  })
}

/** UK_schedule_1_of_2.ods -> "UK_sheet_1_of_2 - £60,446 - GA-2026-0009.xlsx" */
export function simpleSheetName(plan: HmrcFilePlan, ref: string): string {
  const total = '£' + plan.total.toLocaleString('en-GB', { minimumFractionDigits: Number.isInteger(plan.total) ? 0 : 2, maximumFractionDigits: 2 })
  return `${plan.name.replace(/\.ods$/i, '').replace('_schedule_', '_sheet_')} - ${total} - ${ref}.xlsx`
}

/** What was changed on each donation, keyed by the platform's donation ID. */
export function changesByDonation(result: ClaimResult): Map<string, Change[]> {
  const map = new Map<string, Change[]>()
  for (const c of result.changes) {
    if (c.donationId === null || c.donationId === undefined || String(c.donationId) === '') continue
    const key = String(c.donationId)
    map.set(key, [...(map.get(key) ?? []), c])
  }
  return map
}
