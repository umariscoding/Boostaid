/**
 * The claim workbook (2_Claim_workbook.xlsx) - what the finance officer and auditor open.
 *
 * Tabs: Summary (BoostAid summary) · Audit detail · Ready - UK · Ready - Overseas ·
 *       Needs Review · Not Eligible · Changes Log
 */
import ExcelJS from 'exceljs'
import type { ClaimResult, ClaimRow } from './engine'
import { toDDMMYY } from './dates'
import { s } from './text'
import { summaryLines, type Fingerprint } from './summary'

type Col = { header: string; get: (r: any) => unknown; width?: number }

function addTable(wb: ExcelJS.Workbook, name: string, rows: any[], cols: Col[], colour: string) {
  const ws = wb.addWorksheet(name)
  ws.columns = cols.map((c) => ({ header: c.header, width: c.width ?? Math.max(12, Math.min(30, c.header.length + 4)) }))
  const head = ws.getRow(1)
  head.font = { bold: true, color: { argb: 'FFFFFFFF' } }
  head.eachCell((cell) => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + colour } } })
  for (const r of rows) ws.addRow(cols.map((c) => c.get(r)))
  ws.views = [{ state: 'frozen', ySplit: 1 }]
}

const claimCols: Col[] = [
  { header: 'HMRC file', get: (r: ClaimRow) => r.hmrcFile, width: 28 },
  { header: 'HMRC row', get: (r: ClaimRow) => r.hmrcRow },
  { header: 'Title', get: (r: ClaimRow) => r.title || null },
  { header: 'First Name', get: (r: ClaimRow) => r.firstName },
  { header: 'Last Name', get: (r: ClaimRow) => r.lastName },
  { header: 'House', get: (r: ClaimRow) => r.house },
  { header: 'Postcode', get: (r: ClaimRow) => r.postcode },
  { header: 'Date', get: (r: ClaimRow) => (r.date ? toDDMMYY(r.date) : null) },
  { header: 'Amount', get: (r: ClaimRow) => r.amount },
  { header: 'Platform', get: (r: ClaimRow) => r.platform },
  { header: 'Donation ID', get: (r: ClaimRow) => r.donationId },
  { header: 'Source file & row', get: (r: ClaimRow) => r.source, width: 40 },
]

export async function buildClaimWorkbook(res: ClaimResult, fingerprints: Fingerprint[], generated: string): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook()

  // --- Summary (the finance officer's one-glance view)
  const sm = wb.addWorksheet('Summary')
  sm.columns = [{ width: 34 }, { width: 22 }, { width: 90 }]
  sm.addRow(['Boost Aid - Gift Aid Summary']).font = { bold: true, size: 16, color: { argb: 'FF4C1D95' } }
  sm.addRow([`${res.platforms.join(' + ')} · generated ${generated}`])
  sm.addRow([])
  for (const [label, value, note] of summaryLines(res)) {
    const row = sm.addRow([label, value, note])
    row.getCell(2).font = label.startsWith('Estimated') ? { bold: true, size: 14, color: { argb: 'FF146C3A' } } : { bold: true, size: 12 }
  }
  if (res.platforms.length > 1) {
    sm.addRow([]); sm.addRow(['Gift Aid by platform']).font = { bold: true }
    for (const p of res.platforms) {
      const rows = [...res.ready, ...res.overseas].filter((r) => r.platform === p)
      const t = rows.reduce((a, r) => a + r.amount, 0)
      sm.addRow([p, `£${fmt(t * 0.25)}`, `${rows.length.toLocaleString('en-GB')} donations, £${fmt(t)}`])
    }
  }
  sm.addRow([])
  sm.addRow(['Files to upload to HMRC', String(res.files.length), res.files.map((f) => f.name).join(', ')])
  sm.addRow(['More detail', '', 'Audit detail tab: source files, fingerprints, totals per HMRC file, reasons'])

  // --- Audit detail
  const ad = wb.addWorksheet('Audit detail')
  ad.columns = [{ width: 62 }, { width: 22 }, { width: 18 }, { width: 22 }, { width: 14 }, { width: 14 }, { width: 14 }]
  const bold = (r: ExcelJS.Row) => { r.font = { bold: true } }
  ad.addRow([`BoostAid cleaning summary - ${res.platforms.join(' + ')}`]).font = { bold: true, size: 14 }
  ad.addRow([`Generated ${generated}`]); ad.addRow([])
  bold(ad.addRow(['Source file / sheet', 'Platform', 'Lines read', 'Exact copies skipped', 'Ready', 'Needs review', 'Not eligible']))
  for (const sh of res.sheets) {
    const count = (rows: { source: string }[]) => rows.filter((r) => r.source.startsWith(sh.table + '!')).length
    ad.addRow([sh.table, sh.dateOrder ? `${sh.platform} - dates ${sh.dateOrder}` : sh.platform, sh.rowsUsed, sh.copiesSkipped,
      count(res.ready) + count(res.overseas), count(res.review), count(res.ineligible)])
  }
  ad.addRow([]); bold(ad.addRow(['Original file', 'Size (bytes)', 'Fingerprint (SHA-256) - proves the file was not changed']))
  for (const f of fingerprints) ad.addRow([f.name, f.size, f.sha256])
  ad.addRow([]); bold(ad.addRow(['Group', 'Rows', 'Donations (£)', 'Gift Aid at 25% (£)']))
  const t = res.totals
  ad.addRow(['Ready for HMRC (UK)', res.ready.length, r2(t.ukDonations), r2(t.ukDonations * 0.25)])
  ad.addRow(['Ready for HMRC (overseas, separate file)', res.overseas.length, r2(t.overseasDonations), r2(t.overseasDonations * 0.25)])
  ad.addRow(['Needs review', res.review.length, r2(t.reviewDonations), r2(t.reviewDonations * 0.25)])
  ad.addRow(['Not eligible', res.ineligible.length])
  for (const [label, list] of [['Needs review - reasons', res.reviewReasons], ['Not eligible - reasons', res.ineligibleReasons]] as const) {
    ad.addRow([]); bold(ad.addRow([label, 'Rows']))
    for (const [k, n] of list) ad.addRow([k, n])
  }
  ad.addRow([]); bold(ad.addRow(['HMRC file', 'Rows', 'Donations (£)', 'Gift Aid at 25% (£)', 'First date', 'Last date']))
  for (const f of res.files) ad.addRow([f.name, f.rows.length, f.total, r2(f.total * 0.25), toDDMMYY(f.firstDate), toDDMMYY(f.lastDate)])

  // --- detail tabs
  addTable(wb, 'Ready - UK', res.ready, claimCols, '146C3A')
  addTable(wb, 'Ready - Overseas', res.overseas, claimCols, '1E4FA8')
  addTable(wb, 'Needs Review', res.review, [
    { header: 'Reason', get: (r) => r.reason, width: 50 },
    { header: 'Suggested name (billing name - check it is the donor)', get: (r) => r.suggestedName || null, width: 30 },
    ...claimCols.filter((c) => !c.header.startsWith('HMRC')),
    { header: 'Original amount', get: (r) => r.originalAmount },
    { header: 'Original name', get: (r) => r.originalName },
    { header: 'Original address', get: (r) => r.originalAddress, width: 50 },
    { header: 'Original postcode', get: (r) => r.originalPostcode },
    { header: 'Country', get: (r) => r.country },
  ], '9B1C13')
  addTable(wb, 'Not Eligible', res.ineligible, [
    { header: 'Reason', get: (r) => r.reason, width: 50 }, { header: 'Name', get: (r) => r.name || null },
    { header: 'Amount', get: (r) => r.amount }, { header: 'Donation ID', get: (r) => r.donationId },
    { header: 'Source file & row', get: (r) => r.source, width: 40 },
  ], '5B5874')
  addTable(wb, 'Changes Log', res.changes, [
    { header: 'Donation ID', get: (r) => r.donationId }, { header: 'Field', get: (r) => r.field },
    { header: 'Original', get: (r) => (s(r.original) || null), width: 40 }, { header: 'Cleaned', get: (r) => r.cleaned, width: 40 },
  ], '5B5874')

  return new Uint8Array(await wb.xlsx.writeBuffer())
}

const r2 = (x: number) => Math.round(x * 100) / 100
const fmt = (x: number) => x.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
