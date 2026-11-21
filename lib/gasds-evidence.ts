import * as XLSX from 'xlsx'
import JSZip from 'jszip'
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib'
import type { GasdsClaimDto } from '@/lib/gasds-db'
import { countText, p, showNo, taxYearLabel, ukDate, who, type GReceipt } from '@/lib/gasds'
import { RULES_TEXT, evidenceSummary } from '@/lib/gasds-summary'

const safe = (s: string) => s.replace(/[^\x20-\x7E -ÿ]/g, '?')

/** One page: the summary card and the rules. */
export async function buildEvidencePdf(c: GasdsClaimDto, charity: string): Promise<Uint8Array> {
  const pdf = await PDFDocument.create()
  const font = await pdf.embedFont(StandardFonts.Helvetica)
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold)
  const page = pdf.addPage([595, 842])
  const M = 50
  let y = 842 - M
  const t = (s: string, x: number, yy: number, o: { size?: number; f?: typeof font; color?: ReturnType<typeof rgb> } = {}) =>
    page.drawText(safe(s).replace('✓', 'OK').replace('✗', 'NOT ENOUGH'), { x, y: yy, size: o.size ?? 11, font: o.f ?? font, color: o.color ?? rgb(0.09, 0.08, 0.17) })
  t('Small donations (GASDS): evidence', M, y - 12, { size: 20, f: bold, color: rgb(0.3, 0.11, 0.58) })
  t(`${charity} · claim for ${taxYearLabel(c.taxYear)} · HMRC ref ${c.hmrcRef}`, M, y - 34, { size: 10, color: rgb(0.36, 0.35, 0.45) })
  y -= 70
  page.drawLine({ start: { x: M, y: y + 12 }, end: { x: 545, y: y + 12 }, thickness: 0.6, color: rgb(0.85, 0.84, 0.9) })
  for (const [k, v] of evidenceSummary(c, charity)) {
    t(k, M, y)
    const w = bold.widthOfTextAtSize(safe(v).replace('✓', 'OK').replace('✗', 'NOT ENOUGH'), 11)
    t(v, 545 - w, y, { f: bold })
    y -= 24
    page.drawLine({ start: { x: M, y: y + 14 }, end: { x: 545, y: y + 14 }, thickness: 0.3, color: rgb(0.9, 0.9, 0.94) })
  }
  y -= 14
  // wrap the rules text
  const words = RULES_TEXT.split(' ')
  let line = ''
  for (const w of words) {
    if (font.widthOfTextAtSize(line + ' ' + w, 9) > 495) { t(line, M, y, { size: 9, color: rgb(0.36, 0.35, 0.45) }); y -= 13; line = w } else line = line ? line + ' ' + w : w
  }
  if (line) t(line, M, y, { size: 9, color: rgb(0.36, 0.35, 0.45) })
  t('The receipts behind this claim are in the Excel file next to this page.', M, y - 22, { size: 9, color: rgb(0.36, 0.35, 0.45) })
  return pdf.save()
}

/** The table of receipts in the claim, as Excel. */
export function buildEvidenceXlsx(receipts: GReceipt[], total: number): Uint8Array {
  const rows = receipts.map((r) => ({
    Date: ukDate(r.date),
    Receipt: showNo(r.no),
    From: who(r),
    Type: r.gift ? 'Individual' : 'Collection',
    'Paid by': r.card ? 'Card' : 'Cash',
    Count: countText(r),
    'Bank slip': r.slip ?? (r.gift ? 'Banked' : ''),
    'Bank counted (GBP)': r.counted != null && p(r.counted) !== p(r.amount) ? r.counted : '',
    'Receipt (GBP)': r.amount,
    'Used in claim (GBP)': r.used ?? 0,
  }))
  const sheet = XLSX.utils.json_to_sheet(rows)
  sheet['!cols'] = [12, 14, 30, 12, 9, 44, 14, 18, 14, 18].map((wch) => ({ wch }))
  XLSX.utils.sheet_add_aoa(sheet, [['Total', '', '', '', '', '', '', '', '', total]], { origin: -1 })
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, sheet, 'Receipts in claim')
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as Uint8Array
}

export async function buildEvidenceZip(
  c: GasdsClaimDto,
  charity: string,
  receipts: GReceipt[],
  photos: { name: string; mime: string; data: Uint8Array }[]
): Promise<Uint8Array> {
  const zip = new JSZip()
  const base = `GASDS_${taxYearLabel(c.taxYear).replace('/', '-')}_${c.hmrcRef.replace(/[^\w-]/g, '_')}`
  zip.file(`${base}_summary.pdf`, await buildEvidencePdf(c, charity))
  zip.file(`${base}_receipts.xlsx`, buildEvidenceXlsx(receipts, c.amount))
  for (const ph of photos) {
    const ext = ph.mime === 'application/pdf' ? 'pdf' : ph.mime.split('/')[1] ?? 'bin'
    zip.file(`count_sheets/${ph.name.replace(/[^\w-]/g, '_')}.${ext}`, ph.data)
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
