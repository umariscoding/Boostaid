import { PDFDocument, StandardFonts, rgb, type PDFFont, type PDFPage } from 'pdf-lib'

const gbp = (n: number) => '£' + n.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
const ukDate = (d: Date | string | null) => {
  if (!d) return ''
  const x = typeof d === 'string' ? new Date(d + 'T12:00:00Z') : d
  return x.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/London' })
}
// The standard PDF fonts only cover Latin-1: anything else (e.g. Arabic names) is shown as ?
const safe = (s: string) => s.replace(/[^\x20-\x7E -ÿ]/g, '?')

export interface InvoiceData {
  number: string
  sentAt: Date
  dueAt: Date
  charity: string
  claimRef: string
  hmrcRef: string | null
  expected: number
  received: number | null
  recoveredGiftAid: number
  recoveredPaid: number
  fee: number
  donations: { date: string | null; donor: string; amount: number }[]
}

/** The business details printed on invoices. Set these in the environment; sensible defaults are used until then. */
export const business = () => ({
  name: process.env.INVOICE_BUSINESS_NAME || 'BoostAid',
  address: process.env.INVOICE_BUSINESS_ADDRESS || '',
  bank: process.env.INVOICE_BANK_DETAILS || '',
  vat: process.env.INVOICE_VAT_NUMBER || '',
})

/** A simple one-invoice PDF: the fee, how it was worked out, and the list of recovered donations. */
export async function buildInvoicePdf(d: InvoiceData): Promise<Uint8Array> {
  const pdf = await PDFDocument.create()
  const font = await pdf.embedFont(StandardFonts.Helvetica)
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold)
  const ink = rgb(0.09, 0.08, 0.17)
  const muted = rgb(0.36, 0.35, 0.45)
  const brand = rgb(0.3, 0.11, 0.58)
  const W = 595, H = 842, M = 50
  let page: PDFPage = pdf.addPage([W, H])
  let y = H - M

  const text = (t: string, x: number, yy: number, o: { size?: number; f?: PDFFont; color?: ReturnType<typeof rgb> } = {}) =>
    page.drawText(safe(t), { x, y: yy, size: o.size ?? 10, font: o.f ?? font, color: o.color ?? ink })
  const right = (t: string, xr: number, yy: number, o: { size?: number; f?: PDFFont; color?: ReturnType<typeof rgb> } = {}) => {
    const w = (o.f ?? font).widthOfTextAtSize(safe(t), o.size ?? 10)
    text(t, xr - w, yy, o)
  }
  const line = (yy: number) => page.drawLine({ start: { x: M, y: yy }, end: { x: W - M, y: yy }, thickness: 0.6, color: rgb(0.85, 0.84, 0.9) })

  const b = business()
  text('Invoice', M, y - 6, { size: 26, f: bold, color: brand })
  right(b.name, W - M, y - 2, { size: 14, f: bold })
  let by = y - 18
  for (const l of [b.address, b.vat ? `VAT no. ${b.vat}` : ''].filter(Boolean)) { right(l, W - M, by, { size: 9 }); by -= 12 }
  y -= 56

  text('Invoice number', M, y, { size: 9, color: muted }); text(d.number, M, y - 14, { f: bold })
  text('Date', 220, y, { size: 9, color: muted }); text(ukDate(d.sentAt), 220, y - 14)
  text('Due', 360, y, { size: 9, color: muted }); text(ukDate(d.dueAt), 360, y - 14)
  y -= 48
  text('Billed to', M, y, { size: 9, color: muted }); text(d.charity, M, y - 14, { f: bold })
  text('Claim', 360, y, { size: 9, color: muted }); text(`${d.claimRef}${d.hmrcRef ? ` (HMRC ref ${d.hmrcRef})` : ''}`, 360, y - 14)
  y -= 50
  line(y)
  y -= 20
  text('Description', M, y, { size: 9, f: bold, color: muted }); right('Amount', W - M, y, { size: 9, f: bold })
  y -= 20
  text('Gift Aid recovery fee (5%)', M, y, { f: bold })
  right(gbp(d.fee), W - M, y, { f: bold })
  y -= 16
  const explain = [
    `5% of the Gift Aid BoostAid recovered by emailing donors, on claim ${d.claimRef}.`,
    `Gift Aid recovered: ${gbp(d.recoveredGiftAid)}.${d.received != null && d.received < d.expected ? ` HMRC paid ${gbp(d.received)} against ${gbp(d.expected)} claimed,` : ''}`,
    `${d.received != null && d.received < d.expected ? 'so the recovered part paid is ' : 'Recovered Gift Aid paid by HMRC: '}${gbp(d.recoveredPaid)}.`,
  ]
  for (const l of explain) { text(l, M, y, { size: 9, color: muted }); y -= 12 }
  y -= 8
  line(y)
  y -= 22
  text('Total due', M, y, { size: 12, f: bold }); right(gbp(d.fee), W - M, y, { size: 14, f: bold, color: brand })
  y -= 30
  const pay = ['Payment by bank transfer, within 14 days of the invoice date.', b.bank ? `Bank details: ${b.bank}` : '', `Please use ${d.number} as the payment reference.`].filter(Boolean)
  for (const l of pay) { text(l, M, y, { size: 9, color: muted }); y -= 12 }

  // The recovered donations, so the charity can see exactly what the fee is on
  y -= 18
  const header = () => {
    text('Recovered donations', M, y, { size: 11, f: bold }); y -= 18
    text('Date', M, y, { size: 9, f: bold, color: muted }); text('Donor', 140, y, { size: 9, f: bold, color: muted })
    right('Donation', 440, y, { size: 9, f: bold }); right('Gift Aid', W - M, y, { size: 9, f: bold }); y -= 6; line(y); y -= 14
  }
  header()
  for (const r of d.donations) {
    if (y < M + 30) { page = pdf.addPage([W, H]); y = H - M; header() }
    text(ukDate(r.date).replace(/ 20(\d\d)$/, ' 20$1'), M, y, { size: 9 }); text(r.donor.slice(0, 40), 140, y, { size: 9 })
    right(gbp(r.amount), 440, y, { size: 9 }); right(gbp(Math.round(r.amount * 25) / 100), W - M, y, { size: 9 })
    y -= 14
  }
  return pdf.save()
}
