/**
 * The claim summary (3_Claim_summary.pdf).
 *   Page 1 - "Boost Aid - Gift Aid Summary": plain numbers for the finance officer / trustees.
 *   Page 2 - Audit detail: source files, HMRC files, review reasons.
 *
 * Returned as an HTML document. In the app: show it in an iframe and offer "Print / Save as PDF"
 * (window.print), like the current summary screen does. On a server it can be rendered to PDF
 * with a headless browser.
 */
import type { ClaimResult } from './engine'
import { toDDMMYY } from './dates'

export interface Fingerprint { name: string; size: number; sha256: string }

const gbp = (x: number) => '£' + x.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
const int = (x: number) => x.toLocaleString('en-GB')
const esc = (x: string) => x.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

/** The eight headline lines (label, value, plain-English explanation). Used by the PDF and the workbook. */
export function summaryLines(res: ClaimResult): [string, string, string][] {
  const t = res.totals
  const claimed = res.ready.length + res.overseas.length
  return [
    ['Total donations reviewed', gbp(t.reviewedDonations), 'every real donation in the uploaded files, with or without Gift Aid'],
    ['Data lines', int(res.linesRead), 'rows read from the files'],
    ['Records cleaned', int(t.cleanedRows), 'claimed donations where BoostAid corrected a name, house number or postcode'],
    ['Valid records (claimed)', int(claimed), 'donations in the HMRC files'],
    ['Total Gift Aid value', gbp(t.ukDonations + t.overseasDonations), 'donations being claimed on'],
    ['Estimated Gift Aid reclaimable', gbp(t.giftAid), '25% of the above'],
    ['Records needing attention', int(res.review.length), 'held back - see the Needs Review tab for the reason on each row'],
    ['Compliance rate', `${Math.round(t.compliancePct)}%`, 'claimed ÷ (claimed + needing attention)'],
  ]
}

export function buildSummaryHtml(res: ClaimResult, generated: string): string {
  const claimed = [...res.ready, ...res.overseas]
  const platformRows = res.platforms.length > 1
    ? `<h2>Gift Aid by platform</h2><table class=simple>${res.platforms.map((p) => {
        const rows = claimed.filter((r) => r.platform === p)
        return `<tr><td>${esc(p)}</td><td class="n b">${gbp(rows.reduce((a, r) => a + r.amount, 0) * 0.25)}</td><td class=muted>${int(rows.length)} donations</td></tr>`
      }).join('')}</table>` : ''
  const count = (rows: { source: string }[], table: string) => rows.filter((r) => r.source.startsWith(table + '!')).length
  return `<!doctype html><html><head><meta charset="utf-8"><title>Gift Aid claim summary</title><style>
body{font-family:Arial,sans-serif;font-size:10pt;color:#17152B;margin:28px} h1{font-size:17pt;margin:0} h2{font-size:11pt;margin:16px 0 6px;color:#4C1D95}
table{border-collapse:collapse;width:100%} td,th{border-bottom:1px solid #ddd;padding:4px 6px;text-align:left} th{background:#F1EFF6} .n{text-align:right}
.simple td{padding:7px 8px;font-size:11pt} .b{font-weight:bold} .hl{color:#146C3A;font-size:14pt} .muted{color:#5B5874}
@media print{.page2{page-break-before:always}}
</style></head><body>
<h1 style="color:#4C1D95">Boost Aid - Gift Aid Summary</h1><div class=muted>${esc(res.platforms.join(' + '))} · generated ${generated}</div>
<table class=simple>${summaryLines(res).map(([l, v, n]) => `<tr><td>${l}</td><td class="n b${l.startsWith('Estimated') ? ' hl' : ''}">${v}</td><td class=muted>${n}</td></tr>`).join('')}</table>
${platformRows}
<h2>Next steps</h2><ol><li>Upload the ${res.files.length} file(s) in <b>1_Submit_to_HMRC</b> to Charities Online.</li><li>Check the ${int(res.review.length)} records needing attention (Needs Review tab) - fix and include them in a later claim.</li><li>When HMRC pays, record the date and amount.</li></ol>
<div class=page2><h1>Audit detail</h1><div class=muted>Box 1 (earliest donation): ${res.box1Date ? toDDMMYY(res.box1Date) : '-'}</div>
<h2>Source files</h2><table><tr><th>File</th><th>Platform</th><th class=n>Lines</th><th class=n>Claimed</th><th class=n>Review</th><th class=n>Not eligible</th></tr>
${res.sheets.map((sh) => `<tr><td>${esc(sh.table)}</td><td>${esc(sh.platform)}</td><td class=n>${int(sh.rowsUsed)}</td><td class=n>${int(count(claimed, sh.table))}</td><td class=n>${int(count(res.review, sh.table))}</td><td class=n>${int(count(res.ineligible, sh.table))}</td></tr>`).join('')}</table>
<h2>HMRC files</h2><table><tr><th>File</th><th class=n>Rows</th><th class=n>Donations</th><th class=n>Gift Aid</th><th>Dates</th></tr>
${res.files.map((f) => `<tr><td>${f.name}</td><td class=n>${int(f.rows.length)}</td><td class=n>${gbp(f.total)}</td><td class=n>${gbp(f.total * 0.25)}</td><td>${toDDMMYY(f.firstDate)} – ${toDDMMYY(f.lastDate)}</td></tr>`).join('')}</table>
<h2>Main review reasons</h2><ul>${res.reviewReasons.slice(0, 6).map(([k, n]) => `<li>${esc(k)}: ${n}</li>`).join('') || '<li>None</li>'}</ul>
<p class=muted>Every claimed donation is listed in the claim workbook with its HMRC file and row, platform, donation ID and original file row. Original exports and their SHA-256 fingerprints are in 4_Original_exports.</p></div>
</body></html>`
}
