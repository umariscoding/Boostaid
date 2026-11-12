/**
 * The claim pack - one zip per claim, the same layout as the example packs:
 *
 *   1_Submit_to_HMRC/UK_schedule_1_of_8.ods …   (+ Overseas_schedule_… if any)
 *   2_Claim_workbook.xlsx
 *   3_Claim_summary.html                        (open + "Save as PDF"; the app can also render it to PDF)
 *   4_Original_exports/<files exactly as uploaded> + Fingerprints_SHA256.txt
 */
import JSZip from 'jszip'
import { readTables, sha256Hex } from './read'
import { buildClaim, type BuildOptions, type ClaimResult } from './engine'
import { writeHmrcOds } from './ods'
import { buildClaimWorkbook } from './workbook'
import { buildSummaryHtml, type Fingerprint } from './summary'

export interface UploadedFile { name: string; data: ArrayBuffer | Uint8Array }

export interface ClaimPack {
  result: ClaimResult
  fingerprints: Fingerprint[]
  summaryHtml: string
  zip: Uint8Array
}

/**
 * Everything in one call: uploaded files -> cleaned claim -> zip.
 * `templateOds` = bytes of public/hmrc_template.ods (in the browser: `await (await fetch('/hmrc_template.ods')).arrayBuffer()`).
 */
export async function buildClaimPack(files: UploadedFile[], templateOds: ArrayBuffer | Uint8Array, opts: BuildOptions = {}): Promise<ClaimPack> {
  const tables = files.flatMap((f) => readTables(f.name, f.data))
  const result = buildClaim(tables, opts)
  const fingerprints: Fingerprint[] = []
  for (const f of files) fingerprints.push({ name: f.name, size: f.data.byteLength, sha256: await sha256Hex(f.data) })
  const generated = new Date().toLocaleString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }).replace(',', '')

  const zip = new JSZip()
  // Box 1 = earliest donation of that kind (UK files share one date, overseas files another)
  const earliest = (kind: 'UK' | 'Overseas') => res(kind).map((f) => f.firstDate).sort()[0]
  const res = (kind: 'UK' | 'Overseas') => result.files.filter((f) => f.kind === kind)
  for (const f of result.files) zip.file(`1_Submit_to_HMRC/${f.name}`, await writeHmrcOds(templateOds, f, earliest(f.kind)))
  zip.file('2_Claim_workbook.xlsx', await buildClaimWorkbook(result, fingerprints, generated))
  const summaryHtml = buildSummaryHtml(result, generated)
  zip.file('3_Claim_summary.html', summaryHtml)
  for (const f of files) zip.file(`4_Original_exports/${f.name}`, f.data)
  zip.file('4_Original_exports/Fingerprints_SHA256.txt',
    'Original files exactly as uploaded. Re-compute SHA-256 to prove they are unchanged.\n\n' +
    fingerprints.map((f) => `${f.sha256}  ${f.name}  (${f.size} bytes)`).join('\n') + '\n')

  return { result, fingerprints, summaryHtml, zip: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }) }
}
