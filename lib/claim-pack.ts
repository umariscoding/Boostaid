import JSZip from 'jszip'
import {
  buildClaimWorkbook,
  buildSummaryHtml,
  sha256Hex,
  writeHmrcOds,
  type ClaimResult,
  type HmrcFilePlan,
} from '@/lib/giftaid'
import type { Upload } from '@/lib/claim-flow'

export interface Schedule {
  plan: HmrcFilePlan
  /** The ODS file exactly as the engine plans it, e.g. UK_schedule_1_of_2.ods */
  data: Uint8Array
}

export interface NamedSchedule extends Schedule {
  ref: string
  /** Download name: schedule name + the schedule's total + our claim number */
  name: string
}

const money = (n: number) =>
  '£' + n.toLocaleString('en-GB', { minimumFractionDigits: Number.isInteger(n) ? 0 : 2, maximumFractionDigits: 2 })

/** UK_schedule_1_of_2.ods -> "UK_schedule_1_of_2 - £52,851 - GA-2026-0001.ods" */
export function scheduleName(plan: HmrcFilePlan, ref: string): string {
  return `${plan.name.replace(/\.ods$/i, '')} - ${money(plan.total)} - ${ref}.ods`
}

/** Fills HMRC's own template for every planned file. Box 1 is the earliest date of that kind (UK / overseas). */
export async function buildSchedules(result: ClaimResult, templateOds: Uint8Array): Promise<Schedule[]> {
  const earliest = (kind: 'UK' | 'Overseas') =>
    result.files.filter((f) => f.kind === kind).map((f) => f.firstDate).sort()[0]
  const out: Schedule[] = []
  for (const plan of result.files) out.push({ plan, data: await writeHmrcOds(templateOds, plan, earliest(plan.kind)) })
  return out
}

export interface AssembledPack {
  zip: Uint8Array
  workbook: Uint8Array
  summaryHtml: string
}

/** The claim pack: the schedules to submit, the review workbook, the summary, and the original files with fingerprints. */
export async function assemblePack(result: ClaimResult, uploads: Upload[], schedules: NamedSchedule[]): Promise<AssembledPack> {
  const fingerprints = []
  for (const u of uploads) fingerprints.push({ name: u.name, size: u.data.byteLength, sha256: await sha256Hex(u.data) })
  const generated = new Date()
    .toLocaleString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' })
    .replace(',', '')

  const workbook = await buildClaimWorkbook(result, fingerprints, generated)
  const summaryHtml = buildSummaryHtml(result, generated)

  const zip = new JSZip()
  for (const s of schedules) zip.file(`1_Submit_to_HMRC/${s.name}`, s.data)
  zip.file('2_Claim_workbook.xlsx', workbook)
  zip.file('3_Claim_summary.html', summaryHtml)
  for (const u of uploads) zip.file(`4_Original_exports/${u.name}`, u.data)
  zip.file(
    '4_Original_exports/Fingerprints_SHA256.txt',
    'Original files exactly as uploaded. Re-compute SHA-256 to prove they are unchanged.\n\n' +
      fingerprints.map((f) => `${f.sha256}  ${f.name}  (${f.size} bytes)`).join('\n') +
      '\n'
  )
  zip.file(
    '5_Claim_numbers.txt',
    'Each HMRC schedule and the BoostAid claim number it was saved under:\n\n' +
      schedules.map((s) => `${s.ref}  ${s.plan.name}  ${s.plan.rows.length} donations  ${money(s.plan.total)}`).join('\n') +
      '\n'
  )
  return { zip: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), workbook, summaryHtml }
}
