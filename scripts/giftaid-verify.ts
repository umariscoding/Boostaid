/**
 * Acceptance test for the Gift Aid engine.
 *
 *   npx tsx scripts/giftaid-verify.ts <folder with the test exports> [--packs <out folder>] [--baseline <folder>]
 *
 * 1. Runs every address case in lib/giftaid/__tests__/spot_checks.json.
 * 2. Runs each test export (and a 3-platform combined claim) and checks the TARGET numbers
 *    from the handover README.
 * 3. --baseline: compares every claimed / review row with the reference implementation's
 *    workbooks (baseline/<Name>/BoostAid_cleaned.xlsx), field by field.
 * 4. --packs: writes a full claim pack (zip) for each test into that folder.
 *
 * Exit code 0 = everything matches.
 */
import fs from 'node:fs'
import path from 'node:path'
import * as XLSX from 'xlsx'
import { houseFrom, buildClaim, readTables, buildClaimPack } from '../lib/giftaid'

const args = process.argv.slice(2)
const dir = args[0]
const flag = (f: string) => (args.includes(f) ? args[args.indexOf(f) + 1] : undefined)
const packsDir = flag('--packs'), baselineDir = flag('--baseline')
if (!dir) { console.error('usage: npx tsx scripts/giftaid-verify.ts <test files folder> [--packs out] [--baseline folder]'); process.exit(2) }

let failures = 0
const fail = (msg: string) => { failures++; console.log('  ✗ ' + msg) }

// ------------------------------------------------------------------ 1. address spot checks
const spot = JSON.parse(fs.readFileSync(path.join(__dirname, '../lib/giftaid/__tests__/spot_checks.json'), 'utf8')).house as [string[], string | null][]
console.log(`Address spot checks (${spot.length})`)
for (const [lines, expected] of spot) {
  const got = houseFrom(lines).house
  if (got !== expected) fail(`${JSON.stringify(lines)} -> ${JSON.stringify(got)}, expected ${JSON.stringify(expected)}`)
}

// ------------------------------------------------------------------ 2. target numbers per test file
// File names are matched by what they END with, so the hashes the upload tool adds don't matter.
const TESTS: { name: string; files: string[]; ready: number; readyTotal: number; overseas: number; review: number; inelig: number }[] = [
  { name: 'LaunchGood', files: ['LG_Test_data.xlsx'], ready: 3589, readyTotal: 296189.0, overseas: 1, review: 260, inelig: 25258 },
  { name: 'Raisely', files: ['Raisely.xlsx'], ready: 3886, readyTotal: 190372.3, overseas: 0, review: 63, inelig: 7092 },
  { name: 'CAF', files: ['CAF_1.xlsx'], ready: 157, readyTotal: 1511.0, overseas: 0, review: 0, inelig: 92 },
  { name: 'MuslimGiving', files: ['MuslimGiving_1.xlsx'], ready: 41, readyTotal: 1031.32, overseas: 0, review: 1, inelig: 33 },
  { name: 'FundraiseUp', files: ['donations-11-11-2026-2-58-PM.csv'], ready: 686, readyTotal: 12273.0, overseas: 0, review: 44, inelig: 1561 },
  { name: 'DonationsManager', files: ['donations-11112026145923-i86f5ad525e0vsr5h4c9.csv'], ready: 483, readyTotal: 8168.0, overseas: 0, review: 8, inelig: 303 },
  { name: 'CashBook', files: ['gift-aid-schedule-SAMPLE.csv'], ready: 10, readyTotal: 615.0, overseas: 0, review: 0, inelig: 0 },
  { name: 'Combined3', files: ['LG_Test_data.xlsx', 'Raisely.xlsx', 'donations-11112026145923-i86f5ad525e0vsr5h4c9.csv'], ready: 7958, readyTotal: 494729.3, overseas: 1, review: 331, inelig: 32653 },
]
const all = fs.readdirSync(dir)
const find = (suffix: string) => all.find((f) => f.endsWith(suffix))

async function main() {
  for (const t of TESTS) {
    const paths = t.files.map(find)
    if (paths.some((p) => !p)) { console.log(`\n${t.name}: test file missing - skipped`); continue }
    const uploads = paths.map((p) => ({ name: p!, data: new Uint8Array(fs.readFileSync(path.join(dir, p!))) }))
    const res = buildClaim(uploads.flatMap((u) => readTables(u.name, u.data)))
    const readyTotal = Math.round(res.totals.ukDonations * 100) / 100
    console.log(`\n${t.name}: ready ${res.ready.length} (£${readyTotal}) · overseas ${res.overseas.length} · review ${res.review.length} · not eligible ${res.ineligible.length} · ${res.files.length} HMRC file(s)`)
    if (res.ready.length !== t.ready) fail(`ready ${res.ready.length}, expected ${t.ready}`)
    if (Math.abs(readyTotal - t.readyTotal) > 0.005) fail(`ready total £${readyTotal}, expected £${t.readyTotal}`)
    if (res.overseas.length !== t.overseas) fail(`overseas ${res.overseas.length}, expected ${t.overseas}`)
    if (res.review.length !== t.review) fail(`review ${res.review.length}, expected ${t.review}`)
    if (res.ineligible.length !== t.inelig) fail(`not eligible ${res.ineligible.length}, expected ${t.inelig}`)

    // -------------------------------------------------------------- 3. row-by-row against the reference workbooks
    const wbPath = baselineDir && path.join(baselineDir, t.name, 'BoostAid_cleaned.xlsx')
    if (wbPath && fs.existsSync(wbPath)) {
      const wb = XLSX.read(fs.readFileSync(wbPath), { type: 'buffer' })
      const sheetRows = (name: string) => XLSX.utils.sheet_to_json<Record<string, unknown>>(wb.Sheets[name] ?? {}, { defval: null, raw: true })
      const str = (v: unknown) => (v === null || v === undefined ? '' : String(v))
      const ddmmyy = (iso: string | null) => (iso ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(2, 4)}` : '')
      const compare = (label: string, ref: Record<string, unknown>[], ours: any[], fields: [string, (r: any) => unknown][]) => {
        const bySource = new Map(ours.map((r) => [r.source, r]))
        let bad = 0
        for (const rr of ref) {
          const o = bySource.get(str(rr['Source file & row']))
          if (!o) { if (bad++ < 5) fail(`${label}: missing ${str(rr['Source file & row'])}`); continue }
          for (const [col, get] of fields) {
            const a = str(rr[col]), b = str(get(o))
            const same = col === 'Amount' ? Math.abs(Number(a || 0) - Number(b || 0)) < 0.005 : a === b
            if (!same) { if (bad++ < 5) fail(`${label} ${str(rr['Source file & row'])} ${col}: reference "${a}" vs ours "${b}"`); break }
          }
        }
        if (ref.length !== ours.length) fail(`${label}: reference has ${ref.length} rows, ours ${ours.length}`)
        console.log(`  ${bad ? '✗' : '✓'} ${label}: ${ref.length} rows compared with the reference${bad ? `, ${bad} differences` : ', all identical'}`)
      }
      const claimFields: [string, (r: any) => unknown][] = [
        ['HMRC file', (r) => r.hmrcFile], ['HMRC row', (r) => r.hmrcRow], ['Title', (r) => r.title], ['First Name', (r) => r.firstName],
        ['Last Name', (r) => r.lastName], ['House', (r) => r.house], ['Postcode', (r) => r.postcode], ['Date', (r) => ddmmyy(r.date)], ['Amount', (r) => r.amount],
      ]
      compare('Ready - UK', sheetRows('Ready - UK'), res.ready, claimFields)
      compare('Ready - Overseas', sheetRows('Ready - Overseas'), res.overseas, claimFields)
      compare('Needs Review', sheetRows('Needs Review'), res.review, [['Reason', (r) => r.reason], ['First Name', (r) => r.firstName], ['Last Name', (r) => r.lastName], ['House', (r) => r.house], ['Postcode', (r) => r.postcode]])
      compare('Not Eligible', sheetRows('Not Eligible'), res.ineligible, [['Reason', (r) => r.reason]])
    }

    // -------------------------------------------------------------- 4. full claim pack
    if (packsDir) {
      fs.mkdirSync(packsDir, { recursive: true })
      const pack = await buildClaimPack(uploads, new Uint8Array(fs.readFileSync(path.join(__dirname, '../public/hmrc_template.ods'))))
      fs.writeFileSync(path.join(packsDir, `${t.name}_claim_pack.zip`), pack.zip)
      console.log(`  pack written: ${t.name}_claim_pack.zip (${(pack.zip.byteLength / 1e6).toFixed(1)} MB)`)
    }
  }
  console.log(failures ? `\n${failures} problem(s) found` : '\nAll checks passed ✓')
  process.exit(failures ? 1 : 0)
}
main()
