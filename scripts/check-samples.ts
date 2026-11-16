/**
 * Runs BoostAid's Gift Aid engine on every sample file and checks each row against expected.json.
 *
 *   npm run samples:check   (reads samples/)
 *
 * Exit code 0 = every row did exactly what the answer sheet says.
 */
import fs from 'node:fs'
import path from 'node:path'
import { buildClaim, readTables } from '../lib/giftaid'

const dir = process.argv[2]
if (!dir) { console.error('usage: npx tsx check-samples.ts <sample folder>'); process.exit(2) }
const expected = JSON.parse(fs.readFileSync(path.join(dir, 'expected.json'), 'utf8')) as Record<string, string | number>[]

let failures = 0
const files = [...new Set(expected.map((e) => String(e.File)))]
for (const file of files) {
  const tables = readTables(file, new Uint8Array(fs.readFileSync(path.join(dir, file))))
  const res = buildClaim(tables)
  const main = tables[0].sheetName
  const where = (row: string | number) => `${file}:${String(row).includes('!') ? row : `${main}!${row}`}`

  type Got = { bucket: string; reason: string; row?: Record<string, unknown> }
  const got = new Map<string, Got>()
  for (const r of res.ready) got.set(r.source, { bucket: 'ready', reason: '', row: r as unknown as Record<string, unknown> })
  for (const r of res.overseas) got.set(r.source, { bucket: 'overseas', reason: '', row: r as unknown as Record<string, unknown> })
  for (const r of res.review) got.set(r.source, { bucket: 'review', reason: r.reason })
  for (const r of res.ineligible) got.set(r.source, { bucket: 'ineligible', reason: r.reason })

  let ok = 0
  const problems: string[] = []
  for (const e of expected.filter((x) => x.File === file)) {
    const g = got.get(where(e.Row)) ?? { bucket: 'skipped', reason: '' }
    const bad: string[] = []
    if (g.bucket !== e.Expected) bad.push(`went to ${g.bucket}${g.reason ? ` (${g.reason})` : ''}, expected ${e.Expected}`)
    else if (e.Expected !== 'skipped' && e['Reason contains'] && !g.reason.toLowerCase().includes(String(e['Reason contains']).toLowerCase())) {
      bad.push(`reason "${g.reason}" does not mention "${e['Reason contains']}"`)
    }
    if (!bad.length && g.row && e.Expected === 'ready') {
      const want: [string, unknown, unknown][] = [
        ['title', e.Title ?? '', g.row.title], ['first name', e['First name'], g.row.firstName], ['last name', e['Last name'], g.row.lastName],
        ['house', e.House, g.row.house], ['postcode', e.Postcode, g.row.postcode], ['date', e['Donation date'], g.row.date], ['amount', Number(e.Amount), g.row.amount],
      ]
      for (const [field, w, v] of want) if (String(w ?? '') !== String(v ?? '')) bad.push(`${field} is "${v}", expected "${w}"`)
    }
    if (bad.length) problems.push(`  ✗ row ${e.Row} [${e.Case}]: ${bad.join('; ')}`)
    else ok++
  }
  failures += problems.length
  const counts = `ready ${res.ready.length} · overseas ${res.overseas.length} · review ${res.review.length} · not eligible ${res.ineligible.length}`
  console.log(`${problems.length ? '✗' : '✓'} ${file} (${res.platforms.join(', ') || 'not recognised'}): ${ok}/${ok + problems.length} rows as expected - ${counts}`)
  for (const p of problems) console.log(p)
}
console.log(failures ? `\n${failures} row(s) did not match` : '\nAll sample rows matched ✓')
process.exit(failures ? 1 : 0)
