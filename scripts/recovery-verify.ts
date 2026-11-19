/**
 * Acceptance test for Gift Aid recovery and the fee (lib/recovery.ts, lib/fees.ts), using the client's two LaunchGood files.
 *
 *   npm run recovery:verify
 *
 * Exit code 0 = everything matches INSTRUCTIONS_Gift_Aid_recovery.txt.
 */
import fs from 'node:fs'
import path from 'node:path'
import { buildClaim, readTables } from '../lib/giftaid'
import { checkAnswers, groupByDonor, heldBackRows, reconcileUpload, rowKey } from '../lib/recovery'
import { hashCode, newCode } from '../lib/recovery-codes'
import { feeFigures } from '../lib/fees'
import { giftAidOn } from '../lib/claims'

const DIR = path.join(__dirname, '../samples/recovery')
let failures = 0
const ok = (label: string, got: unknown, want: unknown) => {
  const same = typeof want === 'number' ? Math.abs(Number(got) - want) < 0.005 : got === want
  if (!same) failures++
  console.log(`  ${same ? '✓' : '✗'} ${label}: ${got}${same ? '' : `  (expected ${want})`}`)
}
const load = (f: string) => {
  const tables = readTables(f, new Uint8Array(fs.readFileSync(path.join(DIR, f))))
  return { tables, res: buildClaim(tables) }
}

console.log('\nFirst export (30 donations)')
const first = load('1_LaunchGood_first_export.xlsx')
const held1 = heldBackRows(first.res.review, first.tables)
ok('ready', first.res.ready.length + first.res.overseas.length, 18)
ok('needs review', held1.length, 12)
const emailable = held1.filter((r) => r.email)
ok('review rows with an email', emailable.length, 10)
ok('one email per donor (Yusuf Patel has 2 rows)', groupByDonor(emailable).size, 9)
ok('without an email (fix by hand)', held1.filter((r) => !r.email).map((r) => `${r.firstName} ${r.lastName}`).sort().join(' / '), 'Anonymous Donor / Ruth Evans')
ok('emailable donations total', emailable.reduce((s, r) => s + r.amount, 0), 1230)
ok('Gift Aid that can be recovered', giftAidOn(emailable.reduce((s, r) => s + r.amount, 0)), 307.5)

console.log('\nThe donor\'s form (answers go through the engine)')
const good = checkAnswers({ title: 'Ms', first: 'Sumaya', last: 'Mirza', house: '93', postcode: 'bd7 1hd' }, '2026-05-14', 50)
ok('a good answer passes', good.ok, true)
ok('...postcode is tidied', good.ok ? good.cleaned.postcode : '', 'BD7 1HD')
ok('one person only ("Ali & Sara")', checkAnswers({ title: '', first: 'Ali & Sara', last: 'Khan', house: '5', postcode: 'B11 2QT' }, '2026-09-19', 200).ok, false)
ok('a partial postcode is refused', checkAnswers({ title: '', first: 'Imran', last: 'Sheikh', house: '40', postcode: 'E1 6A' }, '2026-07-11', 40).ok, false)
ok('a missing house number is refused', checkAnswers({ title: '', first: 'Ruth', last: 'Evans', house: '', postcode: 'CF10 1EP' }, '2026-10-01', 30).ok, false)
ok('a placeholder name is refused', checkAnswers({ title: '', first: 'Anonymous', last: 'Donor', house: '101', postcode: 'B16 5GU' }, '2026-10-22', 150).ok, false)

console.log('\nLink codes')
const c = newCode()
ok('code is 10 characters', c.length, 10)
ok('hash is stable and not the code', hashCode(c) === hashCode(c.toLowerCase()) && hashCode(c) !== c, true)
ok('codes are random', new Set(Array.from({ length: 200 }, () => newCode())).size, 200)

console.log('\nRe-export a month later (the table in section 5)')
const second = load('2_LaunchGood_re-export_a_month_later.xlsx')
const held2 = heldBackRows(second.res.review, second.tables)
// Everything held back from the first export is stored; the emailed ones are "sent"
const existing = held1.map((r) => ({ key: rowKey(r.platform, r.donationId), status: r.email ? 'sent' : 'none' }))
const idOf = (id: string) => rowKey('LaunchGood', id)
const readyKeys = [...second.res.ready, ...second.res.overseas].map((r) => rowKey(r.platform, String(r.donationId)))
const rec = reconcileUpload(existing, held2.map((r) => rowKey(r.platform, r.donationId)), readyKeys)
ok('re-export: ready', second.res.ready.length + second.res.overseas.length, 24)
ok('re-export: needs review', held2.length, 12)
ok('new needing review (Layla Mirza)', rec.add.length, 1)
ok('...it is Layla, with her email', held2.find((r) => rec.add.includes(rowKey(r.platform, r.donationId)))?.email, 'test10@example.com')
ok('Adam Begum (93000108) fixed from the file', rec.fixedFromFile.includes(idOf('93000108')), true)
ok('Noor Aslam (93000109) still missing: no change', !rec.fixedFromFile.includes(idOf('93000109')) && !rec.add.includes(idOf('93000109')), true)
ok('only Adam is fixed from the file', rec.fixedFromFile.filter((k) => existing.some((e) => e.key === k)).length, 1)

console.log('\nFee (5% of recovered Gift Aid, once received)')
const exp = feeFigures(178.75, 4989, 4980)
ok('recovered part HMRC paid', exp.recoveredPaid, 178.43)
ok('fee', exp.fee, 8.92)
ok('paid in full: fee on all of it', feeFigures(178.75, 4989, 4989).fee, 8.94)
ok('paid MORE: never above what was expected', feeFigures(178.75, 4989, 5200).fee, 8.94)
ok('not received yet: fee shown on the full amount', feeFigures(178.75, 4989, null).fee, 8.94)
ok('without BoostAid you would have had', exp.withoutBoostAid ?? 0, 4801.57)

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll recovery checks passed ✓')
process.exit(failures ? 1 : 0)
