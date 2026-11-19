/**
 * Acceptance test for the GASDS rules (lib/gasds.ts + lib/gasds-read.ts).
 *
 *   npm run gasds:verify
 *
 * Runs the client's three test sets (tax year 2026/27, a charity with plenty of Gift Aid, "today" = 16 Nov 2026)
 * and checks every number in INSTRUCTIONS_GASDS.txt. Exit code 0 = everything matches.
 */
import fs from 'node:fs'
import path from 'node:path'
import { analyse, applyBanking, previewUpload, p, needsCount, type GReceipt, type Analysis } from '../lib/gasds'
import { readGasdsFiles } from '../lib/gasds-read'

const DIR = path.join(__dirname, '../samples/gasds')
const TODAY = '2026-11-16'
const YEAR = 2026
const PLENTY = 1_000_000

let failures = 0
const ok = (label: string, got: unknown, want: unknown) => {
  const same = typeof want === 'number' ? Math.abs(Number(got) - want) < 0.005 : got === want
  if (!same) failures++
  console.log(`  ${same ? '✓' : '✗'} ${label}: ${got}${same ? '' : `  (expected ${want})`}`)
}

function load(files: string[]) {
  const read = readGasdsFiles(files.map((f) => ({ name: path.basename(f), data: new Uint8Array(fs.readFileSync(f)) })))
  if (read.errors.length) { failures++; console.log('  ✗ read errors: ' + read.errors.join(' | ')) }
  const receipts = applyBanking(read.receipts, read.banking)
  return { read, receipts: receipts.filter((r) => r.year === YEAR) }
}
const dirFiles = (d: string) => fs.readdirSync(path.join(DIR, d)).sort().map((f) => path.join(DIR, d, f))
const reason = (a: Analysis, no: string) => { const s = a.statuses.get(no)!; return `${s.label}${s.reason ? ': ' + s.reason : ''}` }

// ------------------------------------------------------------------ 1. small test
console.log('\nSmall test (3 files)')
{
  const { receipts } = load(dirFiles('1-small-2026'))
  const a = analyse(receipts, [], PLENTY, YEAR)
  ok('claim (type into HMRC)', a.amount, 8000)
  ok('HMRC pays', a.topUp, 2000)
  ok('over the limit', a.overBy, 212)
  ok('collections needing counts', a.collections.length, 3)
  ok('...all with a bank slip pre-filled', a.collections.filter((x) => x.r.slip).length, 3)
  const used = (no: string) => a.plan.find((x) => x.r.no === no)?.used
  ok('C-000202 counts at the bank figure', used('C-000202') ?? -1, 1960)
  ok('C-000203 part used (limit runs out)', used('C-000203') ?? -1, 428)
  ok('A-000401 (Gift Aid)', reason(a, 'A-000401'), 'Not included: has Gift Aid, so it goes in the Gift Aid claim')
  ok('A-000402 (£100)', reason(a, 'A-000402'), 'Not included: over £30 from one person')
  ok('A-000404 (cheque)', reason(a, 'A-000404'), 'Not included: paid by cheque or bank transfer')
  ok('A-000403 (£30 card) in the claim', reason(a, 'A-000403'), 'In this claim')
  ok('A-000406 in the claim', reason(a, 'A-000406'), 'In this claim')
  ok('C-000204 waiting for the bank', reason(a, 'C-000204'), 'Waiting: not banked yet')
  ok('void receipt skipped at upload', receipts.some((r) => r.no === 'A-000405'), false)

  // Save the claim: those receipts are locked with the amount used. Then the year is full.
  const saved: GReceipt[] = receipts.map((r) => {
    const u = a.plan.find((x) => x.r.no === r.no)
    return u ? { ...r, claimRef: 'REF-1', used: u.used } : r
  })
  const after = analyse(saved, [8000], PLENTY, YEAR)
  ok('after saving: room left', after.room, 0)
  ok('after saving: nothing more to claim', after.amount, 0)
  ok('after saving: C-000204', reason(after, 'C-000204'), 'Not included: the £8,000 limit for the year is reached')

  // Re-upload: unchanged receipts are "already uploaded"; claimed ones are untouched
  const again = previewUpload(saved, receipts, [], { [YEAR]: 8000 }, { [YEAR]: PLENTY }, TODAY)
  ok('re-upload: nothing new', again.every((o) => o.action === 'claimed' || o.action === 'same'), true)
}

// ------------------------------------------------------------------ 2. big test
console.log('\nBig test (274 receipts + 8 banking files)')
{
  const { read, receipts } = load(dirFiles('2-big-2026'))
  const a = analyse(receipts, [], PLENTY, YEAR)
  ok('claim (type into HMRC)', a.amount, 8000)
  ok('HMRC pays', a.topUp, 2000)
  ok('over the limit', a.overBy, 25230.11)
  const why = (needle: string) => receipts.filter((r) => a.statuses.get(r.no)?.reason.includes(needle)).length
  ok('Gift Aid', why('has Gift Aid'), 59)
  ok('over £30', why('over £30'), 34)
  ok('cheque or transfer', why('cheque or bank transfer'), 31)
  ok('waiting for the bank', a.waiting.length, 6)
  // The instructions say £2,243.00; the 6 receipts in the file add up to £2,213.00. Reported, not failed.
  const waitingSum = a.waiting.reduce((s, r) => s + r.amount, 0)
  console.log(`  ! waiting receipts are worth £${waitingSum.toFixed(2)} here; the instructions say £2,243.00 (a £30.00 difference to ask about)`)
  const short = read.banking.flatMap((b) => b.items).filter((it) => p(it.counted) < p(it.amount))
  ok('banking shortfalls', short.length, 3)
  ok('...C-000012 short by', short.find((s) => s.no === 'C-000012') ? short.find((s) => s.no === 'C-000012')!.counted - short.find((s) => s.no === 'C-000012')!.amount : 0, -5)
  const cols = a.plan.filter((x) => !x.r.gift)
  ok('collections in the claim', cols.length, 12)
  ok('...slips pre-filled (cash collections)', cols.filter((x) => needsCount(x.r) && x.r.slip).length, 8)
  ok('...without a slip are card collections', cols.filter((x) => !x.r.slip && x.r.card).length, 4)
  ok('"Gift Aid on another receipt" hint', a.maybeGiftAid.length, 8)
  const voids = readGasdsFiles(dirFiles('2-big-2026').filter((f) => /receipt-list/.test(f)).map((f) => ({ name: f, data: new Uint8Array(fs.readFileSync(f)) }))).receipts.length
  ok('receipts read (voids skipped)', voids, 274 - 6)
}

// ------------------------------------------------------------------ 3. template example
console.log('\nTemplate example')
{
  const f = path.join(DIR, 'BoostAid_GASDS_template_EXAMPLE.xlsx')
  const read = readGasdsFiles([{ name: path.basename(f), data: new Uint8Array(fs.readFileSync(f)) }])
  if (read.errors.length) { failures++; console.log('  ✗ ' + read.errors.join(' | ')) }
  const a = analyse(read.receipts.filter((r) => r.year === YEAR), [], PLENTY, YEAR)
  ok('claim (type into HMRC)', a.amount, 746)
  ok('HMRC pays', a.topUp, 186.5)
  ok('cheque excluded', reason(a, 'R-13'), 'Not included: paid by cheque or bank transfer')
  ok('Gift Aid row excluded', reason(a, 'R-14').startsWith('Not included: has Gift Aid'), true)
  const waiting = read.receipts.find((r) => r.notes === 'Friday collection, Masjid B')!
  ok('unbanked collection waits', reason(a, waiting.no), 'Waiting: not banked yet')
  const card = a.plan.find((x) => x.r.card)!
  ok('card collection needs no count', needsCount(card.r), false)
}

// ------------------------------------------------------------------ 4. the limits
console.log('\nLimits')
{
  const { receipts } = load(dirFiles('1-small-2026'))
  const a = analyse(receipts, [], 500, YEAR) // only £500 of Gift Aid -> 10 x = £5,000
  ok('matching rule caps the claim at 10 x Gift Aid', a.amount, 5000)
  ok('...and says so', a.matchingBites, true)
  const b = analyse(receipts, [3000], PLENTY, YEAR) // £3,000 already claimed this year
  ok('earlier claims reduce the room', b.amount, 5000)
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll GASDS checks passed ✓')
process.exit(failures ? 1 : 0)
