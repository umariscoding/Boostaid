import { NextResponse } from 'next/server'
import { and, eq, sql } from 'drizzle-orm'
import { gasdsReceipts, getDb } from '@/lib/db'
import { withAuth } from '@/lib/api'
import { claimedByYear, giftAidFigures, loadReceipts, requireCharity } from '@/lib/gasds-db'
import { openYears, previewUpload, taxYearOf, type BankingFile, type GReceipt } from '@/lib/gasds'
import { gasdsClaims } from '@/lib/db'

const DATE_RE = /^\d{4}-\d{2}-\d{2}$/
const str = (v: unknown, max: number) => (typeof v === 'string' ? v.trim().slice(0, max) : '')
const num = (v: unknown) => (typeof v === 'number' && Number.isFinite(v) ? v : NaN)

/** Re-checks what the browser sent: the server never trusts the tax year or amounts blindly. */
function cleanReceipt(x: unknown): GReceipt {
  const r = (x ?? {}) as Record<string, unknown>
  const date = str(r.date, 10)
  const amount = num(r.amount)
  const no = str(r.no, 120)
  if (!no || !DATE_RE.test(date) || !(amount > 0) || amount > 1e8) throw new RangeError('One of the receipts has a missing number, date or amount')
  return {
    no,
    slip: str(r.slip, 60) || undefined,
    date,
    year: taxYearOf(date),
    amount: Math.round(amount * 100) / 100,
    card: r.card === true,
    notCash: r.notCash === true,
    notes: str(r.notes, 300),
    donor: str(r.donor, 200),
    gift: r.gift === true,
    giftAid: r.giftAid === true,
    banked: r.banked === true || !!str(r.slip, 60),
    isVoid: false,
  }
}

function cleanBanking(x: unknown): BankingFile[] {
  if (!Array.isArray(x)) return []
  return x.slice(0, 200).map((b) => {
    const slip = str(b?.slip, 60)
    if (!slip || !Array.isArray(b?.items)) throw new RangeError('A banking file is missing its slip number')
    return {
      slip,
      items: b.items.slice(0, 5000).map((it: Record<string, unknown>) => ({
        no: str(it.no, 120),
        amount: Math.round(num(it.amount) * 100) / 100,
        counted: Math.round((Number.isFinite(num(it.counted)) ? num(it.counted) : num(it.amount)) * 100) / 100,
      })),
    }
  })
}

/**
 * Saves an upload. Receipts are matched by receipt number: unchanged ones are left alone, changed ones (e.g. now banked)
 * are updated, receipts already in a saved claim are never touched. Banking files mark their receipts banked.
 */
export const POST = withAuth(async (req) => {
  const body = await req.json().catch(() => null)
  if (!body || !Array.isArray(body.receipts)) throw new RangeError('Invalid body')
  const charityId = await requireCharity(body.charityId)
  if (body.receipts.length > 20000) throw new RangeError('Too many receipts in one upload. Please upload them in parts.')
  const uploaded = body.receipts.map(cleanReceipt)
  const banking = cleanBanking(body.banking)
  const today = new Date().toISOString().slice(0, 10)
  const years = openYears(today)

  const db = getDb()
  const result = await db.transaction(async (tx) => {
    // One upload at a time per charity, so two people can't overfill the same year
    await tx.execute(sql`select pg_advisory_xact_lock(hashtext(${'gasds:' + charityId}))`)
    const stored = await loadReceipts(charityId, years)
    const claimRows = await tx.select().from(gasdsClaims).where(eq(gasdsClaims.charityId, charityId))
    const claimed = claimedByYear(claimRows.map((c) => ({ taxYear: c.taxYear, amount: Number(c.amount) })))
    const ga = await giftAidFigures(charityId, years)
    const gaTotals = Object.fromEntries(years.map((y) => [y, ga[y].boostaid + ga[y].elsewhere]))

    const outcomes = previewUpload(stored, uploaded, banking, claimed, gaTotals, today)
    const have = new Map(stored.map((r) => [r.no, r]))
    let added = 0
    let updated = 0

    const fresh = outcomes.filter((o) => o.action === 'new').map((o) => o.receipt)
    for (let i = 0; i < fresh.length; i += 500) {
      await tx.insert(gasdsReceipts).values(
        fresh.slice(i, i + 500).map((r) => ({
          charityId,
          receiptNo: r.no,
          receiptDate: r.date,
          taxYear: r.year,
          amount: r.amount.toFixed(2),
          card: r.card,
          notCash: r.notCash,
          notes: r.notes,
          donor: r.donor,
          isGift: r.gift,
          giftAid: r.giftAid,
          banked: r.banked,
          slip: r.slip ?? null,
          counted: r.counted == null ? null : r.counted.toFixed(2),
        }))
      ).onConflictDoNothing()
    }
    added = fresh.length

    const touched = new Set<string>()
    for (const o of outcomes.filter((x) => x.action === 'update')) {
      const prev = have.get(o.receipt.no)
      if (!prev?.id) continue
      const r = o.receipt
      await tx
        .update(gasdsReceipts)
        .set({
          amount: r.amount.toFixed(2),
          giftAid: r.giftAid,
          banked: r.banked,
          card: r.card,
          notCash: r.notCash,
          slip: r.slip ?? prev.slip ?? null,
          counted: r.counted == null ? (prev.counted == null ? null : prev.counted.toFixed(2)) : r.counted.toFixed(2),
        })
        .where(and(eq(gasdsReceipts.id, prev.id), sql`${gasdsReceipts.claimId} is null`))
      touched.add(r.no)
      updated++
    }

    // Banking files can also mark receipts that were uploaded earlier
    for (const b of banking) {
      const nos = b.items.map((it) => it.no).filter((n) => have.has(n) && !touched.has(n) && !have.get(n)!.claimRef)
      for (const it of b.items.filter((i) => nos.includes(i.no))) {
        const prev = have.get(it.no)!
        if (prev.banked && prev.slip === b.slip && (prev.counted ?? -1) === it.counted) continue
        await tx
          .update(gasdsReceipts)
          .set({ banked: true, slip: b.slip, counted: it.counted.toFixed(2) })
          .where(and(eq(gasdsReceipts.id, prev.id!), sql`${gasdsReceipts.claimId} is null`))
        updated++
      }
    }
    return { added, updated, saved: outcomes.length, tooOld: outcomes.filter((o) => o.action === 'too-old').length }
  })
  return NextResponse.json(result)
})

