import { NextResponse } from 'next/server'
import { and, eq, sql } from 'drizzle-orm'
import { gasdsClaims, gasdsReceipts, getDb, users } from '@/lib/db'
import { currentUserId, withAuth } from '@/lib/api'
import { analyse, deadlineOf, openYears, p } from '@/lib/gasds'
import { claimedByYear, giftAidFigures, loadReceipts, requireCharity, toClaimDto } from '@/lib/gasds-db'

/**
 * Step 3: save the claim. The server works the amount out again from what is stored (the page can be out of date),
 * needs the HMRC reference, then locks the receipts in the claim so they can never be claimed again.
 */
export const POST = withAuth(async (req) => {
  const body = await req.json().catch(() => null)
  if (!body) throw new RangeError('Invalid body')
  const charityId = await requireCharity(body.charityId)
  const year = Number(body.taxYear)
  const today = new Date().toISOString().slice(0, 10)
  if (!openYears(today).includes(year) || deadlineOf(year) < today) throw new RangeError('That tax year is closed, so it can’t be claimed any more')
  const ref = typeof body.hmrcRef === 'string' ? body.hmrcRef.trim().slice(0, 100) : ''
  if (!ref) throw new RangeError('Enter the reference HMRC gave you before saving the claim')

  const db = getDb()
  const uid = currentUserId(req)
  const [user] = uid ? await db.select().from(users).where(eq(users.id, uid)) : []
  const sentBy = user ? user.name || user.email : ''

  const saved = await db.transaction(async (tx) => {
    await tx.execute(sql`select pg_advisory_xact_lock(hashtext(${'gasds:' + charityId}))`)
    const receipts = (await loadReceipts(charityId, [year])).filter((r) => r.year === year)
    const claimRows = await tx.select().from(gasdsClaims).where(and(eq(gasdsClaims.charityId, charityId), eq(gasdsClaims.taxYear, year)))
    const ga = (await giftAidFigures(charityId, [year]))[year]
    const giftAid = Math.round((ga.boostaid + ga.elsewhere) * 100) / 100
    const a = analyse(receipts, claimRows.map((c) => Number(c.amount)), giftAid, year)

    if (!(a.amount > 0)) throw new RangeError('There is nothing to claim for this tax year')
    if (p(a.amount) !== p(Number(body.amount))) {
      throw new RangeError(`The amount has changed to £${a.amount.toFixed(2)} since you opened this page. Please check it and try again.`)
    }

    const [claim] = await tx
      .insert(gasdsClaims)
      .values({ charityId, taxYear: year, amount: a.amount.toFixed(2), hmrcRef: ref, sentBy, giftAidFigure: giftAid.toFixed(2), status: 'submitted' })
      .returning()
    for (const { r, used } of a.plan) {
      await tx.update(gasdsReceipts).set({ claimId: claim.id, used: used.toFixed(2) }).where(and(eq(gasdsReceipts.id, r.id!), sql`${gasdsReceipts.claimId} is null`))
    }
    return claim
  })
  const all = await db.select().from(gasdsClaims).where(eq(gasdsClaims.charityId, charityId))
  return NextResponse.json({ claim: toClaimDto(saved), claimedByYear: claimedByYear(all.map(toClaimDto)) })
})
