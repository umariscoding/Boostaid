import { NextResponse } from 'next/server'
import { desc, eq } from 'drizzle-orm'
import { gasdsClaims, getDb } from '@/lib/db'
import { withAuth } from '@/lib/api'
import { openYears } from '@/lib/gasds'
import { claimedByYear, giftAidFigures, loadReceipts, requireCharity, toClaimDto } from '@/lib/gasds-db'

/** Everything the Small donations page needs for one charity: receipts of the open tax years, saved claims, Gift Aid figures. */
export const GET = withAuth(async (req) => {
  const charityId = await requireCharity(req.nextUrl.searchParams.get('charity'))
  const today = new Date().toISOString().slice(0, 10)
  const years = openYears(today)
  const [receipts, claimRows, giftAid] = await Promise.all([
    loadReceipts(charityId, years),
    getDb().select().from(gasdsClaims).where(eq(gasdsClaims.charityId, charityId)).orderBy(desc(gasdsClaims.sentAt)),
    giftAidFigures(charityId, years),
  ])
  const claims = claimRows.map(toClaimDto)
  return NextResponse.json({ today, years, receipts, claims, claimedByYear: claimedByYear(claims), giftAid })
})
