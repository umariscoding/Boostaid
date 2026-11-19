import { NextResponse } from 'next/server'
import { getDb, gasdsYears } from '@/lib/db'
import { withAuth } from '@/lib/api'
import { openYears } from '@/lib/gasds'
import { requireCharity } from '@/lib/gasds-db'

/** Sets "Gift Aid claimed outside BoostAid" for one charity and tax year. */
export const PUT = withAuth(async (req) => {
  const body = await req.json().catch(() => null)
  if (!body) throw new RangeError('Invalid body')
  const charityId = await requireCharity(body.charityId)
  const year = Number(body.taxYear)
  if (!openYears(new Date().toISOString().slice(0, 10)).includes(year)) throw new RangeError('That tax year is closed')
  const amount = Number(body.giftAidElsewhere || 0)
  if (!Number.isFinite(amount) || amount < 0 || amount > 1e10) throw new RangeError('Enter an amount of 0 or more')
  const value = (Math.round(amount * 100) / 100).toFixed(2)
  await getDb()
    .insert(gasdsYears)
    .values({ charityId, taxYear: year, giftAidElsewhere: value })
    .onConflictDoUpdate({ target: [gasdsYears.charityId, gasdsYears.taxYear], set: { giftAidElsewhere: value } })
  return NextResponse.json({ success: true })
})
