import { NextResponse } from 'next/server'
import { UUID_RE, withAuth } from '@/lib/api'
import { fixByHand } from '@/lib/recovery-db'
import { requireCharity } from '@/lib/gasds-db'

/** The charity fixes a held-back row itself. It goes through the same engine checks as a donor's answer. No fee. */
export const POST = withAuth(async (req) => {
  const body = await req.json().catch(() => null)
  if (!body || typeof body.rowId !== 'string' || !UUID_RE.test(body.rowId)) throw new RangeError('Invalid body')
  const charityId = await requireCharity(body.charityId)
  const s = (v: unknown) => (typeof v === 'string' ? v.slice(0, 200) : '')
  await fixByHand(charityId, body.rowId, { title: s(body.title), first: s(body.first), last: s(body.last), house: s(body.house), postcode: s(body.postcode) })
  return NextResponse.json({ success: true })
})
