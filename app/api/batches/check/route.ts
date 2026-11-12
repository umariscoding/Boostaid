import { NextResponse } from 'next/server'
import { getDb } from '@/lib/db'
import { UUID_RE, withAuth } from '@/lib/api'
import { findAlreadyClaimed } from '@/lib/duplicates'

/**
 * Before saving: which of these donations are already in a claim for this charity?
 * Body: { charityId, items: [{ platform, donationId }] }
 */
export const POST = withAuth(async (req) => {
  const body = await req.json().catch(() => null)
  const charityId = String(body?.charityId ?? '')
  if (!UUID_RE.test(charityId)) throw new RangeError('Choose which charity this upload is for')
  if (!Array.isArray(body?.items) || body.items.length > 200000) throw new RangeError('Invalid request')

  const items = (body.items as { platform?: unknown; donationId?: unknown }[]).map((i) => ({
    platform: String(i.platform ?? ''),
    donationId: String(i.donationId ?? ''),
  }))
  return NextResponse.json({ claimed: await findAlreadyClaimed(getDb(), charityId, items) })
})
