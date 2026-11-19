import { NextResponse } from 'next/server'
import { withAuth } from '@/lib/api'
import { appBase } from '@/lib/email'
import { previewRecoveryEmail, sendRecoveryEmails } from '@/lib/recovery-db'
import { requireCharity } from '@/lib/gasds-db'
import { UUID_RE } from '@/lib/api'

/** { charityId, rowIds, preview? }: preview returns the email as it would be sent; otherwise one email per donor goes out. */
export const POST = withAuth(async (req) => {
  const body = await req.json().catch(() => null)
  if (!body || !Array.isArray(body.rowIds)) throw new RangeError('Invalid body')
  const charityId = await requireCharity(body.charityId)
  const ids = body.rowIds.filter((x: unknown): x is string => typeof x === 'string' && UUID_RE.test(x)).slice(0, 5000)
  const base = appBase(req.nextUrl.origin)
  if (body.preview) return NextResponse.json(await previewRecoveryEmail(charityId, ids, base))
  return NextResponse.json(await sendRecoveryEmails(charityId, ids, base))
})
