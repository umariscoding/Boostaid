import { NextResponse } from 'next/server'
import { UUID_RE, withAuth, requireAdmin } from '@/lib/api'
import { previewInvoice, sendInvoice } from '@/lib/invoices'

/** { claimId, preview? }: the preview shows the invoice; otherwise it is emailed to the charity. Only for a received claim. */
export const POST = withAuth(async (req) => {
  await requireAdmin(req)
  const body = await req.json().catch(() => null)
  if (!body || typeof body.claimId !== 'string' || !UUID_RE.test(body.claimId)) throw new RangeError('Invalid body')
  if (body.preview) return NextResponse.json(await previewInvoice(body.claimId))
  const inv = await sendInvoice(body.claimId)
  return NextResponse.json({ invoice: { id: inv.id, number: inv.number, delivered: inv.delivered } })
})
