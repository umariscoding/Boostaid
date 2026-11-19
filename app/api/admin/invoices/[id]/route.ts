import { NextResponse } from 'next/server'
import { UUID_RE, withAuth, requireAdmin } from '@/lib/api'
import { markPaid, resendInvoice } from '@/lib/invoices'

/** { action: 'paid' | 'resend' } */
export const PATCH = withAuth<{ id: string }>(async (req, { params }) => {
  await requireAdmin(req)
  const { id } = await params
  if (!UUID_RE.test(id)) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const body = await req.json().catch(() => null)
  if (body?.action === 'paid') {
    await markPaid(id)
    return NextResponse.json({ success: true })
  }
  if (body?.action === 'resend') return NextResponse.json({ delivered: await resendInvoice(id) })
  throw new RangeError('Unknown action')
})
