import { NextResponse } from 'next/server'
import { UUID_RE, withAuth, requireAdmin } from '@/lib/api'
import { invoicePdfFor } from '@/lib/invoices'

export const GET = withAuth<{ id: string }>(async (req, { params }) => {
  await requireAdmin(req)
  const { id } = await params
  if (!UUID_RE.test(id)) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const f = await invoicePdfFor(id)
  if (!f) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  return new Response(new Uint8Array(f.pdf), { headers: { 'Content-Type': 'application/pdf', 'Content-Disposition': `inline; filename="${f.name}"`, 'Cache-Control': 'no-store' } })
})
