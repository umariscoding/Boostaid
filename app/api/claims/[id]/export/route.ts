import { NextResponse } from 'next/server'
import { eq } from 'drizzle-orm'
import { claims, getDb } from '@/lib/db'
import { UUID_RE, withAuth } from '@/lib/api'
import { buildAuditWorkbook, xlsxResponse } from '@/lib/export'

export const GET = withAuth<{ id: string }>(async (_req, { params }) => {
  const { id } = await params
  if (!UUID_RE.test(id)) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const [claim] = await getDb().select({ ref: claims.ref }).from(claims).where(eq(claims.id, id))
  if (!claim) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  return xlsxResponse(await buildAuditWorkbook(eq(claims.id, id)), `${claim.ref}_audit.xlsx`)
})
