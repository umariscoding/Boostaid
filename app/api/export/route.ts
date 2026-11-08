import { eq, inArray } from 'drizzle-orm'
import { claims } from '@/lib/db'
import { withAuth } from '@/lib/api'
import { buildAuditWorkbook, xlsxResponse } from '@/lib/export'

/** GET /api/export?view=all|active|archive - audit workbook for every matching claim. */
export const GET = withAuth(async (req) => {
  const view = req.nextUrl.searchParams.get('view') ?? 'all'
  const where =
    view === 'archive' ? eq(claims.status, 'received') : view === 'active' ? inArray(claims.status, ['pending', 'submitted']) : undefined
  const buffer = await buildAuditWorkbook(where)
  const stamp = new Date().toISOString().slice(0, 10)
  return xlsxResponse(buffer, `boostaid_audit_${view}_${stamp}.xlsx`)
})
