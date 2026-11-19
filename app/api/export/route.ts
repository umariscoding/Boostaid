import { and, eq, inArray } from 'drizzle-orm'
import { charities, claims, getDb } from '@/lib/db'
import { withAuth } from '@/lib/api'
import { buildAuditWorkbook, xlsxResponse } from '@/lib/export'
import { claimsOfCharity, parseCharityParam } from '@/lib/scope'

const slug = (s: string) => s.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 40) || 'charity'

/** GET /api/export?view=all|active|archive&charity=<id|all> - audit workbook for every matching claim. */
export const GET = withAuth(async (req) => {
  const view = req.nextUrl.searchParams.get('view') ?? 'all'
  const charityId = parseCharityParam(req.nextUrl.searchParams.get('charity'))
  const status =
    view === 'archive' ? eq(claims.status, 'received') : view === 'active' ? inArray(claims.status, ['pending', 'submitted']) : undefined

  let label = 'all-charities'
  if (charityId) {
    const [c] = await getDb().select({ name: charities.name }).from(charities).where(eq(charities.id, charityId))
    label = slug(c?.name ?? 'charity')
  }

  const buffer = await buildAuditWorkbook(and(status, charityId ? claimsOfCharity(charityId) : undefined))
  const stamp = new Date().toISOString().slice(0, 10)
  return xlsxResponse(buffer, `boostaid_audit_${label}_${view}_${stamp}.xlsx`)
})
