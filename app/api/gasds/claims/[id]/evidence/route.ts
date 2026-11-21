import { NextResponse } from 'next/server'
import { asc, eq, inArray } from 'drizzle-orm'
import { charities, gasdsClaims, gasdsPhotos, gasdsReceipts, getDb } from '@/lib/db'
import { UUID_RE, withAuth } from '@/lib/api'
import { buildEvidenceZip } from '@/lib/gasds-evidence'
import { toClaimDto, toReceipt } from '@/lib/gasds-db'

/** The evidence download: the summary printout (PDF) and the table of receipts (Excel), plus any count-sheet photos. */
export const GET = withAuth<{ id: string }>(async (_req, { params }) => {
  const { id } = await params
  if (!UUID_RE.test(id)) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const db = getDb()
  const [row] = await db
    .select({ claim: gasdsClaims, charity: charities.name })
    .from(gasdsClaims)
    .innerJoin(charities, eq(charities.id, gasdsClaims.charityId))
    .where(eq(gasdsClaims.id, id))
  if (!row) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const rows = await db.select().from(gasdsReceipts).where(eq(gasdsReceipts.claimId, id)).orderBy(asc(gasdsReceipts.receiptDate), asc(gasdsReceipts.receiptNo))
  const receipts = rows.map((r) => toReceipt(r, row.claim.hmrcRef))
  const withPhoto = rows.filter((r) => r.photoName)
  const photoRows = withPhoto.length ? await db.select().from(gasdsPhotos).where(inArray(gasdsPhotos.receiptId, withPhoto.map((r) => r.id))) : []
  const photos = photoRows.map((p) => ({ name: withPhoto.find((r) => r.id === p.receiptId)!.receiptNo, mime: p.mime, data: new Uint8Array(p.data) }))
  const claim = toClaimDto(row.claim)
  const zip = await buildEvidenceZip(claim, row.charity, receipts, photos)
  const name = `GASDS_evidence_${row.claim.hmrcRef.replace(/[^\w-]/g, '_')}.zip`
  return new Response(new Uint8Array(zip), { headers: { 'Content-Type': 'application/zip', 'Content-Disposition': `attachment; filename="${name}"`, 'Cache-Control': 'no-store' } })
})
