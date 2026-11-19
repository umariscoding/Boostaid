import { NextResponse } from 'next/server'
import { eq } from 'drizzle-orm'
import { gasdsPhotos, gasdsReceipts, getDb } from '@/lib/db'
import { UUID_RE, withAuth } from '@/lib/api'

const MAX_BYTES = 3 * 1024 * 1024
const TYPES = ['image/jpeg', 'image/png', 'image/webp', 'application/pdf']

/** A photo (or PDF) of the count sheet. The body is the file itself; its name is in the x-file-name header. */
export const POST = withAuth<{ id: string }>(async (req, { params }) => {
  const { id } = await params
  if (!UUID_RE.test(id)) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const mime = (req.headers.get('content-type') ?? '').split(';')[0].trim().toLowerCase()
  if (!TYPES.includes(mime)) throw new RangeError('Attach a photo (JPG, PNG or WebP) or a PDF')
  const data = Buffer.from(await req.arrayBuffer())
  if (!data.length) throw new RangeError('That file is empty')
  if (data.length > MAX_BYTES) throw new RangeError('That file is over 3 MB. Please attach a smaller photo.')
  const name = decodeURIComponent(req.headers.get('x-file-name') ?? 'count-sheet').replace(/[^\w .()-]/g, '_').slice(0, 120)

  const db = getDb()
  const [receipt] = await db.select().from(gasdsReceipts).where(eq(gasdsReceipts.id, id))
  if (!receipt) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  if (receipt.claimId) throw new RangeError('This receipt is in a claim sent to HMRC, so it is locked')
  await db.transaction(async (tx) => {
    await tx.insert(gasdsPhotos).values({ receiptId: id, mime, data }).onConflictDoUpdate({ target: gasdsPhotos.receiptId, set: { mime, data } })
    await tx.update(gasdsReceipts).set({ photoName: name }).where(eq(gasdsReceipts.id, id))
  })
  return NextResponse.json({ photo: name })
})

export const GET = withAuth<{ id: string }>(async (_req, { params }) => {
  const { id } = await params
  if (!UUID_RE.test(id)) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const [p] = await getDb().select().from(gasdsPhotos).where(eq(gasdsPhotos.receiptId, id))
  if (!p) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  return new Response(new Uint8Array(p.data), {
    headers: { 'Content-Type': p.mime, 'Cache-Control': 'private, no-store', 'X-Content-Type-Options': 'nosniff', 'Content-Disposition': 'inline' },
  })
})
