import { NextResponse } from 'next/server'
import { sql } from 'drizzle-orm'
import { charities, getDb } from '@/lib/db'
import { withAuth } from '@/lib/api'

/** All charities with how many records each one holds (shown in the delete warning). */
export const GET = withAuth(async () => {
  const rows = await getDb().execute<{
    id: string
    name: string
    created_at: Date
    claims: number
    donations: number
    donated: string
  }>(sql`
    select c.id, c.name, c.created_at,
      (select count(*)::int from claims cl join batches b on b.id = cl.batch_id where b.charity_id = c.id) as claims,
      (select count(*)::int from donations d join claims cl on cl.id = d.claim_id join batches b on b.id = cl.batch_id
        where b.charity_id = c.id) as donations,
      (select coalesce(sum(cl.total_amount), 0)::text from claims cl join batches b on b.id = cl.batch_id
        where b.charity_id = c.id) as donated
    from charities c
    order by lower(c.name)
  `)

  return NextResponse.json({
    charities: rows.rows.map((r) => ({
      id: r.id,
      name: r.name,
      createdAt: new Date(r.created_at).toISOString(),
      claimCount: r.claims,
      donationCount: r.donations,
      donated: Number(r.donated),
    })),
  })
})

export const POST = withAuth(async (req) => {
  const body = await req.json().catch(() => null)
  const name = String(body?.name ?? '').replace(/\s+/g, ' ').trim()
  if (!name) throw new RangeError('Enter the charity name')
  if (name.length > 100) throw new RangeError('Charity names can be up to 100 characters')

  const db = getDb()
  const [existing] = await db
    .select({ id: charities.id })
    .from(charities)
    .where(sql`lower(${charities.name}) = lower(${name})`)
  if (existing) return NextResponse.json({ error: `"${name}" already exists` }, { status: 409 })

  const [created] = await db.insert(charities).values({ name }).returning()
  return NextResponse.json(
    { charity: { id: created.id, name: created.name, createdAt: created.createdAt.toISOString(), claimCount: 0, donationCount: 0, donated: 0 } },
    { status: 201 }
  )
})
