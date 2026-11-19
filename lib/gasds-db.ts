import { and, eq, inArray, sql } from 'drizzle-orm'
import { charities, gasdsClaims, gasdsReceipts, gasdsYears, getDb } from '@/lib/db'
import { UUID_RE } from '@/lib/api'
import { taxYearRange, type Count, type GReceipt } from '@/lib/gasds'

type Row = typeof gasdsReceipts.$inferSelect
type ClaimRow = typeof gasdsClaims.$inferSelect

export function toReceipt(r: Row, claimRef?: string | null): GReceipt {
  return {
    id: r.id,
    no: r.receiptNo,
    slip: r.slip,
    date: r.receiptDate,
    year: r.taxYear,
    amount: Number(r.amount),
    card: r.card,
    notCash: r.notCash,
    notes: r.notes,
    donor: r.donor,
    gift: r.isGift,
    giftAid: r.giftAid,
    banked: r.banked,
    isVoid: false,
    counted: r.counted == null ? null : Number(r.counted),
    count: (r.count as Count | null) ?? null,
    photo: r.photoName,
    claimRef: r.claimId ? (claimRef ?? 'claimed') : null,
    used: r.used == null ? null : Number(r.used),
  }
}

export interface GasdsClaimDto {
  id: string
  taxYear: number
  amount: number
  hmrcRef: string
  sentAt: string
  sentBy: string
  giftAidFigure: number
  status: 'submitted' | 'received'
  receivedAmount: number | null
  receivedAt: string | null
}
export const toClaimDto = (c: ClaimRow): GasdsClaimDto => ({
  id: c.id,
  taxYear: c.taxYear,
  amount: Number(c.amount),
  hmrcRef: c.hmrcRef,
  sentAt: c.sentAt.toISOString(),
  sentBy: c.sentBy,
  giftAidFigure: Number(c.giftAidFigure),
  status: c.status === 'received' ? 'received' : 'submitted',
  receivedAmount: c.receivedAmount == null ? null : Number(c.receivedAmount),
  receivedAt: c.receivedAt?.toISOString() ?? null,
})

/** The charity id from a request, checked to exist. */
export async function requireCharity(id: unknown): Promise<string> {
  if (typeof id !== 'string' || !UUID_RE.test(id)) throw new RangeError('Choose a charity first')
  const [c] = await getDb().select({ id: charities.id }).from(charities).where(eq(charities.id, id))
  if (!c) throw new RangeError('That charity no longer exists')
  return c.id
}

/** Gift Aid figure for a tax year: donations in the charity's submitted and received Gift Aid claims (by donation date), plus the figure claimed outside BoostAid. */
export async function giftAidFigures(charityId: string, years: number[]): Promise<Record<number, { boostaid: number; elsewhere: number }>> {
  const db = getDb()
  const out: Record<number, { boostaid: number; elsewhere: number }> = {}
  const elsewhere = years.length
    ? await db.select().from(gasdsYears).where(and(eq(gasdsYears.charityId, charityId), inArray(gasdsYears.taxYear, years)))
    : []
  for (const y of years) {
    const { from, to } = taxYearRange(y)
    const res = await db.execute<{ total: string }>(sql`
      select coalesce(sum(d.amount), 0)::text as total
      from donations d
      join claims c on c.id = d.claim_id
      join batches b on b.id = c.batch_id
      where b.charity_id = ${charityId} and c.status in ('submitted', 'received')
        and d.donation_date >= ${from}::date and d.donation_date <= ${to}::date
    `)
    out[y] = {
      boostaid: Number(res.rows[0].total),
      elsewhere: Number(elsewhere.find((e) => e.taxYear === y)?.giftAidElsewhere ?? 0),
    }
  }
  return out
}

/** All stored receipts of a charity for the given tax years, with the HMRC ref of any claim they are in. */
export async function loadReceipts(charityId: string, years: number[]): Promise<GReceipt[]> {
  if (!years.length) return []
  const db = getDb()
  const rows = await db
    .select({ r: gasdsReceipts, ref: gasdsClaims.hmrcRef })
    .from(gasdsReceipts)
    .leftJoin(gasdsClaims, eq(gasdsClaims.id, gasdsReceipts.claimId))
    .where(and(eq(gasdsReceipts.charityId, charityId), inArray(gasdsReceipts.taxYear, years)))
  return rows.map((x) => toReceipt(x.r, x.ref))
}

/** Total of the saved claims per tax year. */
export function claimedByYear(claims: { taxYear: number; amount: number }[]): Record<number, number> {
  const out: Record<number, number> = {}
  for (const c of claims) out[c.taxYear] = Math.round(((out[c.taxYear] ?? 0) + c.amount) * 100) / 100
  return out
}
