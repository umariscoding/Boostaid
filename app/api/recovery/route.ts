import { NextResponse } from 'next/server'
import { and, desc, eq, ne } from 'drizzle-orm'
import { getDb, reviewRows } from '@/lib/db'
import { withAuth } from '@/lib/api'
import { emailConfigured } from '@/lib/email'
import { giftAidOn } from '@/lib/claims'
import { requireCharity } from '@/lib/gasds-db'

export interface RecoveryRowDto {
  id: string
  platform: string
  donationId: string
  date: string | null
  amount: number
  reason: string
  name: string
  house: string
  postcode: string
  email: string
  status: string
  emailSentAt: string | null
  reminderAt: string | null
  repliedAt: string | null
  fixedBy: string | null
  fixed: { title: string; first: string; last: string; house: string; postcode: string } | null
}

/** The held-back donations of one charity (everything not yet in a claim), with the figures for the tiles. */
export const GET = withAuth(async (req) => {
  const charityId = await requireCharity(req.nextUrl.searchParams.get('charity'))
  const rows = await getDb()
    .select()
    .from(reviewRows)
    .where(and(eq(reviewRows.charityId, charityId), ne(reviewRows.status, 'claimed')))
    .orderBy(desc(reviewRows.createdAt), reviewRows.donationDate)
  const dto: RecoveryRowDto[] = rows.map((r) => ({
    id: r.id,
    platform: r.platform,
    donationId: r.donationId,
    date: r.donationDate,
    amount: Number(r.amount),
    reason: r.reason,
    name: `${r.firstName} ${r.lastName}`.trim(),
    house: r.house,
    postcode: r.postcode,
    email: r.email ?? '',
    status: r.status,
    emailSentAt: r.emailSentAt?.toISOString() ?? null,
    reminderAt: r.reminderAt?.toISOString() ?? null,
    repliedAt: r.repliedAt?.toISOString() ?? null,
    fixedBy: r.fixedBy,
    fixed: r.status === 'recovered' ? { title: r.fixedTitle ?? '', first: r.fixedFirst ?? '', last: r.fixedLast ?? '', house: r.fixedHouse ?? '', postcode: r.fixedPostcode ?? '' } : null,
  }))
  const sum = (f: (r: RecoveryRowDto) => boolean) => giftAidOn(dto.filter(f).reduce((s, r) => s + r.amount, 0))
  const n = (f: (r: RecoveryRowDto) => boolean) => dto.filter(f).length
  return NextResponse.json({
    rows: dto,
    emailConfigured: emailConfigured(),
    tiles: {
      heldBack: { count: n((r) => r.status === 'none'), giftAid: sum((r) => r.status === 'none') },
      waiting: { count: n((r) => r.status === 'sent' || r.status === 'reminded'), giftAid: sum((r) => r.status === 'sent' || r.status === 'reminded') },
      recovered: { count: n((r) => r.status === 'recovered'), giftAid: sum((r) => r.status === 'recovered') },
      noResponse: { count: n((r) => r.status === 'no_response') },
    },
  })
})
