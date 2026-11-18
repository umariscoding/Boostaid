import { NextResponse } from 'next/server'
import { asc, eq } from 'drizzle-orm'
import * as XLSX from 'xlsx'
import { claims, donations, getDb } from '@/lib/db'
import { UUID_RE, withAuth } from '@/lib/api'
import { xlsxResponse } from '@/lib/export'

/** The donations in one claim, as shown on the claim page, as an Excel file. */
export const GET = withAuth<{ id: string }>(async (_req, { params }) => {
  const { id } = await params
  if (!UUID_RE.test(id)) return NextResponse.json({ error: 'Not found' }, { status: 404 })
  const db = getDb()
  const [claim] = await db.select({ ref: claims.ref }).from(claims).where(eq(claims.id, id))
  if (!claim) return NextResponse.json({ error: 'Not found' }, { status: 404 })

  const rows = await db.select().from(donations).where(eq(donations.claimId, id)).orderBy(asc(donations.lastName), asc(donations.firstName))
  const sheet = XLSX.utils.json_to_sheet(
    rows.map((d) => ({
      Title: d.title,
      'First name': d.firstName,
      'Last name': d.lastName,
      Address: d.address,
      Postcode: d.postcode,
      'Donation date': d.donationDate ? d.donationDate.split('-').reverse().join('/') : '',
      'Amount (GBP)': Number(d.amount),
    }))
  )
  sheet['!cols'] = [8, 16, 18, 28, 12, 14, 14].map((wch) => ({ wch }))
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, sheet, 'Donations')
  return xlsxResponse(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer, `${claim.ref}_donations.xlsx`)
})
