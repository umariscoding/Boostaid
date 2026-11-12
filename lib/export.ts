import * as XLSX from 'xlsx'
import { asc, desc, eq, inArray, type SQL } from 'drizzle-orm'
import { batches, charities, claimEvents, claims, donations, getDb } from '@/lib/db'

const fmtDateTime = (d: Date | null) =>
  d ? d.toLocaleString('en-GB', { timeZone: 'Europe/London', dateStyle: 'short', timeStyle: 'short' }) : ''

const fmtDate = (iso: string | null) => (iso ? iso.split('-').reverse().join('/') : '')

const STATUS_LABEL = { pending: 'Pending', submitted: 'Submitted', received: 'Received' } as const

/** Builds the audit workbook (Claims, Donations, Status History) for the given claims filter. */
export async function buildAuditWorkbook(where?: SQL): Promise<Buffer> {
  const db = getDb()

  const claimRows = await db
    .select({ claim: claims, filename: batches.filename, charity: charities.name })
    .from(claims)
    .innerJoin(batches, eq(batches.id, claims.batchId))
    .innerJoin(charities, eq(charities.id, batches.charityId))
    .where(where)
    .orderBy(desc(claims.createdAt), desc(claims.chunkIndex))

  const ids = claimRows.map((r) => r.claim.id)
  const refById = new Map(claimRows.map((r) => [r.claim.id, r.claim]))
  const charityById = new Map(claimRows.map((r) => [r.claim.id, r.charity]))

  // Chunked so a large export never exceeds the Postgres parameter limit
  const donationRows: (typeof donations.$inferSelect)[] = []
  const eventRows: (typeof claimEvents.$inferSelect)[] = []
  for (let i = 0; i < ids.length; i += 200) {
    const slice = ids.slice(i, i + 200)
    donationRows.push(
      ...(await db.select().from(donations).where(inArray(donations.claimId, slice)).orderBy(asc(donations.lastName), asc(donations.firstName)))
    )
    eventRows.push(...(await db.select().from(claimEvents).where(inArray(claimEvents.claimId, slice)).orderBy(asc(claimEvents.at))))
  }

  const wb = XLSX.utils.book_new()

  const claimsSheet = XLSX.utils.json_to_sheet(
    claimRows.map(({ claim: c, filename, charity }) => ({
      Charity: charity,
      'Claim Reference': c.ref,
      'Schedule file': c.fileName ?? '',
      Type: c.kind,
      'HMRC Reference': c.hmrcRef ?? '',
      Status: STATUS_LABEL[c.status],
      'Created': fmtDateTime(c.createdAt),
      'Submitted': fmtDateTime(c.submittedAt),
      'Received': fmtDateTime(c.receivedAt),
      Donations: c.donationCount,
      'Gift Aid Amount (GBP)': Number(c.totalAmount),
      'Source File': filename,
      Notes: c.notes ?? '',
    }))
  )
  claimsSheet['!cols'] = [26, 18, 26, 10, 18, 12, 18, 18, 18, 10, 20, 30, 40].map((wch) => ({ wch }))
  XLSX.utils.book_append_sheet(wb, claimsSheet, 'Claims')

  const donationsSheet = XLSX.utils.json_to_sheet(
    donationRows.map((d) => ({
      Charity: charityById.get(d.claimId) ?? '',
      'Claim Reference': refById.get(d.claimId)?.ref ?? '',
      'Claim Status': STATUS_LABEL[refById.get(d.claimId)?.status ?? 'pending'],
      Title: d.title,
      'First Name': d.firstName,
      'Last Name': d.lastName,
      Address: d.address,
      Postcode: d.postcode,
      Platform: d.platform,
      'Platform donation ID': d.platformDonationId ?? '',
      'Donation Date': fmtDate(d.donationDate),
      'Amount (GBP)': Number(d.amount),
    }))
  )
  donationsSheet['!cols'] = [26, 18, 12, 8, 16, 16, 34, 12, 16, 22, 14, 14].map((wch) => ({ wch }))
  XLSX.utils.book_append_sheet(wb, donationsSheet, 'Donations')

  const historySheet = XLSX.utils.json_to_sheet(
    eventRows.map((e) => ({
      Charity: charityById.get(e.claimId) ?? '',
      'Claim Reference': refById.get(e.claimId)?.ref ?? '',
      Status: STATUS_LABEL[e.status],
      'Changed At': fmtDateTime(e.at),
      Note: e.note ?? '',
    }))
  )
  historySheet['!cols'] = [26, 18, 12, 18, 40].map((wch) => ({ wch }))
  XLSX.utils.book_append_sheet(wb, historySheet, 'Status History')

  return XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }) as Buffer
}

export function xlsxResponse(buffer: Buffer, filename: string): Response {
  return new Response(new Uint8Array(buffer), {
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${filename}"`,
      'Cache-Control': 'no-store',
    },
  })
}
