import * as XLSX from 'xlsx'
import { NextResponse } from 'next/server'
import { eq } from 'drizzle-orm'
import { charities, getDb, invoices } from '@/lib/db'
import { withAuth, requireAdmin } from '@/lib/api'
import { xlsxResponse } from '@/lib/export'

/** The paid invoices (the Paid archive) as Excel, for the accountant. `?month=2026-11` limits it to one month (by date paid). */
export const GET = withAuth(async (req) => {
  await requireAdmin(req)
  const month = req.nextUrl.searchParams.get('month')
  const rows = await getDb()
    .select({ inv: invoices, charity: charities.name })
    .from(invoices)
    .innerJoin(charities, eq(charities.id, invoices.charityId))
    .where(eq(invoices.status, 'paid'))
  const uk = (d: Date | null) => (d ? d.toLocaleDateString('en-GB', { timeZone: 'Europe/London' }) : '')
  const picked = rows
    .filter((r) => !month || (r.inv.paidAt && r.inv.paidAt.toISOString().slice(0, 7) === month))
    .sort((a, b) => (b.inv.paidAt?.getTime() ?? 0) - (a.inv.paidAt?.getTime() ?? 0))
  const sheet = XLSX.utils.json_to_sheet(
    picked.map((r) => ({
      Charity: r.charity,
      Invoice: r.inv.number,
      'Recovered Gift Aid paid (GBP)': Number(r.inv.recoveredPaid),
      'Fee 5% (GBP)': Number(r.inv.fee),
      Sent: uk(r.inv.sentAt),
      Paid: uk(r.inv.paidAt),
    }))
  )
  sheet['!cols'] = [30, 24, 28, 16, 14, 14].map((wch) => ({ wch }))
  XLSX.utils.sheet_add_aoa(sheet, [['Total', '', '', picked.reduce((s, r) => s + Number(r.inv.fee), 0)]], { origin: -1 })
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, sheet, 'Paid invoices')
  return xlsxResponse(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer, `boostaid_paid_invoices${month ? '_' + month : ''}.xlsx`)
})
