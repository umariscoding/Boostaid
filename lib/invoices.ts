import { and, asc, eq, sql } from 'drizzle-orm'
import { charities, donations, getDb, invoices } from '@/lib/db'
import { addNotification, claimsWithRecovery, type ClaimRecovery } from '@/lib/admin-db'
import { buildInvoicePdf, business, type InvoiceData } from '@/lib/invoice-pdf'
import { addDays } from '@/lib/recovery'
import { esc, sendMail } from '@/lib/email'

export const INVOICE_DAYS = 14
const gbp = (n: number) => '£' + n.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
export const invoiceNumber = (claimRef: string) => `BA-${claimRef}`

async function invoiceData(c: ClaimRecovery, sentAt: Date, dueAt: Date): Promise<InvoiceData> {
  const rows = await getDb()
    .select({ date: donations.donationDate, first: donations.firstName, last: donations.lastName, amount: donations.amount })
    .from(donations)
    .where(and(eq(donations.claimId, c.claimId), eq(donations.recovered, true)))
    .orderBy(asc(donations.donationDate))
  return {
    number: invoiceNumber(c.ref),
    sentAt,
    dueAt,
    charity: c.charityName,
    claimRef: c.ref,
    hmrcRef: c.hmrcRef,
    expected: c.expected,
    received: c.received,
    recoveredGiftAid: c.recoveredGiftAid,
    recoveredPaid: c.recoveredPaid,
    fee: c.fee,
    donations: rows.map((r) => ({ date: r.date, donor: `${r.first} ${r.last}`.trim(), amount: Number(r.amount) })),
  }
}

async function emailInvoice(c: ClaimRecovery, to: string, data: InvoiceData, resend = false) {
  const pdf = await buildInvoicePdf(data)
  const b = business()
  const subject = `${resend ? 'Reminder: i' : 'I'}nvoice ${data.number}: Gift Aid recovered on ${c.ref}`
  const text =
    `Hello,\n\nHMRC has paid claim ${c.ref}. BoostAid recovered ${gbp(c.recoveredGiftAid)} Gift Aid in it by emailing donors, and HMRC paid ${gbp(c.recoveredPaid)} of that. ` +
    `Our fee is 5% of it: ${gbp(c.fee)}.\n\nThe invoice is attached, with the list of recovered donations. It is due on ${data.dueAt.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })}.\n\nThank you,\n${b.name}`
  const html = `<div style="font-family:Arial,sans-serif;font-size:15px;color:#17152B;max-width:560px">${text.split('\n\n').map((p) => `<p>${esc(p).replace(/\n/g, '<br>')}</p>`).join('')}</div>`
  return sendMail({ to, fromName: b.name, subject, text, html, attachments: [{ name: `${data.number}.pdf`, contentType: 'application/pdf', data: pdf }] })
}

/** What sending would do, for the preview. */
export async function previewInvoice(claimId: string) {
  const [c] = await claimsWithRecovery({ claimId })
  if (!c) throw new RangeError('That claim has no recovered Gift Aid')
  const [ch] = await getDb().select().from(charities).where(eq(charities.id, c.charityId))
  const sentAt = new Date()
  const data = await invoiceData(c, sentAt, addDays(sentAt, INVOICE_DAYS))
  return { ...data, sentAt: data.sentAt.toISOString(), dueAt: data.dueAt.toISOString(), to: ch?.contactEmail ?? '', canSend: c.status === 'received' }
}

/**
 * Sends the invoice for one claim. Only possible once the claim is received, and only once: the invoice row is claimed
 * first, under a lock for that claim, so a double click (or two people at once) can never email it twice.
 */
export async function sendInvoice(claimId: string) {
  const db = getDb()
  const [c] = await claimsWithRecovery({ claimId })
  if (!c) throw new RangeError('That claim has no recovered Gift Aid')
  if (c.status !== 'received') throw new RangeError('An invoice can only be sent once the claim is marked received')
  const [ch] = await db.select().from(charities).where(eq(charities.id, c.charityId))
  if (!ch?.contactEmail) throw new RangeError(`Add a contact email for ${c.charityName} in Settings first (the invoice is sent there)`)

  const sentAt = new Date()
  const dueAt = addDays(sentAt, INVOICE_DAYS)
  const data = await invoiceData(c, sentAt, dueAt)

  // 1. Claim the invoice. A second request for the same claim waits here, then finds it taken and stops.
  const inv = await db.transaction(async (tx) => {
    await tx.execute(sql`select pg_advisory_xact_lock(hashtext(${'invoice:' + claimId}))`)
    const [existing] = await tx.select().from(invoices).where(eq(invoices.claimId, claimId))
    if (existing) throw new RangeError(`${existing.number} has already been sent`)
    const [created] = await tx
      .insert(invoices)
      .values({ charityId: c.charityId, claimId, number: data.number, recoveredPaid: c.recoveredPaid.toFixed(2), fee: c.fee.toFixed(2), sentAt, dueAt, sentTo: ch.contactEmail!, delivered: false })
      .returning()
    return created
  })

  // 2. Email it. If that fails, give the claim back so it can be sent again.
  let delivered: boolean
  try {
    delivered = (await emailInvoice(c, ch.contactEmail, data)).delivered
  } catch (e) {
    await db.delete(invoices).where(eq(invoices.id, inv.id))
    throw e
  }
  await db.update(invoices).set({ delivered }).where(eq(invoices.id, inv.id))
  await addNotification({
    type: 'invoice_sent',
    charityId: c.charityId,
    claimId,
    text: delivered ? `Invoice ${data.number} emailed to ${c.charityName}.` : `Invoice ${data.number} recorded for ${c.charityName}. It was NOT emailed because no email service is connected.`,
  })
  return { ...inv, delivered }
}

export async function resendInvoice(invoiceId: string) {
  const db = getDb()
  const [inv] = await db.select().from(invoices).where(eq(invoices.id, invoiceId))
  if (!inv) throw new RangeError('Invoice not found')
  const [c] = await claimsWithRecovery({ claimId: inv.claimId })
  if (!c) throw new RangeError('That claim has no recovered Gift Aid')
  const data = await invoiceData({ ...c, fee: Number(inv.fee), recoveredPaid: Number(inv.recoveredPaid) }, inv.sentAt, inv.dueAt)
  const { delivered } = await emailInvoice(c, inv.sentTo, data, true)
  await db.update(invoices).set({ delivered }).where(eq(invoices.id, invoiceId))
  return delivered
}

export async function markPaid(invoiceId: string) {
  const db = getDb()
  const [inv] = await db.select().from(invoices).where(eq(invoices.id, invoiceId))
  if (!inv) throw new RangeError('Invoice not found')
  if (inv.status === 'paid') throw new RangeError('That invoice is already marked paid')
  await db.update(invoices).set({ status: 'paid', paidAt: new Date() }).where(eq(invoices.id, invoiceId))
  const [ch] = await db.select().from(charities).where(eq(charities.id, inv.charityId))
  await addNotification({ type: 'invoice_paid', charityId: inv.charityId, claimId: inv.claimId, text: `Invoice ${inv.number} paid by ${ch?.name ?? 'the charity'} (${gbp(Number(inv.fee))}).` })
}

export async function invoicePdfFor(invoiceId: string) {
  const [inv] = await getDb().select().from(invoices).where(eq(invoices.id, invoiceId))
  if (!inv) return null
  const [c] = await claimsWithRecovery({ claimId: inv.claimId })
  if (!c) return null
  const data = await invoiceData({ ...c, fee: Number(inv.fee), recoveredPaid: Number(inv.recoveredPaid) }, inv.sentAt, inv.dueAt)
  return { name: `${inv.number}.pdf`, pdf: await buildInvoicePdf(data) }
}

