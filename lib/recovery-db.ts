import { and, eq, inArray, isNotNull, lte, sql } from 'drizzle-orm'
import { charities, getDb, recoveryLog, recoveryRequests, reviewRows } from '@/lib/db'
import { emailConfigured, esc, sendMail } from '@/lib/email'
import {
  LINK_VALID_DAYS,
  NO_RESPONSE_AFTER_DAYS,
  REMINDER_AFTER_DAYS,
  addDays,
  checkAnswers,
  groupByDonor,
  reconcileUpload,
  rowKey,
  type HeldBackRow,
} from '@/lib/recovery'
import { hashCode, newCode } from '@/lib/recovery-codes'
import { giftAidOn } from '@/lib/claims'

type Db = ReturnType<typeof getDb>
type Tx = Parameters<Parameters<Db['transaction']>[0]>[0]
export type ReviewRowDb = typeof reviewRows.$inferSelect

const gbp = (n: number) => '£' + n.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
const ukDate = (iso: string | null) => (iso ? iso.split('-').reverse().join('/') : '')

// ---------------------------------------------------------------------------------------------- the email
/** One email per donor: it lists every donation they have held back, with one link that fixes them all. */
export function donorEmail(p: {
  charity: string
  contactEmail: string | null
  rows: { donationDate: string | null; amount: string | number }[]
  link: string
  firstName: string
  reminder: boolean
}) {
  const total = p.rows.reduce((s, r) => s + Number(r.amount), 0)
  const one = p.rows.length === 1
  const subject = `${p.reminder ? 'Reminder: ' : ''}Help us claim Gift Aid on your ${one ? gbp(total) + ' donation' : p.rows.length + ' donations'}`
  const lines = p.rows.map((r) => `${gbp(Number(r.amount))} on ${ukDate(r.donationDate)}`)
  const hi = p.firstName ? `Dear ${p.firstName},` : 'Hello,'
  const text =
    `${hi}\n\nThank you for your ${one ? 'donation' : 'donations'} to ${p.charity}: ${lines.join('; ')}. You ticked Gift Aid, so HMRC adds ${gbp(giftAidOn(total))} at no cost to you.\n\n` +
    `We're missing a detail HMRC needs. It takes under a minute to confirm:\n${p.link}\n\n` +
    `This is a one-off email about your ${one ? 'donation' : 'donations'}, not marketing. The link works once and expires after ${LINK_VALID_DAYS} days.` +
    (p.contactEmail ? `\nQuestions? Reply to this email or write to ${p.contactEmail}.` : '')
  const html =
    `<div style="font-family:Arial,sans-serif;font-size:15px;color:#17152B;max-width:560px"><p>${esc(hi)}</p>` +
    `<p>Thank you for your ${one ? 'donation' : 'donations'} to <b>${esc(p.charity)}</b>:</p><ul>${lines.map((l) => `<li>${esc(l)}</li>`).join('')}</ul>` +
    `<p>You ticked Gift Aid, so HMRC adds <b>${gbp(giftAidOn(total))}</b> at no cost to you.</p>` +
    `<p>We're missing a detail HMRC needs. It takes under a minute to confirm:</p>` +
    `<p><a href="${esc(p.link)}" style="display:inline-block;background:#4C1D95;color:#fff;padding:12px 22px;border-radius:999px;text-decoration:none;font-weight:bold">Confirm my details</a></p>` +
    `<p style="font-size:13px;color:#5B5874">This is a one-off email about your ${one ? 'donation' : 'donations'}, not marketing. The link works once and expires after ${LINK_VALID_DAYS} days.` +
    (p.contactEmail ? ` Questions? Reply to this email or write to ${esc(p.contactEmail)}.` : '') +
    `</p></div>`
  return { subject, text, html }
}

export const confirmUrl = (base: string, code: string) => `${base}/confirm/${code}`
const fromName = (charity: string) => `${charity} (sent by BoostAid)`

// ---------------------------------------------------------------------------------------------- sending
export interface SendOutcome {
  sent: number
  failed: number
  emailConfigured: boolean
  /** Only when no email service is connected: the links, so the form can be tested. Never returned when real emails go out. */
  testLinks: { to: string; name: string; url: string }[]
}

/** Emails the donors of the chosen rows: one email per donor address. Rows must still be "Needs review" and have an email. */
export async function sendRecoveryEmails(charityId: string, rowIds: string[], base: string): Promise<SendOutcome> {
  const db = getDb()
  const [charity] = await db.select().from(charities).where(eq(charities.id, charityId))
  if (!charity) throw new RangeError('That charity no longer exists')
  const rows = rowIds.length
    ? await db.select().from(reviewRows).where(and(eq(reviewRows.charityId, charityId), inArray(reviewRows.id, rowIds), eq(reviewRows.status, 'none'), isNotNull(reviewRows.email)))
    : []
  if (!rows.length) throw new RangeError('None of those donations can be emailed (they need an email address and must still be waiting for review)')

  const out: SendOutcome = { sent: 0, failed: 0, emailConfigured: emailConfigured(), testLinks: [] }
  for (const [email, group] of groupByDonor(rows)) {
    const code = newCode()
    const now = new Date()
    const requestId = await db.transaction(async (tx) => {
      const [req] = await tx
        .insert(recoveryRequests)
        .values({ charityId, codeHash: hashCode(code), email, expiresAt: addDays(now, LINK_VALID_DAYS) })
        .returning()
      await tx.update(reviewRows).set({ status: 'sent', requestId: req.id, emailSentAt: now }).where(inArray(reviewRows.id, group.map((g) => g.id)))
      return req.id
    })
    const link = confirmUrl(base, code)
    const mail = donorEmail({ charity: charity.name, contactEmail: charity.contactEmail, rows: group, link, firstName: group[0].firstName.split(' ')[0] ?? '', reminder: false })
    try {
      const { delivered } = await sendMail({ to: email, fromName: fromName(charity.name), replyTo: charity.contactEmail, ...mail })
      await db.insert(recoveryLog).values({ charityId, requestId, type: 'email_sent', toEmail: email, delivered, detail: `${group.length} donation${group.length === 1 ? '' : 's'}${delivered ? '' : ' (no email service connected: not delivered)'}` })
      out.sent += group.length
      if (!delivered) out.testLinks.push({ to: email, name: `${group[0].firstName} ${group[0].lastName}`.trim(), url: link })
    } catch (e) {
      // Not sent: put the rows back so they can be tried again, and forget the link
      await db.transaction(async (tx) => {
        await tx.update(reviewRows).set({ status: 'none', requestId: null, emailSentAt: null }).where(inArray(reviewRows.id, group.map((g) => g.id)))
        await tx.delete(recoveryRequests).where(eq(recoveryRequests.id, requestId))
        await tx.insert(recoveryLog).values({ charityId, type: 'email_sent', toEmail: email, delivered: false, detail: `FAILED: ${(e as Error).message}` })
      })
      out.failed += group.length
    }
  }
  return out
}

/** What the email would look like, for the preview before sending. Uses a placeholder link: no code exists yet. */
export async function previewRecoveryEmail(charityId: string, rowIds: string[], base: string) {
  const db = getDb()
  const [charity] = await db.select().from(charities).where(eq(charities.id, charityId))
  if (!charity) throw new RangeError('That charity no longer exists')
  const rows = rowIds.length
    ? await db.select().from(reviewRows).where(and(eq(reviewRows.charityId, charityId), inArray(reviewRows.id, rowIds), eq(reviewRows.status, 'none'), isNotNull(reviewRows.email)))
    : []
  const groups = [...groupByDonor(rows)]
  if (!groups.length) throw new RangeError('None of those donations can be emailed (they need an email address and must still be waiting for review)')
  const [email, group] = groups[0]
  const mail = donorEmail({ charity: charity.name, contactEmail: charity.contactEmail, rows: group, link: confirmUrl(base, 'XXXXXXXXXX'), firstName: group[0].firstName.split(' ')[0] ?? '', reminder: false })
  return {
    donors: groups.length,
    donations: rows.length,
    from: fromName(charity.name),
    replyTo: charity.contactEmail,
    to: email,
    subject: mail.subject,
    html: mail.html,
    emailConfigured: emailConfigured(),
  }
}

// ---------------------------------------------------------------------------------------------- the daily job
export interface DueOutcome { reminders: number; noResponse: number }

/**
 * Runs every day: one reminder to anyone who hasn't replied after 14 days (never more than 2 emails), then "No response"
 * 14 days after the reminder, when the link stops working. `now` can be set to test.
 */
export async function processDue(base: string, now = new Date()): Promise<DueOutcome> {
  const db = getDb()
  const out: DueOutcome = { reminders: 0, noResponse: 0 }

  // 1. reminders
  const due = await db
    .select()
    .from(reviewRows)
    .where(and(eq(reviewRows.status, 'sent'), lte(reviewRows.emailSentAt, addDays(now, -REMINDER_AFTER_DAYS)), isNotNull(reviewRows.requestId)))
  const byRequest = new Map<string, ReviewRowDb[]>()
  for (const r of due) byRequest.set(r.requestId!, [...(byRequest.get(r.requestId!) ?? []), r])
  for (const [requestId, rows] of byRequest) {
    const [root] = await db.select().from(recoveryRequests).where(eq(recoveryRequests.id, requestId))
    const [charity] = await db.select().from(charities).where(eq(charities.id, rows[0].charityId))
    if (!root || !charity || root.usedAt || root.expiresAt <= now) continue
    const code = newCode()
    await db.transaction(async (tx) => {
      await tx.insert(recoveryRequests).values({ charityId: charity.id, codeHash: hashCode(code), email: root.email, parentId: root.id, expiresAt: root.expiresAt })
      await tx.update(reviewRows).set({ status: 'reminded', reminderAt: now }).where(inArray(reviewRows.id, rows.map((r) => r.id)))
    })
    const mail = donorEmail({ charity: charity.name, contactEmail: charity.contactEmail, rows, link: confirmUrl(base, code), firstName: rows[0].firstName.split(' ')[0] ?? '', reminder: true })
    let delivered = false
    try {
      delivered = (await sendMail({ to: root.email, fromName: fromName(charity.name), replyTo: charity.contactEmail, ...mail })).delivered
    } catch (e) {
      await db.insert(recoveryLog).values({ charityId: charity.id, requestId, type: 'reminder_sent', toEmail: root.email, delivered: false, detail: `FAILED: ${(e as Error).message}` })
      continue
    }
    await db.insert(recoveryLog).values({ charityId: charity.id, requestId, type: 'reminder_sent', toEmail: root.email, delivered, detail: `${rows.length} donation${rows.length === 1 ? '' : 's'}${delivered ? '' : ' (no email service connected: not delivered)'}` })
    out.reminders += rows.length
  }

  // 2. no response, and the link stops working
  const giveUp = await db
    .select()
    .from(reviewRows)
    .where(and(eq(reviewRows.status, 'reminded'), lte(reviewRows.reminderAt, addDays(now, -NO_RESPONSE_AFTER_DAYS))))
  if (giveUp.length) {
    await db.transaction(async (tx) => {
      await tx.update(reviewRows).set({ status: 'no_response' }).where(inArray(reviewRows.id, giveUp.map((r) => r.id)))
      const ids = [...new Set(giveUp.map((r) => r.requestId).filter((x): x is string => !!x))]
      if (ids.length) {
        await tx.update(recoveryRequests).set({ expiresAt: now }).where(sql`${recoveryRequests.id} in ${ids} or ${recoveryRequests.parentId} in ${ids}`)
      }
      await tx.insert(recoveryLog).values(giveUp.map((r) => ({ charityId: r.charityId, requestId: r.requestId, type: 'no_response', toEmail: r.email, detail: 'No reply after the reminder; the link has expired' })))
    })
    out.noResponse = giveUp.length
  }
  return out
}

// ---------------------------------------------------------------------------------------------- a charity fixes a row by hand
export async function fixByHand(charityId: string, rowId: string, a: { title: string; first: string; last: string; house: string; postcode: string }) {
  const db = getDb()
  return db.transaction(async (tx) => {
    const [row] = await tx.select().from(reviewRows).where(and(eq(reviewRows.id, rowId), eq(reviewRows.charityId, charityId))).for('update')
    if (!row) throw new RangeError('That donation no longer exists')
    if (!['none', 'no_response', 'sent', 'reminded'].includes(row.status)) throw new RangeError('That donation is already fixed')
    const checked = checkAnswers(a, row.donationDate ?? '2026-01-01', Number(row.amount))
    if (!checked.ok) throw new RangeError(checked.error)
    const c = checked.cleaned
    await tx
      .update(reviewRows)
      .set({ status: 'recovered', fixedBy: 'hand', fixedTitle: c.title, fixedFirst: c.first, fixedLast: c.last, fixedHouse: c.house, fixedPostcode: c.postcode })
      .where(eq(reviewRows.id, rowId))
    await tx.insert(recoveryLog).values({ charityId, requestId: row.requestId, type: 'fixed_by_hand', detail: 'Fixed by the charity (no fee)' })
  })
}

// ---------------------------------------------------------------------------------------------- saving a claim meets the held-back rows
/**
 * Called inside the claim-saving transaction.
 *  - new rows that need review are stored (a row already stored stays as it is: no new row, no second email)
 *  - a stored row that is now clean in this file is fixed from the file: it is in this claim, no fee, its link stops working
 * `claimedByKey` maps platform+donation ID to the claim that now holds it.
 */
export async function applyUpload(tx: Tx, charityId: string, held: HeldBackRow[], claimedByKey: Map<string, string>) {
  const existing = await tx
    .select({ id: reviewRows.id, platform: reviewRows.platform, donationId: reviewRows.donationId, status: reviewRows.status })
    .from(reviewRows)
    .where(eq(reviewRows.charityId, charityId))
  const rec = reconcileUpload(
    existing.map((e) => ({ key: rowKey(e.platform, e.donationId), status: e.status, id: e.id })),
    held.map((h) => rowKey(h.platform, h.donationId)),
    [...claimedByKey.keys()]
  )

  const addKeys = new Set(rec.add)
  const fresh = held.filter((h, i, all) => addKeys.has(rowKey(h.platform, h.donationId)) && all.findIndex((x) => rowKey(x.platform, x.donationId) === rowKey(h.platform, h.donationId)) === i)
  for (let i = 0; i < fresh.length; i += 500) {
    await tx
      .insert(reviewRows)
      .values(
        fresh.slice(i, i + 500).map((h) => ({
          charityId,
          platform: h.platform,
          donationId: h.donationId,
          donationDate: h.donationDate,
          amount: h.amount.toFixed(2),
          title: h.title,
          firstName: h.firstName,
          lastName: h.lastName,
          house: h.house,
          postcode: h.postcode,
          reason: h.reason,
          sourceFile: h.sourceFile,
          email: h.email || null,
        }))
      )
      .onConflictDoNothing()
  }

  const fixedKeys = new Set(rec.fixedFromFile)
  const toFix = existing.filter((e) => fixedKeys.has(rowKey(e.platform, e.donationId)) && e.status !== 'claimed')
  for (const e of toFix) {
    await tx
      .update(reviewRows)
      .set({ status: 'claimed', fixedBy: 'file', claimId: claimedByKey.get(rowKey(e.platform, e.donationId)) ?? null })
      .where(eq(reviewRows.id, e.id))
  }
  if (toFix.length) {
    await tx.insert(recoveryLog).values(toFix.map(() => ({ charityId, type: 'fixed_from_file', detail: 'Fixed by a newer export: no fee' })))
  }
  return { added: fresh.length, fixedFromFile: toFix.length }
}
