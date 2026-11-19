import { and, eq, sql } from 'drizzle-orm'
import { adminNotifications, charities, claims, donations, getDb, users } from '@/lib/db'
import { giftAidOn } from '@/lib/claims'
import { feeFigures } from '@/lib/fees'
import { esc, sendMail } from '@/lib/email'

const gbp = (n: number) => '£' + n.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })

/** Everyone with the admin role (BoostAid's owner). Notifications are emailed to them as well as shown on the admin page. */
export async function adminEmails(): Promise<string[]> {
  const rows = await getDb().select({ email: users.email }).from(users).where(eq(users.role, 'admin'))
  return rows.map((r) => r.email)
}

/** Records a notification on the admin page and emails it to the admin(s). Email trouble never stops the action itself. */
export async function addNotification(n: { type: string; charityId: string | null; claimId: string | null; text: string }) {
  const db = getDb()
  await db.insert(adminNotifications).values(n)
  try {
    for (const to of await adminEmails()) {
      await sendMail({
        to,
        fromName: 'BoostAid',
        subject: n.text.length > 90 ? n.text.slice(0, 87) + '...' : n.text,
        text: n.text,
        html: `<div style="font-family:Arial,sans-serif;font-size:15px;color:#17152B"><p>${esc(n.text)}</p></div>`,
      })
    }
  } catch (e) {
    console.error('notification email failed', e)
  }
}

export interface ClaimRecovery {
  claimId: string
  ref: string
  charityId: string
  charityName: string
  status: 'pending' | 'submitted' | 'received'
  hmrcRef: string | null
  createdAt: Date
  submittedAt: Date | null
  receivedAt: Date | null
  expected: number // Gift Aid expected on the whole claim
  received: number | null
  recoveredCount: number
  recoveredGiftAid: number // Gift Aid on the recovered donations (25%)
  recoveredPaid: number // the part HMRC actually paid
  fee: number
}

/** Every claim that has recovered Gift Aid in it (optionally one charity or one claim), with the 5% fee worked out. */
export async function claimsWithRecovery(filter?: { claimId?: string; charityId?: string }): Promise<ClaimRecovery[]> {
  const db = getDb()
  const rows = await db
    .select({
      claim: claims,
      charityId: charities.id,
      charityName: charities.name,
      recAmount: sql<string>`(select coalesce(sum(d.amount), 0)::text from ${donations} d where d.claim_id = ${claims.id} and d.recovered)`,
      recCount: sql<number>`(select count(*)::int from ${donations} d where d.claim_id = ${claims.id} and d.recovered)`,
    })
    .from(claims)
    .innerJoin(sql`batches b`, sql`b.id = ${claims.batchId}`)
    .innerJoin(charities, sql`${charities.id} = b.charity_id`)
    .where(
      and(
        sql`exists (select 1 from ${donations} d where d.claim_id = ${claims.id} and d.recovered)`,
        filter?.claimId ? eq(claims.id, filter.claimId) : undefined,
        filter?.charityId ? sql`b.charity_id = ${filter.charityId}` : undefined
      )
    )
  return rows.map((r) => {
    const expected = giftAidOn(Number(r.claim.totalAmount))
    const received = r.claim.receivedAmount == null ? null : Number(r.claim.receivedAmount)
    const f = feeFigures(giftAidOn(Number(r.recAmount)), expected, r.claim.status === 'received' ? received : null)
    return {
      claimId: r.claim.id,
      ref: r.claim.ref,
      charityId: r.charityId,
      charityName: r.charityName,
      status: r.claim.status,
      hmrcRef: r.claim.hmrcRef,
      createdAt: r.claim.createdAt,
      submittedAt: r.claim.submittedAt,
      receivedAt: r.claim.receivedAt,
      expected,
      received,
      recoveredCount: r.recCount,
      recoveredGiftAid: f.recovered,
      recoveredPaid: f.recoveredPaid,
      fee: f.fee,
    }
  })
}

/** Tells BoostAid's owner when a claim with recovered Gift Aid is submitted or received. */
export async function notifyClaimEvent(claimId: string, event: 'submitted' | 'received') {
  const [c] = await claimsWithRecovery({ claimId })
  if (!c) return // nothing recovered in this claim: nothing for BoostAid to be told
  if (event === 'submitted') {
    await addNotification({
      type: 'submitted',
      charityId: c.charityId,
      claimId,
      text: `${c.charityName} submitted ${c.ref} to HMRC (ref ${c.hmrcRef ?? ''}). Includes ${gbp(c.recoveredGiftAid)} recovered Gift Aid.`,
    })
  } else {
    const less = c.received != null && c.received < c.expected ? ` (${gbp(c.expected - c.received)} less than claimed)` : ''
    await addNotification({
      type: 'received',
      charityId: c.charityId,
      claimId,
      text: `${c.charityName} marked ${c.ref} received: HMRC paid ${gbp(c.received ?? c.expected)}${less}. Fee of ${gbp(c.fee)} ready to invoice.`,
    })
  }
}
