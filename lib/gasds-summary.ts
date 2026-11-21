import type { GasdsClaimDto } from '@/lib/gasds-db'
import { p, taxYearLabel, topUpOf } from '@/lib/gasds'

const gbp = (n: number) => '£' + n.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
const day = (iso: string | null) => (iso ? new Date(iso).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/London' }) : '')

/** The lines of the evidence summary card. One list, so the page, the PDF and the download always agree. */
export function evidenceSummary(c: GasdsClaimDto, charity: string): [string, string][] {
  const needed = Math.ceil(p(c.amount) / 10) / 100
  const enough = p(c.giftAidFigure) >= p(needed)
  return [
    ['Charity', charity],
    ['Tax year', taxYearLabel(c.taxYear)],
    ['Sent to HMRC', `${day(c.sentAt)}${c.sentBy ? ` by ${c.sentBy}` : ''}`],
    ['HMRC reference', c.hmrcRef],
    ['Small donations claimed', gbp(c.amount)],
    ['HMRC pays (25%)', gbp(topUpOf(c.amount))],
    ['Status', c.status === 'received' ? `Received ${day(c.receivedAt)}${c.receivedAmount != null ? ` · ${gbp(c.receivedAmount)}` : ''}` : 'Waiting for HMRC'],
    ['Gift Aid needed (1/10 of the claim)', gbp(needed)],
    [`Gift Aid donations claimed for ${taxYearLabel(c.taxYear)}`, `${gbp(c.giftAidFigure)} ${enough ? '✓' : '✗'}`],
  ]
}

export const RULES_TEXT = 'Rules applied: cash or card only; £30 or less from one person; no Gift Aid declaration; banked in the charity’s UK account; £50 notes taken off; cancelled receipts skipped.'

