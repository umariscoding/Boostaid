/**
 * BoostAid's fee: 5% of the Gift Aid it recovered, and only once HMRC has paid the claim.
 *
 * HMRC can pay a little less than claimed. The recovered part is then reduced by the same share
 * (recovered x actual received / expected, to the penny) and the fee is 5% of that. If HMRC pays more,
 * the fee stays at 5% of what was recovered: it never goes above what was expected.
 *
 * Example: expected £4,989.00, recovered £178.75, HMRC paid £4,980.00 -> £178.43 paid -> fee £8.92.
 */
export const FEE_RATE = 0.05
const r2 = (n: number) => Math.round(n * 100) / 100

export interface FeeFigures {
  /** Gift Aid recovered by BoostAid in this claim (as claimed). */
  recovered: number
  /** The part of it HMRC actually paid (equals `recovered` until the claim is received, or if HMRC paid in full or more). */
  recoveredPaid: number
  /** 5% of `recoveredPaid`, rounded to the penny. */
  fee: number
  /** What the charity would have had without BoostAid: what HMRC paid less the recovered part. */
  withoutBoostAid: number | null
}

/**
 * `recovered` = Gift Aid on the recovered donations (25%), `expected` = Gift Aid expected on the whole claim,
 * `received` = what HMRC actually paid, or null if it hasn't been marked received.
 */
export function feeFigures(recovered: number, expected: number, received: number | null): FeeFigures {
  const share = received != null && expected > 0 ? Math.min(1, received / expected) : 1
  const recoveredPaid = r2(recovered * share)
  return {
    recovered: r2(recovered),
    recoveredPaid,
    fee: r2(recoveredPaid * FEE_RATE),
    withoutBoostAid: received != null ? r2(received - recoveredPaid) : null,
  }
}
