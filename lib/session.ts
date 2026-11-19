import { createHmac, timingSafeEqual } from 'node:crypto'

export const SESSION_COOKIE = 'boostaid_session'
export const SESSION_MAX_AGE = 60 * 60 * 24 * 7 // 7 days, in seconds

function secret(): string {
  const s = process.env.SESSION_SECRET
  if (!s || s.length < 16) {
    throw new Error('SESSION_SECRET must be set (16+ characters)')
  }
  return s
}

function sign(payload: string): string {
  return createHmac('sha256', secret()).update(payload).digest('base64url')
}

/** Cookie value: `<user id>.<expiry-unix-seconds>.<hmac>`. Cannot be forged without SESSION_SECRET. */
export function createSessionToken(userId: string): string {
  const exp = Math.floor(Date.now() / 1000) + SESSION_MAX_AGE
  const payload = `${userId}.${exp}`
  return `${payload}.${sign(payload)}`
}

/** The signed-in user's id if the cookie is genuine and not expired, otherwise null. */
export function readSession(token: string | undefined): { userId: string } | null {
  if (!token) return null
  const parts = token.split('.')
  if (parts.length !== 3) return null
  const [userId, expRaw, sig] = parts
  if (!userId || !expRaw || !sig) return null

  let expected: string
  try {
    expected = sign(`${userId}.${expRaw}`)
  } catch {
    return null
  }
  const a = Buffer.from(sig)
  const b = Buffer.from(expected)
  if (a.length !== b.length || !timingSafeEqual(a, b)) return null

  const exp = Number(expRaw)
  return Number.isFinite(exp) && exp > Math.floor(Date.now() / 1000) ? { userId } : null
}

export function verifySessionToken(token: string | undefined): boolean {
  return readSession(token) !== null
}
