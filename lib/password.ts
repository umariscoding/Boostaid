import { randomBytes, scrypt, timingSafeEqual, type ScryptOptions } from 'node:crypto'

// scrypt, memory-hard. Stored as: scrypt$N$r$p$<salt base64>$<hash base64>
const N = 32768
const R = 8
const P = 1
const KEYLEN = 64

function derive(password: string, salt: Buffer, n: number, r: number, p: number): Promise<Buffer> {
  const opts: ScryptOptions = { N: n, r, p, maxmem: 128 * n * r * 2 }
  return new Promise((resolve, reject) =>
    scrypt(password.normalize('NFKC'), salt, KEYLEN, opts, (err, key) => (err ? reject(err) : resolve(key)))
  )
}

export async function hashPassword(password: string): Promise<string> {
  const salt = randomBytes(16)
  const key = await derive(password, salt, N, R, P)
  return `scrypt$${N}$${R}$${P}$${salt.toString('base64')}$${key.toString('base64')}`
}

export async function verifyPassword(password: string, stored: string): Promise<boolean> {
  const [scheme, n, r, p, saltB64, hashB64] = stored.split('$')
  if (scheme !== 'scrypt' || !saltB64 || !hashB64) return false
  const expected = Buffer.from(hashB64, 'base64')
  const key = await derive(password, Buffer.from(saltB64, 'base64'), Number(n), Number(r), Number(p))
  return key.length === expected.length && timingSafeEqual(key, expected)
}

let dummy: Promise<string> | null = null
/** Runs a real hash comparison when the email is unknown, so timing doesn't reveal which emails exist. */
export async function burnPasswordCheck(password: string): Promise<void> {
  dummy ??= hashPassword('not-a-real-password')
  await verifyPassword(password, await dummy)
}

/** Returns a message if the password is too weak, otherwise null. */
export function passwordProblem(password: string, email?: string): string | null {
  if (password.length < 10) return 'Use at least 10 characters.'
  if (password.length > 200) return 'That password is too long.'
  if (email && password.toLowerCase() === email.toLowerCase()) return 'The password can\'t be your email address.'
  if (/^(.)\1+$/.test(password)) return 'Choose a password that is not one repeated character.'
  return null
}

/** A strong random password for first-time accounts (shown once, then changed by the user). */
export function generatePassword(): string {
  return randomBytes(15).toString('base64url')
}
