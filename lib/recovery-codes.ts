/** Secret link codes. Server only (uses node:crypto). */
import { createHash, randomInt } from 'node:crypto'

const ALPHABET = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789' // no 0/O/1/I: easy to read, still 32^10 possibilities

/** 10 characters from a cryptographically secure source. Only the hash is stored. */
export function newCode(len = 10): string {
  let c = ''
  for (let i = 0; i < len; i++) c += ALPHABET[randomInt(ALPHABET.length)]
  return c
}
export const hashCode = (code: string) => createHash('sha256').update(code.trim().toUpperCase()).digest('hex')
