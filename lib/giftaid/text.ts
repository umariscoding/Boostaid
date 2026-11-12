/**
 * Text helpers shared by every cleaning rule.
 *
 * These mirror the reference implementation (boostaid_clean.py) exactly, so the app
 * produces the same results as the expected outputs in the handover pack.
 */

/** Any cell value -> trimmed string ('' for null/undefined). */
export function s(v: unknown): string {
  if (v === null || v === undefined) return ''
  if (v instanceof DateValue) return v.iso
  return String(v).trim()
}

/** A date read from a spreadsheet date cell (kept as YYYY-MM-DD so time zones can never shift it). */
export class DateValue {
  constructor(public readonly iso: string) {}
  toString() { return this.iso }
}

const YES = new Set(['yes', 'y', 'true', '1', 'tax effective'])

/** Platforms mark Gift Aid as yes / Y / TRUE / 1 / "Tax Effective" - accept them all. */
export function yes(v: unknown): boolean {
  return v === true || YES.has(s(v).toLowerCase())
}

/** Number from a cell, 0 if it isn't one. */
export function num(v: unknown): number {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0
  const t = s(v)
  if (t === '') return 0
  const n = Number(t)
  return Number.isFinite(n) ? n : 0
}

// Windows-1252 bytes 0x80-0x9F (the rest of 0x00-0xFF map straight to U+0000-U+00FF).
const CP1252_HIGH: Record<string, number> = {
  '€': 0x80, '‚': 0x82, 'ƒ': 0x83, '„': 0x84, '…': 0x85, '†': 0x86, '‡': 0x87,
  'ˆ': 0x88, '‰': 0x89, 'Š': 0x8a, '‹': 0x8b, 'Œ': 0x8c, 'Ž': 0x8e, '‘': 0x91,
  '’': 0x92, '“': 0x93, '”': 0x94, '•': 0x95, '–': 0x96, '—': 0x97, '˜': 0x98,
  '™': 0x99, 'š': 0x9a, '›': 0x9b, 'œ': 0x9c, 'ž': 0x9e, 'Ÿ': 0x9f,
}
const UTF8 = new TextDecoder('utf-8', { fatal: true })

/**
 * Repairs "mojibake": UTF-8 text that was wrongly read as Windows-1252.
 *   "St Johnâ€™s"  -> "St John’s"     "ðŸ˜‡" -> "😇"     "ErdoÄŸan" -> "Erdoğan"
 * If the text isn't mojibake (or can't be repaired) it is returned unchanged.
 */
export function demojibake(x: string): string {
  if (!/[\u0080-ÿŒ-™]/.test(x)) return x // pure ASCII etc - nothing to do
  const bytes: number[] = []
  for (const ch of x) {
    const c = ch.codePointAt(0)!
    if (c < 0x80 || (c >= 0xa0 && c <= 0xff)) bytes.push(c)
    else if (CP1252_HIGH[ch] !== undefined) bytes.push(CP1252_HIGH[ch])
    else return x // character that can't be Windows-1252 -> not mojibake
  }
  try { return UTF8.decode(new Uint8Array(bytes)) } catch { return x }
}

/** Fix mojibake, turn curly quotes into ', strip accents and drop emoji/other non-ASCII. */
export function toAscii(x: string): string {
  const fixed = demojibake(x).replace(/[’‘`]/g, "'")
  return fixed.normalize('NFKD').replace(/[^\x00-\x7F]/g, '')
}

/** True if the text contains letters from a non-Latin script (e.g. Arabic) - HMRC needs English letters. */
export function nonLatin(x: string): boolean {
  for (const ch of x) if (/\p{L}/u.test(ch) && !/\p{Script=Latin}/u.test(ch)) return true
  return false
}

/** Python-style str.capitalize(): first letter upper, rest lower. */
export function capitalize(w: string): string {
  return w ? w[0].toUpperCase() + w.slice(1).toLowerCase() : w
}

const isLower = (w: string) => w !== w.toUpperCase() && w === w.toLowerCase()
const isUpper = (w: string) => w !== w.toLowerCase() && w === w.toUpperCase()

/** Capitalise names typed all in lower or upper case: "rahman" -> "Rahman", "MCDONALD" -> "McDonald", "o'brien" -> "O'Brien". */
export function cap(x: string): string {
  if (x.length > 1 && (x === x.toLowerCase() || x === x.toUpperCase())) {
    return x.toLowerCase().split(/(\s+|-)/).map((w) => {
      if (/^mc[a-z]/.test(w)) return 'Mc' + capitalize(w.slice(2))
      if (/^o'[a-z]/.test(w)) return "O'" + capitalize(w.slice(2))
      return capitalize(w)
    }).join('')
  }
  return x
}

/** Tidy an address fragment word by word: "2b" -> "2B", "gordon mansions" -> "Gordon Mansions". */
export function tidyWords(x: string): string {
  return x.split(/\s+/).filter(Boolean).map((w) => {
    if (/^\d+[a-z]$/i.test(w)) return w.toUpperCase()
    if ((isLower(w) || isUpper(w)) && !/^\d+$/.test(w)) return capitalize(w)
    return w
  }).join(' ')
}

/** Full UK postcode format (HMRC: upper case, with a space). */
export const UK_POSTCODE = /^(GIR 0AA|[A-PR-UWYZ][A-HK-Y]?[0-9][0-9A-Z]? [0-9][ABD-HJLNP-UW-Z]{2})$/

/** "ig27ea" -> "IG2 7EA". Doesn't validate - use UK_POSTCODE for that. */
export function formatPostcode(p: unknown): string {
  let t = toAscii(s(p)).replace(/[^A-Za-z0-9]/g, '').toUpperCase()
  if (t.length >= 5 && t.length <= 7) t = t.slice(0, -3) + ' ' + t.slice(-3)
  return t
}

/** Python's "%g" for amounts shown in review reasons (10.8 -> "10.8", 10 -> "10"). */
export function fmtG(n: number): string {
  return String(Number(n.toPrecision(6)))
}

/** Python-style strip of specific characters from both ends. */
export function stripChars(x: string, chars: string): string {
  let a = 0, b = x.length
  while (a < b && chars.includes(x[a])) a++
  while (b > a && chars.includes(x[b - 1])) b--
  return x.slice(a, b)
}
