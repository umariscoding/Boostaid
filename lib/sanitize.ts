/**
 * Cleans text that came out of spreadsheets with broken encoding or unwanted symbols.
 *
 * 1. Repairs mojibake: UTF-8 text that was wrongly decoded as Windows-1252, e.g.
 *    "â€™" -> "’" and "ðŸ˜‡" -> "😇".
 * 2. Strips emoji / pictographs and other invisible or control characters.
 * 3. Converts accented letters to plain letters (Erdogan with a dotless g -> Erdogan, Zoe with a diaeresis -> Zoe).
 * 4. Normalises smart quotes, dashes and non-breaking spaces to plain ASCII. Hyphens and apostrophes are kept.
 */

// Windows-1252 bytes 0x80-0x9F and the characters they decode to.
const CP1252_SPECIALS: Record<number, number> = {
  0x20ac: 0x80, 0x201a: 0x82, 0x0192: 0x83, 0x201e: 0x84, 0x2026: 0x85, 0x2020: 0x86,
  0x2021: 0x87, 0x02c6: 0x88, 0x2030: 0x89, 0x0160: 0x8a, 0x2039: 0x8b, 0x0152: 0x8c,
  0x017d: 0x8e, 0x2018: 0x91, 0x2019: 0x92, 0x201c: 0x93, 0x201d: 0x94, 0x2022: 0x95,
  0x2013: 0x96, 0x2014: 0x97, 0x02dc: 0x98, 0x2122: 0x99, 0x0161: 0x9a, 0x203a: 0x9b,
  0x0153: 0x9c, 0x017e: 0x9e, 0x0178: 0x9f,
}

/** Byte value if `ch` could have come from a cp1252/latin1 decode of a UTF-8 byte. */
function toByte(ch: string): number | null {
  const code = ch.charCodeAt(0)
  if (code <= 0xff) return code
  return CP1252_SPECIALS[code] ?? null
}

// Continuation bytes are 0x80-0xBF once decoded: U+0080-U+00BF plus the cp1252 specials.
const CONT = '[\\u0080-\\u00bf\\u0152\\u0153\\u0160\\u0161\\u0178\\u017d\\u017e\\u0192\\u02c6\\u02dc\\u2013\\u2014\\u2018-\\u201a\\u201c-\\u201e\\u2020-\\u2022\\u2026\\u2030\\u2039\\u203a\\u20ac\\u2122]'
// A UTF-8 lead byte (0xC2-0xF4) followed by the right number of continuation bytes.
const MOJIBAKE_RE = new RegExp(
  `[\\u00c2-\\u00df]${CONT}|[\\u00e0-\\u00ef]${CONT}{2}|[\\u00f0-\\u00f4]${CONT}{3}`,
  'g'
)
const decoder = new TextDecoder('utf-8', { fatal: true })

export function repairMojibake(input: string): string {
  return input.replace(MOJIBAKE_RE, (match) => {
    const bytes = new Uint8Array(match.length)
    for (let i = 0; i < match.length; i++) {
      const b = toByte(match[i])
      if (b === null) return match
      bytes[i] = b
    }
    try {
      return decoder.decode(bytes)
    } catch {
      return match // Not valid UTF-8, so it was a genuine accented letter
    }
  })
}

const EMOJI_RE = new RegExp(
  '[\\p{Extended_Pictographic}\\p{Regional_Indicator}\\u{1F3FB}-\\u{1F3FF}\\u{E0020}-\\u{E007F}\\u20E3\\uFE0E\\uFE0F\\u200D]',
  'gu'
)
const INVISIBLE_RE = new RegExp(
  '[\\u0000-\\u0008\\u000B\\u000C\\u000E-\\u001F\\u007F-\\u009F\\u00AD\\u200B\\u200C\\u200E\\u200F\\u2028\\u2029\\u202A-\\u202E\\u2060\\u2066-\\u2069\\uFEFF\\uFFFD]',
  'g'
)
const SINGLE_QUOTES_RE = new RegExp('[\\u2018\\u2019\\u201A\\u2032\\u0060\\u00B4]', 'g')
const DOUBLE_QUOTES_RE = new RegExp('[\\u201C\\u201D\\u201E\\u2033]', 'g')
const DASHES_RE = new RegExp('[\\u2010-\\u2015\\u2212]', 'g')
const ELLIPSIS_RE = new RegExp('\\u2026', 'g')
const SPACES_RE = new RegExp('[\\u00A0\\u1680\\u2000-\\u200A\\u202F\\u205F\\u3000]', 'g')

// Letters that don't decompose into base letter + accent, so Unicode normalisation can't flatten them
const SPECIAL_LETTERS: [number, string][] = [
  [0x00df, 'ss'], [0x00e6, 'ae'], [0x00c6, 'AE'], [0x0153, 'oe'], [0x0152, 'OE'],
  [0x00f8, 'o'], [0x00d8, 'O'], [0x0111, 'd'], [0x0110, 'D'], [0x00f0, 'd'], [0x00d0, 'D'],
  [0x0142, 'l'], [0x0141, 'L'], [0x0131, 'i'], [0x0130, 'I'], [0x00fe, 'th'], [0x00de, 'Th'],
  [0x0127, 'h'], [0x0126, 'H'],
]
const SPECIAL_MAP = new Map(SPECIAL_LETTERS.map(([code, rep]) => [String.fromCharCode(code), rep]))
const SPECIAL_RE = new RegExp('[' + SPECIAL_LETTERS.map(([code]) => String.fromCharCode(code)).join('') + ']', 'g')

/** "Zoe" with a diaeresis -> "Zoe", "Erdogan" with a breve -> "Erdogan". Letters outside Latin (Arabic, Cyrillic...) are untouched. */
export function foldAccents(input: string): string {
  return input
    .normalize('NFD')
    .replace(/\p{M}+/gu, '')
    .replace(SPECIAL_RE, (c) => SPECIAL_MAP.get(c) ?? c)
    .normalize('NFC')
}

export function sanitizeText(input: string): string {
  let s = repairMojibake(input)
  s = foldAccents(s.replace(EMOJI_RE, '').replace(INVISIBLE_RE, ''))
  s = s
    .replace(SINGLE_QUOTES_RE, "'")
    .replace(DOUBLE_QUOTES_RE, '"')
    .replace(DASHES_RE, '-')
    .replace(ELLIPSIS_RE, '...')
    .replace(SPACES_RE, ' ')
  return s.replace(/[ \t]{2,}/g, ' ').trim()
}

/** Sanitises every string value of every record. Returns the count of cells that changed. */
export function sanitizeRecords<T extends Record<string, unknown>>(
  records: T[]
): { records: T[]; cellsCleaned: number } {
  let cellsCleaned = 0
  const cleaned = records.map((rec) => {
    const out: Record<string, unknown> = {}
    for (const [key, value] of Object.entries(rec)) {
      if (typeof value === 'string') {
        const next = sanitizeText(value)
        if (next !== value.trim()) cellsCleaned++
        out[key] = next
      } else {
        out[key] = value
      }
    }
    return out as T
  })
  return { records: cleaned, cellsCleaned }
}
