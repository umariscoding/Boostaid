/**
 * HMRC's schedule has a single "House name or number" column. This reduces a full address to just that part:
 *
 *   7 Wordsworth Rd                      -> 7
 *   302 hadding lodge, 9 knapp road      -> 302 Hadding Lodge
 *   Flat 2, 25 Waylen Street             -> Flat 2
 *   Rose Cottage, Mill Lane              -> Rose Cottage
 *   25                                   -> 25            (a bare number is correct)
 *   Shelley Ave                          -> null          (no number or house name: needs review)
 *
 * Maximum 40 characters.
 */

export const MAX_HOUSE_LENGTH = 40

const STREET_WORDS = new Set(
  (
    'road rd street st avenue ave av lane ln drive dr close cl way crescent cres terrace ter terr grove gr ' +
    'place pl square sq walk row mews rise hill view parade gate green park path passage alley bank ' +
    'boulevard blvd broadway circus embankment esplanade highway hwy lawn promenade quay ridge ' +
    'gardens garden gdns approach estate parkway wharf meadow meadows fields field common heath orchard chase ' +
    'croft dene vale hollow lea mead paddock yard end side'
  ).split(' ')
)

// Words that mean the text is a building name, not a street, e.g. "302 Hadding Lodge"
const BUILDING_WORDS = new Set(
  (
    'lodge house court ct mansions mansion towers tower building buildings apartments apartment heights villas villa ' +
    'cottage cottages manor block centre center hall flats farm barn mill chambers residence plaza point studios'
  ).split(' ')
)

const DIRECTIONS = new Set(['north', 'south', 'east', 'west', 'n', 's', 'e', 'w', 'ne', 'nw', 'se', 'sw'])

const FLAT_RE = /^(flat|apartment|apt|unit|room|suite|studio|annexe|annex)\.?\s+([A-Za-z]?\d+[A-Za-z]?|[A-Za-z])(?![A-Za-z0-9])/i
const NUMBER_RE = /^(\d+[A-Za-z]?(?:\s*[-/]\s*\d+[A-Za-z]?)?)(?![A-Za-z0-9])\s*(.*)$/

const clean = (w: string) => w.replace(/[.,]+$/g, '').toLowerCase()
const words = (s: string) => s.split(/\s+/).filter(Boolean).map(clean)

/** Last word, ignoring a trailing compass direction ("High Road North" -> "road"). */
function lastWord(s: string): string {
  const w = words(s)
  while (w.length > 1 && DIRECTIONS.has(w[w.length - 1])) w.pop()
  return w[w.length - 1] ?? ''
}

const isBuilding = (s: string) => BUILDING_WORDS.has(lastWord(s))

/**
 * Looks like a street: ends in a street word ("Wordsworth Rd"), ends in "...way" ("Westway"), or has one after the
 * first word ("Ley Street Ilford"). A building word at the end wins ("Hadding Lodge", "Park House").
 */
function isStreet(s: string): boolean {
  if (isBuilding(s)) return false
  const w = words(s)
  const last = lastWord(s)
  if (STREET_WORDS.has(last) || (last.length > 4 && last.endsWith('way'))) return true
  return w.slice(1).some((x) => STREET_WORDS.has(x))
}

/** "hadding lodge" -> "Hadding Lodge", "o'brien" -> "O'Brien", "17a" -> "17A". Mixed-case input is left alone. */
export function tidyCase(s: string): string {
  const letters = s.replace(/[^A-Za-z]/g, '')
  const isMixed = letters !== letters.toLowerCase() && letters !== letters.toUpperCase()
  if (isMixed) return s
  return s
    .toLowerCase()
    .split(' ')
    .map((word) => {
      if (/^\d+[a-z]$/.test(word)) return word.toUpperCase()
      return word.replace(/(^|[-/])([a-z])|'([a-z])/g, (m, pre: string, c1: string, c2: string, offset: number) => {
        if (c2) {
          // O'Brien / D'Arcy capitalise after a single-letter prefix; King's stays lower case
          const before = word.slice(0, offset).replace(/^.*[-/]/, '')
          return before.length === 1 ? `'${c2.toUpperCase()}` : m
        }
        return `${pre}${c1.toUpperCase()}`
      })
    })
    .join(' ')
}

export interface HouseResult {
  /** The house name/number to send, or null when the row needs review. */
  value: string | null
  /** True when the text had to be cut to fit 40 characters. */
  truncated: boolean
}

function fit(value: string): HouseResult {
  if (value.length <= MAX_HOUSE_LENGTH) return { value, truncated: false }
  const cut = value.slice(0, MAX_HOUSE_LENGTH + 1)
  const atWord = cut.lastIndexOf(' ')
  const trimmed = (atWord > 0 ? cut.slice(0, atWord) : cut.slice(0, MAX_HOUSE_LENGTH)).replace(/[\s,.-]+$/, '')
  return { value: trimmed, truncated: true }
}

export function extractHouse(raw: unknown): HouseResult {
  const text = String(raw ?? '')
    .replace(/[\r\n]+/g, ', ')
    .replace(/\s+/g, ' ')
    .trim()
  if (!text) return { value: null, truncated: false }

  const parts = text.split(/\s*,\s*/).filter(Boolean)
  const first = (parts[0] ?? '').replace(/^[\s.,;-]+|[\s.,;]+$/g, '')
  if (!first) return { value: null, truncated: false }

  // "Flat 2, 25 Waylen Street" -> "Flat 2"
  const flat = FLAT_RE.exec(first)
  if (flat) {
    const id = flat[2].length <= 3 && /\d/.test(flat[2]) ? flat[2].toUpperCase() : flat[2]
    const label = flat[1].charAt(0).toUpperCase() + flat[1].slice(1).toLowerCase()
    return fit(`${label} ${id}`)
  }

  // Starts with a number
  const num = NUMBER_RE.exec(first)
  if (num) {
    const number = num[1].replace(/\s*([-/])\s*/g, '$1').replace(/[a-z]$/, (c) => c.toUpperCase())
    const rest = num[2].trim()
    if (!rest) return fit(number) // a bare number is correct

    // "302 hadding lodge, 9 knapp road": a building name is part of the house, a street is not
    if (isStreet(rest)) return fit(number)
    if (isBuilding(rest)) return fit(tidyCase(first))
    // Unrecognised words: with a comma after them they're a building name, otherwise assume a street
    return parts.length > 1 ? fit(tidyCase(first)) : fit(number)
  }

  // No number: either a house name ("Rose Cottage") or just a street ("Shelley Ave")
  if (isStreet(first)) return { value: null, truncated: false }
  return fit(tidyCase(first))
}
