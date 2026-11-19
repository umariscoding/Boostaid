/**
 * HMRC "House name or number" rule.
 *
 * HMRC guidance: "Enter the donor's house name or number using no more than 40 characters.
 * If a donor lives in the UK, do not enter a street, road or area name."
 *
 * Platforms give us 1-4 messy address lines, often street first and number second
 * ("Galsworthy Avenue, 17 Donoghue Cottages"). We look at EVERY part of EVERY line,
 * score each candidate, and keep the best. Every case in spot_checks.json must pass.
 */
import { s, toAscii, tidyWords, formatPostcode, stripChars, UK_POSTCODE } from './text'

// Words that end a street name. (A few common typos from real data are included on purpose.)
const STREET = '(road|rd|street|st|avenue|ave|lane|ln|close|cl|gardens?|gdns|grove|gr|drive|dr|way|place|pl|crescent|cres|' +
  'terrace|terr|walk|square|sq|hill|park|rise|row|mews|green|grn|parade|vale|view|fields|highway|approach|' +
  'boulevard|blvd|broadway|embankment|path|wynd|hey|croft|chase|end|gate|yard|market|steeet|gardend)\\.?'
// Words that end a building name ("Hollybush House", "Donoghue Cottages").
const BUILDING = '(house|lodge|cottages?|mansions|apartments|building|buildings|tower|hall|farm|barn|villa|manor|heights|court|ct|point|place|block|wharf|quay|centre|works|mill)'
// Stricter list used when a building name stands on its own (no number).
const NAMED_BUILDING = '(house|lodge|cottages?|mansions|apartments|building|tower|hall|farm|barn|villa|manor|heights|point|court)'

const re = (src: string) => new RegExp(src, 'i')
const ENDS_BUILDING = re(`\\b${BUILDING}$`)
const ENDS_STREET = re(`\\b${STREET}$`)
const HAS_STREET = re(`\\b${STREET}\\b`)
const ENDS_NAMED_BUILDING = re(`\\b${NAMED_BUILDING}$`)
const FLAT_THEN_BUILDING = re(`\\b((?:flat|apartment|apt|unit)\\s*\\w+\\s+.*\\b${BUILDING})$`)
const BUILDING_THEN_STREET = re(`^(.*?\\b${NAMED_BUILDING})\\b\\s+.*\\b${STREET}$`)
const UP_TO_BUILDING = re(`^(.*?\\b${NAMED_BUILDING})\\b`)
const JOINABLE = /^(\d+[A-Z]?(?:-\d+[A-Z]?)?|(?:Flat|Apartment|Apt|Unit|Room) ?\w+|[A-Z]{1,2}\d{1,3})$/i

export type HouseResult = { house: string; reason: null } | { house: null; reason: string }

export function houseFrom(lines: unknown[]): HouseResult {
  // 1. Split every line on commas: "Flat 10, Imperial House" -> ["Flat 10", "Imperial House"]
  let parts: string[] = []
  for (const line of lines) {
    for (const p of toAscii(s(line)).split(',')) {
      const t = stripChars(p, ' .')
      if (t) parts.push(t)
    }
  }
  // 2. Ignore postcodes typed into an address line ("London", "E14 0JB")
  parts = parts.filter((p) => !UK_POSTCODE.test(formatPostcode(p)) || (!p.trim().includes(' ') && p.trim().length < 5))

  // 3. Score each part; the highest score wins (ties keep the first)
  let best: { text: string; score: number } | null = null
  for (const raw of parts) {
    const p = raw.replace(/\s+/g, ' ')
    let cand: string | null = null
    let score = 0

    // Flat codes like "G03" only count when followed by nothing or by a building name
    const code = p.match(/^[a-z]{1,2}\d{1,3}\b(.*)$/i)
    if (code && code[1].trim() && !ENDS_BUILDING.test(stripChars(code[1].trim(), ' ,.'))) continue

    const mFlat = p.match(/^((?:flat|apartment|apt|unit|room)\s*[a-z0-9]+|[a-z]{1,2}\d{1,3})\b(.*)$/i)
    const mNum = p.match(/^(\d+[a-z]?(?:-\d+[a-z]?)?)\b\s*(.*)$/i)
    if (mFlat) {
      const flat = mFlat[1]
      const rest = stripChars(mFlat[2].trim(), ' ,.')
      if (rest && ENDS_BUILDING.test(rest) && !/^\d/.test(rest)) { cand = `${flat} ${rest}`; score = 4 }      // "Flat 13 Frobisher House"
      else if (/\d/.test(flat)) { cand = flat; score = 3 }                                                    // "Flat 26"
    } else if (mNum) {
      const n = mNum[1]
      const rest = stripChars(mNum[2].trim(), ' ,.')
      const flatBuilding = rest.match(FLAT_THEN_BUILDING)
      const buildingStreet = rest.match(BUILDING_THEN_STREET)
      if (!rest) { cand = n; score = 2 }                                                                    // "27"
      else if (ENDS_BUILDING.test(rest) && !HAS_STREET.test(rest)) { cand = `${n} ${rest}`; score = 4 }      // "79 Hollybush House"
      else if (flatBuilding) { cand = flatBuilding[1]; score = 4 }                                         // "180C York Way Flat 405 Darley Building"
      else if (buildingStreet && !HAS_STREET.test(rest.match(UP_TO_BUILDING)![1])) {
        cand = `${n} ${rest.match(UP_TO_BUILDING)![1]}`; score = 4                                         // "13 Foundry Court Tredegar Road"
      }
      else if (ENDS_STREET.test(rest)) { cand = n; score = 2 }                                              // "9 Brokesley Street" -> "9"
      else { cand = n; score = 1.5 }
    } else if (/\d/.test(p)) {
      const m = p.match(/\b(\d+[a-z]?)\b/i)
      if (m) { cand = m[1]; score = 1 }
    } else if (ENDS_NAMED_BUILDING.test(p) && !ENDS_STREET.test(p)) {
      cand = p; score = 1.2                                                                                 // "Hector House"
    }
    if (cand && (best === null || score > best.score)) best = { text: tidyWords(cand), score }
  }

  // 4. A bare number or lone flat + a building name elsewhere -> join them ("7" + "Keats House")
  if (best && (best.score === 2 || best.score === 3) && JOINABLE.test(best.text)) {
    for (const raw of parts) {
      const p = raw.trim()
      if (!/\d/.test(p) && ENDS_NAMED_BUILDING.test(p) && !ENDS_STREET.test(p)) {
        const comb = `${best.text} ${tidyWords(p)}`
        if (comb.length <= 40) best = { text: comb, score: best.score }
        break
      }
    }
  }

  if (!best) return { house: null, reason: parts.length ? 'No house number or name' : 'No address' }
  const h = best.text.replace(/\./g, '').trim()
  if (h.length > 40) return { house: null, reason: 'House name/number over 40 characters' }
  return { house: h, reason: null }
}
