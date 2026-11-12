/**
 * Dates. Platforms disagree: LaunchGood and Fundraise Up export US dates (MM/DD/YYYY),
 * Donations Manager and the cash book use UK dates, and Excel date cells have no order at all.
 * We detect the order PER FILE and keep every date as "YYYY-MM-DD" (no time zone surprises).
 */
import { s, DateValue } from './text'

export type DateFormat = 'DMY4' | 'DMY2' | 'MDY4' | 'MDY2'

/** UK unless the file proves it is US (a second number over 12, e.g. 09/13/2025). */
export function detectDateOrder(values: unknown[]): DateFormat {
  let us = 0, uk = 0, y4 = false
  for (const v of values) {
    const t = s(v)
    const m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/)
    if (!m) continue
    if (Number(m[1]) > 12) uk++
    if (Number(m[2]) > 12) us++
    if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(t)) y4 = true
  }
  if (us > uk) return y4 ? 'MDY4' : 'MDY2'
  return y4 ? 'DMY4' : 'DMY2'
}

function valid(y: number, m: number, d: number): string | null {
  if (m < 1 || m > 12 || d < 1) return null
  const dim = new Date(Date.UTC(y, m, 0)).getUTCDate()
  if (d > dim) return null
  return `${String(y).padStart(4, '0')}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`
}

function tryFormat(x: string, f: DateFormat | 'YMD'): string | null {
  if (f === 'YMD') {
    const m = x.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/)
    return m ? valid(+m[1], +m[2], +m[3]) : null
  }
  const four = f.endsWith('4')
  const m = x.match(four ? /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/ : /^(\d{1,2})\/(\d{1,2})\/(\d{2})$/)
  if (!m) return null
  let y = +m[3]
  if (!four) y = y < 69 ? 2000 + y : 1900 + y // same pivot as Python's %y
  return f.startsWith('D') ? valid(y, +m[2], +m[1]) : valid(y, +m[1], +m[2])
}

/** Any cell value -> "YYYY-MM-DD", or null if it can't be read. Time of day is dropped. */
export function parseDate(v: unknown, fmt: DateFormat): string | null {
  if (v instanceof DateValue) return v.iso
  const x = s(v).split(' ')[0]
  for (const f of [fmt, 'YMD', 'DMY2', 'DMY4', 'MDY4'] as const) {
    const r = tryFormat(x, f)
    if (r) return r
  }
  return null
}

/** "2025-04-18" -> "18/04/25" (HMRC's DD/MM/YY display). */
export function toDDMMYY(iso: string): string {
  return `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(2, 4)}`
}
