/**
 * Reading uploaded files (xlsx / xls / csv) into plain tables. Works in the browser and in Node.
 *
 * - CSVs are read as UTF-8; if that fails (older Windows exports) as Windows-1252. This is what
 *   stops "â€™" / "ðŸ˜‡" appearing.
 * - CSV values are kept as TEXT (no automatic number/date guessing - dates are handled per file).
 * - Excel date cells become DateValue("YYYY-MM-DD") without time-zone conversion.
 */
import * as XLSX from 'xlsx'
import { DateValue } from './text'

export interface Table {
  fileName: string
  sheetName: string           // "CSV" for csv files
  rows: unknown[][]           // rows[0] = headings
}

const UTF8 = new TextDecoder('utf-8', { fatal: true })
const CP1252 = new TextDecoder('windows-1252')

export function readTables(fileName: string, data: ArrayBuffer | Uint8Array): Table[] {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data)
  if (fileName.toLowerCase().endsWith('.csv')) {
    let text: string
    try { text = UTF8.decode(bytes) } catch { text = CP1252.decode(bytes) }
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1) // byte-order mark
    const wb = XLSX.read(text, { type: 'string', raw: true })
    return [{ fileName, sheetName: 'CSV', rows: sheetRows(wb.Sheets[wb.SheetNames[0]]) }]
  }
  const wb = XLSX.read(bytes, { type: 'array', cellDates: false, cellNF: true }) // cellNF: keep number formats so date cells can be recognised
  return wb.SheetNames.map((name) => ({ fileName, sheetName: name, rows: sheetRows(wb.Sheets[name]) }))
}

/** Every row of a sheet as an array, full width, with date cells converted. */
function sheetRows(ws: XLSX.WorkSheet | undefined): unknown[][] {
  if (!ws || !ws['!ref']) return []
  const range = XLSX.utils.decode_range(ws['!ref'])
  const rows: unknown[][] = []
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row: unknown[] = []
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined
      row.push(cellValue(cell))
    }
    rows.push(row)
  }
  return rows
}

function cellValue(cell: XLSX.CellObject | undefined): unknown {
  if (!cell || cell.v === undefined || cell.v === null) return null
  if (cell.t === 'n' && cell.z && XLSX.SSF.is_date(cell.z as string)) {
    const d = XLSX.SSF.parse_date_code(cell.v as number)
    if (d) return new DateValue(`${String(d.y).padStart(4, '0')}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`)
  }
  if (cell.t === 'd' && cell.v instanceof Date) {
    const v = cell.v
    return new DateValue(`${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`)
  }
  if (cell.t === 'e') return null
  return cell.v
}

/** SHA-256 fingerprint of an uploaded file (proves the original wasn't changed). */
export async function sha256Hex(data: ArrayBuffer | Uint8Array): Promise<string> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data)
  const digest = await globalThis.crypto.subtle.digest('SHA-256', bytes as Uint8Array<ArrayBuffer>)
  return Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, '0')).join('')
}
