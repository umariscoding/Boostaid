/**
 * Writes HMRC Gift Aid schedules as .ods files.
 *
 * HMRC's template says: "Save the schedule as a .ods file not a Microsoft Excel .xlsx file"
 * and "Don't change the layout of the schedule or change the name of the worksheet".
 *
 * So we DON'T build a spreadsheet from scratch. We take HMRC's own template (converted once to
 * .ods with LibreOffice and saved as public/hmrc_template.ods) and only fill in cells:
 *   D13        Box 1 - earliest donation date in the claim
 *   C25:K1024  one donation per row: Title, First name, Last name, House name or number, Postcode,
 *              (Aggregated - left empty), (Sponsored - left empty), Donation date, Amount
 *   J20        the cached value of HMRC's "Total donations" formula
 * Everything else (sheet name, layout, protection, styles, instructions) stays byte-for-byte HMRC's.
 */
import JSZip from 'jszip'
import type { HmrcFilePlan } from './engine'
import { toDDMMYY } from './dates'

const ROW_RE = /<table:table-row\b[^>]*?(?:\/>|>[\s\S]*?<\/table:table-row>)/g
const CELL_RE = /<table:(?:covered-)?table-cell\b[^>]*?(?:\/>|>[\s\S]*?<\/table:(?:covered-)?table-cell>)/g
const FIRST_DATA_ROW = 25

const esc = (x: string) => x.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

type Cell = { kind: 'string'; text: string } | { kind: 'date'; iso: string } | { kind: 'float'; value: number }

function cellXml(style: string | null, c: Cell): string {
  const st = style ? ` table:style-name="${style}"` : ''
  if (c.kind === 'string') return `<table:table-cell${st} office:value-type="string" calcext:value-type="string"><text:p>${esc(c.text)}</text:p></table:table-cell>`
  if (c.kind === 'date') return `<table:table-cell${st} office:value-type="date" office:date-value="${c.iso}" calcext:value-type="date"><text:p>${toDDMMYY(c.iso)}</text:p></table:table-cell>`
  return `<table:table-cell${st} office:value-type="float" office:value="${c.value}" calcext:value-type="float"><text:p>${c.value.toFixed(2)}</text:p></table:table-cell>`
}

/** Replace cells in one <table:table-row> by column index (0 = A). Columns not given are kept as they are. */
function setCells(rowXml: string, values: Map<number, Cell>): string {
  const open = rowXml.match(/^<table:table-row\b[^>]*>/)![0]
  const cells = rowXml.match(CELL_RE) ?? []
  const maxCol = Math.max(...values.keys())
  const out: string[] = []
  let col = 0
  for (const cell of cells) {
    const rep = Number(cell.match(/table:number-columns-repeated="(\d+)"/)?.[1] ?? 1)
    if (col > maxCol) { out.push(cell); col += rep; continue }           // beyond the columns we touch: keep as-is
    const single = cell.replace(/\s*table:number-columns-repeated="\d+"/, '')
    for (let k = 0; k < rep; k++, col++) {
      if (col > maxCol) {                                                 // rest of a repeated run
        const left = rep - k
        out.push(left > 1 ? single.replace(/^<table:([\w-]+)/, `<table:$1 table:number-columns-repeated="${left}"`) : single)
        col += left
        break
      }
      const v = values.get(col)
      if (v) out.push(cellXml(single.match(/table:style-name="([^"]+)"/)?.[1] ?? null, v))
      else out.push(single)
    }
  }
  return `${open}${out.join('')}</table:table-row>`
}

/**
 * Fill one HMRC file. `templateOds` = bytes of public/hmrc_template.ods.
 * `box1Date` = earliest donation date across all files of the same kind (UK or overseas).
 */
export async function writeHmrcOds(templateOds: ArrayBuffer | Uint8Array, file: HmrcFilePlan, box1Date: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(templateOds)
  const content = await zip.file('content.xml')!.async('string')
  const total = Math.round(file.rows.reduce((a, r) => a + r.amount, 0) * 100) / 100
  let rowNo = 1
  const filled = content.replace(ROW_RE, (rowXml) => {
    const rep = Number(rowXml.match(/table:number-rows-repeated="(\d+)"/)?.[1] ?? 1)
    const here = rowNo
    rowNo += rep
    if (rep !== 1) return rowXml                                          // rows 25-1024 and 13/20 are single rows in the template
    if (here === 13) return setCells(rowXml, new Map([[3, { kind: 'date', iso: box1Date }]]))
    if (here === 20) return rowXml.replace(/(table:formula="of:=SUM\(\[\.K25:\.K1024\]\)"[^>]*office:value=")[^"]*("[^>]*><text:p>)[^<]*/, `$1${total}$2£${total.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`)
    const i = here - FIRST_DATA_ROW
    if (i < 0 || i >= file.rows.length) return rowXml
    const r = file.rows[i]
    const v = new Map<number, Cell>()
    if (r.title) v.set(2, { kind: 'string', text: r.title })              // C Title
    v.set(3, { kind: 'string', text: r.firstName })                       // D First name
    v.set(4, { kind: 'string', text: r.lastName })                        // E Last name
    v.set(5, { kind: 'string', text: r.house })                           // F House name or number
    v.set(6, { kind: 'string', text: r.postcode })                        // G Postcode
    //                                                                       H Aggregated, I Sponsored: left truly empty
    v.set(9, { kind: 'date', iso: r.date! })                              // J Donation date (real date, shown DD/MM/YY)
    v.set(10, { kind: 'float', value: r.amount })                         // K Amount (number, no £ sign)
    return setCells(rowXml, v)
  })
  zip.file('content.xml', filled)
  zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet', { compression: 'STORE' }) // must stay first + uncompressed
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.oasis.opendocument.spreadsheet' })
}
