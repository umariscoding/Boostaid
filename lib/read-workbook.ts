import * as XLSX from 'xlsx'

const BOM = String.fromCharCode(0xfeff)

/**
 * Reads an uploaded spreadsheet. CSV files are decoded as UTF-8 explicitly (SheetJS otherwise guesses a legacy
 * code page, which is how "O'Brien" with a curly apostrophe turns into garbage). If a CSV isn't valid UTF-8
 * it's treated as Windows-1252, which is what older Excel versions export.
 */
export async function readWorkbook(file: File): Promise<XLSX.WorkBook> {
  const buffer = await file.arrayBuffer()

  if (/\.csv$/i.test(file.name)) {
    let text: string
    try {
      text = new TextDecoder('utf-8', { fatal: true }).decode(buffer)
    } catch {
      text = new TextDecoder('windows-1252').decode(buffer)
    }
    if (text.startsWith(BOM)) text = text.slice(1)
    return XLSX.read(text, { type: 'string' })
  }

  return XLSX.read(buffer, { type: 'array', sheetRows: 0 })
}
