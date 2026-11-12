/**
 * BoostAid Gift Aid engine - public entry point.
 *
 *   import { buildClaimPack } from '@/lib/giftaid'
 *   const pack = await buildClaimPack(uploadedFiles, templateOdsBytes)
 *   // pack.result  -> numbers + rows for the screens (Ready / Needs Review / Not Eligible)
 *   // pack.zip     -> the claim pack to download
 *
 * Files:
 *   text.ts      text/postcode helpers (mojibake repair, capitalisation, UK postcode format)
 *   address.ts   HMRC "House name or number" rule (tested by spot_checks.json)
 *   dates.ts     per-file US/UK date detection
 *   profiles.ts  one profile per platform (LaunchGood, Raisely, CAF, MuslimGiving, Fundraise Up, Donations Manager, cash book)
 *   read.ts      xlsx/csv -> tables, SHA-256 fingerprints
 *   engine.ts    eligibility + cleaning + planning the HMRC files
 *   ods.ts       fills HMRC's own template (.ods)
 *   workbook.ts  the claim workbook (Summary, Audit detail, Ready, Needs Review, Not Eligible, Changes Log)
 *   summary.ts   the BoostAid summary / audit detail page
 *   pack.ts      puts it all in one zip
 */
export { buildClaim, JustGivingError, MAX_ROWS_PER_HMRC_FILE } from './engine'
export type { ClaimResult, ClaimRow, ReviewRow, IneligibleRow, Change, HmrcFilePlan, SheetNote, BuildOptions } from './engine'
export { buildClaimPack } from './pack'
export type { UploadedFile, ClaimPack } from './pack'
export { readTables, sha256Hex } from './read'
export { houseFrom } from './address'
export { PROFILES, detectProfile } from './profiles'
export { writeHmrcOds } from './ods'
export { buildClaimWorkbook } from './workbook'
export { buildSummaryHtml, summaryLines } from './summary'
