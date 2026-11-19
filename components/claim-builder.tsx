'use client'

import { useEffect, useRef, useState } from 'react'
import Link from 'next/link'
import { useSearchParams } from 'next/navigation'
import { AlertCircle, ChevronRight, Columns, Download, Eye, FileSpreadsheet, FileText, FileUp, Plus, RotateCcw } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Card, PageHeader, Pill, StatTile, Tabs, td, tdr, th, thr } from '@/components/ui/primitives'
import { Pagination } from '@/components/ui/pagination'
import { ReviewSkeleton, Skeleton } from '@/components/ui/skeleton'
import { Select } from '@/components/ui/select'
import NoCharityNotice from '@/components/charity/no-charity-notice'
import SubmitDialog from '@/components/claims/submit-dialog'
import ColumnMapper, { type ColumnMapping, type GiftAidConfig } from '@/components/column-mapper'
import { useCharity } from '@/lib/store/charity'
import { StatusBadge, apiJson, fmtPlainDate, gbp, type Status } from '@/components/claims/shared'
import type { ClaimDto } from '@/lib/claims'
import { giftAidOn } from '@/lib/claims'
import { JustGivingError, PROFILES, readTables, type ClaimResult } from '@/lib/giftaid'
import type { Table } from '@/lib/giftaid/read'
import { mapToCashBook, tableKey, tableRecords, withoutAlreadyClaimed, type AlreadyClaimed, type Upload } from '@/lib/claim-flow'
import { assemblePack, buildSchedules, scheduleName, type NamedSchedule } from '@/lib/claim-pack'
import { buildSimpleSheets, simpleSheetName } from '@/lib/claim-extras'
import OptionalChecks from '@/components/claim-optional-checks'
import RecoveryPanel from '@/components/recovery/recovery-panel'
import { heldBackRows, rowKey } from '@/lib/recovery'
import { CASH_BOOK_HEADINGS } from '@/lib/claim-flow'
import type { RecoveryRowDto } from '@/app/api/recovery/route'

type Phase = 'idle' | 'reading' | 'review' | 'saving' | 'saved'

interface BlockedRow {
  name: string
  amount: number
  platform: string
  donationId: string
  claimRef: string
}

interface SavedClaim {
  id: string
  ref: string
  kind: string
  fileName: string
  donationCount: number
  totalAmount: number
}

interface SavedState {
  charityName: string
  claims: SavedClaim[]
  possibleRepeats: { count: number; claimRefs: string[] }
  schedules: NamedSchedule[]
  simpleSheets: { name: string; data: Uint8Array }[]
  zip: Uint8Array
  workbook: Uint8Array
  summaryHtml: string
}

const ODS = 'application/vnd.oasis.opendocument.spreadsheet'
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const tick = (ms: number) => new Promise((r) => setTimeout(r, ms))

/**
 * Donations the donor fixed ("Ready – recovered") join a claim as one more file, laid out like HMRC's cash book, so the
 * normal cleaning rules and schedule building apply. Row N of this file is `pool[N - 2]`.
 */
const RECOVERED_FILE = 'Recovered_by_BoostAid.csv'
const csvCell = (v: unknown) => { const t = String(v ?? ''); return /[",\n]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t }
function recoveredUpload(pool: RecoveryRowDto[]): Upload {
  const lines = [CASH_BOOK_HEADINGS.join(',')]
  for (const r of pool) {
    const f = r.fixed!
    lines.push([f.title, f.first, f.last, f.house, f.postcode, r.date ?? '', r.amount.toFixed(2)].map(csvCell).join(','))
  }
  return { name: RECOVERED_FILE, data: new TextEncoder().encode(lines.join('\n')) }
}
const isRecoveredSource = (source: string) => source.startsWith(`${RECOVERED_FILE}:CSV!`)

function download(name: string, data: BlobPart, type: string) {
  const url = URL.createObjectURL(new Blob([data], { type }))
  const a = document.createElement('a')
  a.href = url
  a.download = name
  a.click()
  setTimeout(() => URL.revokeObjectURL(url), 10_000)
}

export default function ClaimBuilder() {
  const { loaded, charities, selectedId, uploadTargetId, chooseUploadTarget, refresh } = useCharity()

  const [phase, setPhase] = useState<Phase>('idle')
  const [uploads, setUploads] = useState<Upload[]>([])
  const [mappedTables, setMappedTables] = useState<Record<string, Table>>({})
  const [result, setResult] = useState<ClaimResult | null>(null)
  const [blocked, setBlocked] = useState<BlockedRow[]>([])
  const [fileCharityId, setFileCharityId] = useState('')
  const [error, setError] = useState<string | null>(null)
  const [saved, setSaved] = useState<SavedState | null>(null)
  const [mapperTable, setMapperTable] = useState<Table | null>(null)
  const [tab, setTab] = useState<'review' | 'ineligible'>('review')
  const [showNotes, setShowNotes] = useState(false)
  const [savedTab, setSavedTab] = useState<'files' | 'review' | 'ineligible'>('files')
  const [claimState, setClaimState] = useState<Record<string, ClaimDto>>({})
  const [submitFor, setSubmitFor] = useState<{ id: string; ref: string; hmrcRef: string | null } | null>(null)
  const [downloaded, setDownloaded] = useState(false)
  const [message, setMessage] = useState('Reading your files...')
  // Donations fixed by their donors that could still join this claim, and the ones that did (in file order)
  const [poolAvail, setPoolAvail] = useState<RecoveryRowDto[]>([])
  const [panelKey, setPanelKey] = useState(0)
  const poolInClaim = useRef<RecoveryRowDto[]>([])
  const tableCache = useRef(new Map<string, Table[]>())
  const allTables = useRef<Table[]>([])

  const noCharities = loaded && charities.length === 0
  const canUpload = !!uploadTargetId && phase === 'idle'
  const needsChoice = loaded && charities.length > 1 && selectedId === 'all' && !uploadTargetId
  const fileCharity = charities.find((c) => c.id === fileCharityId)
  const uploadTargetName = charities.find((c) => c.id === uploadTargetId)?.name

  /** Reads the files, applies the Gift Aid rules, and stops anything already claimed for this charity. */
  const analyse = async (ups: Upload[], mapped: Record<string, Table>, charityId: string) => {
    const tables = ups.flatMap((u) => {
      const key = `${u.name}:${u.data.byteLength}`
      let t = tableCache.current.get(key)
      if (!t) {
        t = readTables(u.name, u.data)
        tableCache.current.set(key, t)
      }
      return t
    })
    allTables.current = tables
    const effective = tables.map((t) => mapped[tableKey(t)] ?? t)

    // Before anything is saved: donations already claimed for this charity are taken out of this claim
    // (repeatedly, because taking one out can bring back a copy of it from another sheet)
    const { result: res, blocked: blockedRows } = await withoutAlreadyClaimed(effective, async (items) =>
      (await apiJson<{ claimed: AlreadyClaimed[] }>('/api/batches/check', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ charityId, items }),
      })).claimed
    )

    // Donations fixed by donors that could join this claim (not ones already in it, and not ones this file now has ready)
    const hasPool = ups.some((u) => u.name === RECOVERED_FILE)
    if (hasPool) setPoolAvail([])
    else {
      const readyKeys = new Set([...res.ready, ...res.overseas].map((r) => rowKey(r.platform, r.donationId == null ? '' : String(r.donationId))))
      const pool = await fetchPool(charityId)
      setPoolAvail(pool.filter((r) => !readyKeys.has(rowKey(r.platform, r.donationId))))
    }

    setUploads(ups)
    setResult(res)
    setBlocked(blockedRows)
    setTab(res.review.length ? 'review' : 'ineligible')
    setPhase('review')
  }

  const run = async (fn: () => Promise<void>, msg: string) => {
    setError(null)
    setMessage(msg)
    setPhase('reading')
    try {
      await tick(30) // let the loading state paint before the heavy work starts
      await fn()
    } catch (e) {
      setError(e instanceof JustGivingError ? e.message : (e as Error).message)
      setPhase(uploads.length && result ? 'review' : 'idle')
    }
  }

  const startFiles = async (list: FileList | null, append: boolean) => {
    if (!list || list.length === 0) return
    const charityId = append ? fileCharityId : uploadTargetId
    if (!charityId) {
      setError('Choose which charity this upload is for before adding files.')
      return
    }
    const picked = Array.from(list)
    await run(async () => {
      const incoming: Upload[] = []
      for (const f of picked) incoming.push({ name: f.name, data: new Uint8Array(await f.arrayBuffer()) })
      const base = append ? uploads : []
      const fresh = incoming.filter((n) => !base.some((u) => u.name === n.name && u.data.byteLength === n.data.byteLength))
      const next = [...base, ...fresh]
      const mapped = append ? mappedTables : {}
      if (!append) {
        setMappedTables({})
        setSaved(null)
        setFileCharityId(charityId)
      }
      await analyse(next, mapped, charityId)
    }, 'Reading your files and applying the Gift Aid rules...')
  }

  const onMapped = async (mapping: ColumnMapping, config: GiftAidConfig) => {
    if (!mapperTable) return
    const mapped = { ...mappedTables, [tableKey(mapperTable)]: mapToCashBook(mapperTable, mapping, config) }
    setMappedTables(mapped)
    setMapperTable(null)
    await run(() => analyse(uploads, mapped, fileCharityId), 'Applying the Gift Aid rules...')
  }

  const fetchPool = async (charityId: string): Promise<RecoveryRowDto[]> => {
    try {
      const d = await apiJson<{ rows: RecoveryRowDto[] }>(`/api/recovery?charity=${charityId}`)
      return d.rows.filter((r) => r.status === 'recovered' && r.fixed)
    } catch {
      return []
    }
  }

  /** A claim made only of the donations donors fixed. */
  const startRecovered = async () => {
    const charityId = uploadTargetId
    if (!charityId) return
    await run(async () => {
      const pool = await fetchPool(charityId)
      if (!pool.length) throw new Error('There are no recovered donations waiting. Refresh and try again.')
      poolInClaim.current = pool
      setMappedTables({})
      setSaved(null)
      setFileCharityId(charityId)
      await analyse([recoveredUpload(pool)], {}, charityId)
    }, 'Preparing your recovered donations...')
  }

  /** Adds the donations donors fixed to the claim being reviewed. */
  const addRecovered = async () => {
    if (!poolAvail.length) return
    await run(async () => {
      poolInClaim.current = poolAvail
      await analyse([...uploads.filter((u) => u.name !== RECOVERED_FILE), recoveredUpload(poolAvail)], mappedTables, fileCharityId)
    }, 'Adding the recovered donations to this claim...')
  }

  // /?recovered=1 (from the Claims page) goes straight to a claim made of the donations donors fixed
  const params = useSearchParams()
  const autoStarted = useRef(false)
  useEffect(() => {
    if (params.get('recovered') === '1' && !autoStarted.current && loaded && uploadTargetId && phase === 'idle') {
      autoStarted.current = true
      startRecovered()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [params, loaded, uploadTargetId, phase])

  const reset = () => {
    poolInClaim.current = []
    setPoolAvail([])
    chooseUploadTarget('')
    setPhase('idle')
    setUploads([])
    setMappedTables({})
    setResult(null)
    setBlocked([])
    setSaved(null)
    setClaimState({})
    setSavedTab('files')
    setDownloaded(false)
    setError(null)
    setMapperTable(null)
    setFileCharityId('')
    allTables.current = []
  }

  const save = async () => {
    if (!result || !fileCharityId) return
    setError(null)
    setMessage('Saving the claim and building your HMRC files...')
    setPhase('saving')
    try {
      await tick(30)
      const template = new Uint8Array(await (await fetch('/hmrc_template.ods')).arrayBuffer())
      const schedules = await buildSchedules(result, template)
      const claimed = result.ready.length + result.overseas.length

      const res = await fetch('/api/batches', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          charityId: fileCharityId,
          filename: uploads.map((u) => u.name).join(' + '),
          stats: { totalRecords: result.linesRead, eligibleRecords: claimed + result.review.length, validRecords: claimed, invalidRecords: result.review.length },
          claims: result.files.map((f) => ({
            kind: f.kind,
            fileName: f.name,
            donations: f.rows.map((r) => {
              // A donation the donor fixed keeps its real platform and ID, and is tagged so the 5% fee can be worked out
              const rec = isRecoveredSource(r.source) ? poolInClaim.current[Number(r.source.slice(r.source.lastIndexOf('!') + 1)) - 2] : undefined
              return {
                title: r.title,
                firstName: r.firstName,
                lastName: r.lastName,
                address: r.house,
                postcode: r.postcode,
                donationDate: r.date,
                amount: r.amount,
                platform: rec ? rec.platform : r.platform,
                donationId: rec ? rec.donationId : r.donationId === null || r.donationId === undefined ? '' : String(r.donationId),
                hmrcRow: r.hmrcRow,
                recoveredRowId: rec?.id,
              }
            }),
          })),
          // Everything held back is kept, with the donor's email, so it can be emailed (rows from the recovered file are already stored)
          reviewRows: heldBackRows(result.review.filter((r) => !isRecoveredSource(r.source)), allTables.current),
        }),
      })
      const data = await res.json().catch(() => null)
      if (res.status === 409) {
        // Something was claimed in the meantime: re-check and show the updated claim
        await analyse(uploads, mappedTables, fileCharityId)
        throw new Error('Some donations in this claim turned out to be claimed already. They have been taken out. Please check the numbers and save again.')
      }
      if (!res.ok) throw new Error(data?.error || `Server responded ${res.status}`)

      const named: NamedSchedule[] = schedules.map((s, i) => {
        const ref: string = data.claims[i].ref
        return { ...s, ref, name: scheduleName(s.plan, ref) }
      })
      const pack = await assemblePack(result, uploads, named)
      setSaved({
        charityName: data.charity?.name ?? fileCharity?.name ?? '',
        claims: data.claims,
        possibleRepeats: data.possibleRepeats,
        schedules: named,
        simpleSheets: buildSimpleSheets(result).map((s, i) => ({ name: simpleSheetName(s.plan, data.claims[i].ref), data: s.data })),
        zip: pack.zip,
        workbook: pack.workbook,
        summaryHtml: pack.summaryHtml,
      })
      setClaimState({})
      setPanelKey((k) => k + 1)
      setSavedTab('files')
      setDownloaded(false)
      setPhase('saved')
      refresh().catch(() => {})
    } catch (e) {
      setError((e as Error).message)
      setPhase(result ? 'review' : 'idle')
    }
  }

  // ---------------------------------------------------------------- derived
  const claimedCount = result ? result.ready.length + result.overseas.length : 0
  const unknownSheets = result ? result.sheets.filter((s) => s.platform.startsWith('Unknown layout')) : []

  const subtitle =
    phase === 'saved'
      ? 'Your claim is saved.'
      : phase === 'idle' || phase === 'reading'
        ? 'Upload your donation exports and get them ready for HMRC.'
        : mapperTable
          ? 'Match the columns of this file to the fields HMRC needs.'
          : 'Check the numbers, then save the claim.'

  const uploadAction = noCharities ? (
    <Link href="/settings" className="inline-flex h-12 items-center rounded-full bg-ink px-6 text-base font-medium text-white hover:bg-black">
      Add a charity
    </Link>
  ) : phase === 'idle' ? (
    <label
      title={uploadTargetId ? undefined : 'Choose a charity first'}
      className={`inline-flex h-12 items-center rounded-full bg-ink px-6 text-base font-medium text-white ${canUpload ? 'cursor-pointer hover:bg-black' : 'cursor-not-allowed opacity-50'}`}
    >
      Upload files
      <input type="file" multiple accept=".xlsx,.xls,.csv" className="hidden" disabled={!canUpload} onChange={(e) => { startFiles(e.target.files, false); e.target.value = '' }} />
    </label>
  ) : (
    <Button variant="dark" size="lg" onClick={reset} disabled={phase === 'saving' || phase === 'reading'}>
      <RotateCcw className="h-4 w-4" /> Start again
    </Button>
  )

  return (
    <>
      <PageHeader title="Clean" subtitle={subtitle} action={uploadAction} />

      <div className="space-y-6">
        {error && (
          <div className="flex items-center gap-3 rounded-2xl border border-bad/20 bg-bad-soft p-5">
            <AlertCircle className="h-5 w-5 flex-shrink-0 text-bad" />
            <div>
              <div className="font-medium text-bad">Something went wrong</div>
              <p className="text-sm text-bad/90">{error}</p>
            </div>
          </div>
        )}

        {/* ---------- Idle ---------- */}
        {phase === 'idle' && (
          <>
            {noCharities ? (
              <NoCharityNotice />
            ) : needsChoice ? (
              <>
                <div className="rounded-2xl border border-line bg-white p-8">
                  <h2 className="font-display text-2xl font-semibold text-ink">Choose a charity first</h2>
                  <p className="mt-1 text-muted">Every upload is attached to one charity. Which one is this for?</p>
                  <div className="mt-5 grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">
                    {charities.map((c) => (
                      <button
                        key={c.id}
                        onClick={() => chooseUploadTarget(c.id)}
                        className="cursor-pointer rounded-xl border border-line px-5 py-4 text-left transition-colors hover:border-brand/50 hover:bg-brand-soft/40"
                      >
                        <div className="text-ink">{c.name}</div>
                        <div className="text-sm text-muted">{c.claimCount.toLocaleString()} claim{c.claimCount === 1 ? '' : 's'}</div>
                      </button>
                    ))}
                  </div>
                </div>
              </>
            ) : (
              <>
                {selectedId !== 'all' || charities.length <= 1 ? (
                  uploadTargetName ? (
                    <p className="text-ink">
                      Uploading for <span className="font-medium">{uploadTargetName}</span>
                    </p>
                  ) : null
                ) : (
                  <div className="flex flex-wrap items-center gap-3">
                    <label htmlFor="upload-charity" className="text-ink">Uploading for</label>
                    <Select id="upload-charity" value={uploadTargetId} onChange={(e) => chooseUploadTarget(e.target.value)} wrapperClassName="min-w-[16rem]">
                      {charities.map((c) => (
                        <option key={c.id} value={c.id}>{c.name}</option>
                      ))}
                    </Select>
                  </div>
                )}

                <label className={`block ${uploadTargetId ? '' : 'cursor-not-allowed opacity-50'}`}>
                  <div className={`rounded-2xl border-2 border-dashed border-[#cfc9e8] bg-[#faf9ff] p-12 text-center transition-colors ${uploadTargetId ? 'cursor-pointer hover:border-brand/50 hover:bg-brand-soft/50' : 'pointer-events-none'}`}>
                    <div className="mx-auto mb-4 flex h-14 w-14 items-center justify-center rounded-full border border-line bg-white">
                      <FileUp className="h-6 w-6 text-brand" />
                    </div>
                    <div className="font-display text-xl font-semibold text-ink">Drop your files here, or click to browse</div>
                    <div className="mt-1 text-muted">Excel or CSV. You can pick several files, even from different platforms, for one claim.</div>
                    <input type="file" multiple accept=".xlsx,.xls,.csv" className="hidden" disabled={!uploadTargetId} onChange={(e) => { startFiles(e.target.files, false); e.target.value = '' }} />
                  </div>
                </label>
              </>
            )}

            {!needsChoice && (
              <>
                {uploadTargetId && <RecoveryPanel charityId={uploadTargetId} refreshKey={panelKey} onCreateClaim={startRecovered} />}
                <div className="grid gap-5 md:grid-cols-3">
                  {[
                    { icon: Columns, title: '1. Upload', text: 'Add your exports. We recognise the platform from its columns.' },
                    { icon: Eye, title: '2. Check the rows', text: 'See what is ready, what needs a decision, and why.' },
                    { icon: Download, title: '3. Save the claim', text: 'Get the HMRC .ods schedules, a review workbook and a summary.' },
                  ].map(({ icon: Icon, title, text }) => (
                    <div key={title} className="flex gap-4 rounded-2xl border border-line bg-white p-6">
                      <div className="flex h-11 w-11 flex-shrink-0 items-center justify-center rounded-full bg-brand-soft text-brand">
                        <Icon className="h-5 w-5" />
                      </div>
                      <div>
                        <div className="font-display text-lg font-semibold text-ink">{title}</div>
                        <p className="mt-1 text-muted">{text}</p>
                      </div>
                    </div>
                  ))}
                </div>
                <p className="text-sm text-muted">
                  Platforms we recognise: {PROFILES.map((p) => p.name).join(', ')}. Any other layout can be matched by hand.
                </p>
              </>
            )}
          </>
        )}

        {/* ---------- Reading / saving ---------- */}
        {(phase === 'reading' || phase === 'saving') && <ReviewSkeleton message={message} />}

        {/* ---------- Review ---------- */}
        {phase === 'review' && result && mapperTable && (
          <>
            <ColumnMapper detectedColumns={(mapperTable.rows[0] ?? []).map((h, i) => String(h ?? '').trim() || `Column ${i + 1}`)} sampleData={tableRecords(mapperTable)} onMappingComplete={onMapped} />
            <Button variant="outline" onClick={() => setMapperTable(null)}>Cancel</Button>
          </>
        )}

        {phase === 'review' && result && !mapperTable && (
          <>
            <FilesRead result={result} />

            {poolAvail.length > 0 && (
              <div className="flex flex-wrap items-center justify-between gap-4 rounded-2xl border border-good-line bg-good-soft px-6 py-5 text-good">
                <div>
                  <div className="font-display text-xl font-bold tabular-nums">{gbp(giftAidOn(poolAvail.reduce((a, r) => a + r.amount, 0)))} recovered and waiting</div>
                  <div>{poolAvail.length} donation{poolAvail.length === 1 ? '' : 's'} fixed by donors. Add {poolAvail.length === 1 ? 'it' : 'them'} to this claim?</div>
                </div>
                <Button variant="success" onClick={addRecovered}>Add to this claim</Button>
              </div>
            )}

            <ResultTiles
              result={result}
              onReview={() => { setTab('review'); document.getElementById('rows-panel')?.scrollIntoView({ behavior: 'smooth', block: 'start' }) }}
            />

            <div className="rounded-2xl border-2 border-dashed border-[#cfc9e8] bg-[#faf9ff] p-6">
              <div className="flex flex-wrap items-center justify-between gap-4">
                <div className="min-w-0">
                  <div className="font-display text-2xl font-semibold text-ink">
                    {uploads.length === 1 ? uploads[0].name : `${uploads.length} files`}
                  </div>
                  <div className="mt-1 text-muted">
                    {fileCharity ? `${fileCharity.name} · ` : ''}
                    {result.files.length} HMRC file{result.files.length === 1 ? '' : 's'}
                    {result.overseas.length > 0 ? ` (${result.overseas.length} overseas donation${result.overseas.length === 1 ? '' : 's'} in their own file)` : ''}
                  </div>
                </div>
                <div className="flex flex-wrap gap-3">
                  <label className="inline-flex h-12 cursor-pointer items-center gap-2 rounded-full border border-line-strong bg-white px-6 text-base font-medium text-ink hover:bg-wash">
                    <Plus className="h-4 w-4" /> Add more files
                    <input type="file" multiple accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => { startFiles(e.target.files, true); e.target.value = '' }} />
                  </label>
                  <Button size="lg" onClick={save} disabled={claimedCount === 0}>
                    Save claim &amp; make files
                  </Button>
                </div>
              </div>
            </div>

            {unknownSheets.length > 0 && (
              <div className="rounded-2xl border border-warn-line bg-warn-soft p-5">
                <div className="flex gap-3">
                  <AlertCircle className="mt-0.5 h-5 w-5 flex-shrink-0 text-warn" />
                  <div className="w-full">
                    <div className="font-medium text-warn">We don&apos;t recognise {unknownSheets.length === 1 ? 'this sheet' : 'these sheets'}</div>
                    <p className="text-sm text-warn/90">Match its columns by hand and the same Gift Aid rules will be applied to it.</p>
                    <div className="mt-3 space-y-2">
                      {unknownSheets.map((s) => {
                        const t = allTables.current.find((x) => tableKey(x) === s.table)
                        return (
                          <div key={s.table} className="flex flex-wrap items-center justify-between gap-3 rounded-xl bg-white/70 px-4 py-3">
                            <span className="text-ink">{s.table}</span>
                            {t && <Button size="sm" variant="outline" onClick={() => setMapperTable(t)}>Match columns</Button>}
                          </div>
                        )
                      })}
                    </div>
                  </div>
                </div>
              </div>
            )}

            {blocked.length > 0 && (
              <div className="rounded-2xl border border-warn-line bg-warn-soft p-5">
                <div className="flex gap-3">
                  <AlertCircle className="mt-0.5 h-5 w-5 flex-shrink-0 text-warn" />
                  <div className="w-full">
                    <div className="font-medium text-warn">
                      {blocked.length.toLocaleString()} donation{blocked.length === 1 ? '' : 's'} already claimed, left out of this claim
                    </div>
                    <p className="text-sm text-warn/90">A donation can only be claimed once. These are in an earlier claim for this charity and were taken out before saving.</p>
                    <ul className="mt-3 space-y-1 text-sm text-ink">
                      {blocked.slice(0, 5).map((b) => (
                        <li key={`${b.platform}-${b.donationId}`}>
                          {b.name || 'Donor'} · {gbp(b.amount)} · {b.platform} {b.donationId} · already in <span className="font-medium">{b.claimRef}</span>
                        </li>
                      ))}
                      {blocked.length > 5 && <li className="text-muted">and {blocked.length - 5} more</li>}
                    </ul>
                  </div>
                </div>
              </div>
            )}

            <div className="overflow-hidden rounded-2xl border border-line bg-white">
              <button onClick={() => setShowNotes((o) => !o)} className="flex w-full cursor-pointer items-center justify-between px-5 py-4 text-left hover:bg-wash/60">
                <div>
                  <h3 className="font-display text-lg font-semibold text-ink">How we read your files</h3>
                  <p className="text-sm text-muted">Platform, date style and any copied rows skipped, per sheet</p>
                </div>
                <ChevronRight className={`h-5 w-5 text-muted transition-transform ${showNotes ? 'rotate-90' : ''}`} />
              </button>
              {showNotes && (
                <div className="overflow-x-auto border-t border-line">
                  <table className="w-full text-base">
                    <thead>
                      <tr className="border-b border-line">
                        <th className={th}>Sheet</th>
                        <th className={th}>Platform</th>
                        <th className={thr}>Rows used</th>
                        <th className={thr}>Copies skipped</th>
                        <th className={th}>Date order</th>
                      </tr>
                    </thead>
                    <tbody>
                      {result.sheets.map((s) => (
                        <tr key={s.table} className="border-b border-line-soft last:border-0">
                          <td className={`${td} text-ink`}>{s.table}</td>
                          <td className={td}>{s.platform.startsWith('Unknown') ? <Pill tone="warn">Unknown layout</Pill> : <Pill tone="brand">{s.platform}</Pill>}</td>
                          <td className={`${tdr} text-ink`}>{s.rowsUsed.toLocaleString()}</td>
                          <td className={`${tdr} text-ink`}>{s.copiesSkipped.toLocaleString()}</td>
                          <td className={`${td} text-ink`}>{s.dateOrder || '-'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            <OptionalChecks result={result} />

            <div id="rows-panel" className="space-y-4">
              <Tabs
                tabs={[
                  { id: 'review' as const, label: `Needs review (${result.review.length.toLocaleString()})` },
                  { id: 'ineligible' as const, label: `Not eligible (${result.ineligible.length.toLocaleString()})` },
                ]}
                value={tab}
                onChange={(id) => { setTab(id) }}
              />
              <RowsPanel key={tab} result={result} tab={tab} />
            </div>
          </>
        )}

        {/* ---------- Saved ---------- */}
        {phase === 'saved' && saved && result && (
          <>
            <div className="space-y-1.5">
              <span className="text-sm font-bold text-good">✓ Claim saved</span>
              <h2 className="font-display text-3xl font-bold tracking-tight text-ink">
                {saved.claims.length === 1 ? `Claim ${saved.claims[0].ref}` : `${saved.claims.length} claims, ${saved.claims[0].ref} to ${saved.claims[saved.claims.length - 1].ref}`} · {gbp(result.totals.giftAid)} Gift Aid
              </h2>
              <p className="text-muted">
                {saved.charityName} · {result.platforms.join(', ')} · {new Date().toLocaleDateString('en-GB')}
              </p>
            </div>

            <NextSteps
              claimIds={saved.claims.map((c) => c.id)}
              downloaded={downloaded}
              submittedCount={saved.claims.filter((c) => (claimState[c.id]?.status ?? 'pending') !== 'pending').length}
              onDownloadAll={() => { download(`BoostAid_claim_pack_${saved.claims[0]?.ref ?? ''}.zip`, saved.zip as BlobPart, 'application/zip'); setDownloaded(true) }}
              onGoToFiles={() => { setSavedTab('files'); document.getElementById('saved-panel')?.scrollIntoView({ behavior: 'smooth', block: 'start' }) }}
            />

            <FilesRead result={result} />
            <ResultTiles result={result} saved onReview={() => { setSavedTab('review'); document.getElementById('saved-panel')?.scrollIntoView({ behavior: 'smooth', block: 'start' }) }} />

            {saved.possibleRepeats.count > 0 && (
              <div className="flex gap-3 rounded-2xl border border-warn-line bg-warn-soft p-5 text-sm text-warn">
                <AlertCircle className="h-5 w-5 flex-shrink-0" />
                <div>
                  <span className="font-medium">{saved.possibleRepeats.count.toLocaleString()} donation{saved.possibleRepeats.count === 1 ? '' : 's'} without a platform ID look identical</span>{' '}
                  to ones already claimed ({saved.possibleRepeats.claimRefs.join(', ')}). Check you are not claiming them twice.
                </div>
              </div>
            )}

            <div id="saved-panel" className="space-y-4">
              <Tabs
                tabs={[
                  { id: 'files' as const, label: `HMRC files (${saved.schedules.length})` },
                  { id: 'review' as const, label: `Needs review (${result.review.length.toLocaleString()})` },
                  { id: 'ineligible' as const, label: `Not eligible (${result.ineligible.length.toLocaleString()})` },
                ]}
                value={savedTab}
                onChange={setSavedTab}
              />

              {savedTab === 'files' ? (
                <>
                  <div className="overflow-x-auto rounded-2xl border border-line bg-white">
                    <table className="w-full min-w-[920px] text-base">
                      <thead>
                        <tr className="border-b border-line">
                          <th className={th}>File to upload</th>
                          <th className={th} title="Earliest donation in this file">Date</th>
                          <th className={thr}>Donations</th>
                          <th className={thr}>Amount</th>
                          <th className={thr}>Gift Aid</th>
                          <th className={th}>Status</th>
                          <th className={th}><span className="sr-only">Actions</span></th>
                        </tr>
                      </thead>
                      <tbody>
                        {saved.schedules.map((sc, i) => {
                          const c = saved.claims[i]
                          const live = claimState[c.id]
                          const status: Status = live?.status ?? 'pending'
                          return (
                            <tr key={sc.name} className="border-b border-line-soft last:border-0">
                              <td className={td}>
                                <Link href={`/claims/${c.id}`} className="text-ink hover:text-brand">{sc.name}</Link>
                                <div className="text-sm text-muted">
                                  {c.kind} · {c.ref}{live?.hmrcRef ? ` · HMRC ${live.hmrcRef}` : ''}
                                </div>
                              </td>
                              <td className={`${td} whitespace-nowrap text-ink`}>{fmtPlainDate(sc.plan.firstDate || null)}</td>
                              <td className={`${tdr} text-ink`}>{c.donationCount.toLocaleString()}</td>
                              <td className={`${tdr} text-ink`}>{gbp(c.totalAmount)}</td>
                              <td className={`${tdr} text-ink`}>{gbp(giftAidOn(c.totalAmount))}</td>
                              <td className={td}><StatusBadge status={status} /></td>
                              <td className={`${td} whitespace-nowrap text-right`}>
                                <Button variant="outline" size="sm" onClick={() => download(sc.name, sc.data as BlobPart, ODS)}>Download</Button>{' '}
                                {status === 'pending' ? (
                                  <Button variant="dark" size="sm" onClick={() => setSubmitFor({ id: c.id, ref: c.ref, hmrcRef: live?.hmrcRef ?? null })}>Mark submitted</Button>
                                ) : (
                                  <Link href={`/claims/${c.id}`} className="inline-flex h-9 items-center rounded-full border border-line-strong bg-white px-4 text-sm font-medium text-ink hover:bg-wash">Open</Link>
                                )}
                              </td>
                            </tr>
                          )
                        })}
                      </tbody>
                    </table>
                  </div>

                  <div className="flex flex-wrap gap-3">
                    <Button variant="outline" onClick={() => download('2_Claim_workbook.xlsx', saved.workbook as BlobPart, XLSX_MIME)}>
                      <FileSpreadsheet className="h-4 w-4 text-info" /> Claim workbook (.xlsx)
                    </Button>
                    <Button variant="outline" onClick={() => window.open(URL.createObjectURL(new Blob([saved.summaryHtml], { type: 'text/html' })), '_blank')}>
                      <FileText className="h-4 w-4" /> Summary (print or save as PDF)
                    </Button>
                  </div>

                  <Card className="p-6">
                    <div className="flex flex-wrap items-center gap-3">
                      <h3 className="font-display text-xl font-semibold text-ink">Simple output sheets</h3>
                      <Pill tone="neutral">Optional</Pill>
                    </div>
                    <p className="mb-4 mt-1 text-muted">
                      Not needed for HMRC. Plain Excel copies of each schedule&apos;s clean data, without HMRC&apos;s template. Handy as a backup or for your own records.
                    </p>
                    <div className="grid grid-cols-1 gap-2 md:grid-cols-2">
                      {saved.simpleSheets.map((sh) => (
                        <Button key={sh.name} variant="outline" className="h-auto justify-start whitespace-normal py-3 text-left" onClick={() => download(sh.name, sh.data as BlobPart, XLSX_MIME)}>
                          <FileSpreadsheet className="h-4 w-4 flex-shrink-0 text-info" /> {sh.name}
                        </Button>
                      ))}
                    </div>
                    {saved.simpleSheets.length > 1 && (
                      <Button
                        variant="outline"
                        className="mt-3"
                        onClick={async () => {
                          for (const sh of saved.simpleSheets) {
                            download(sh.name, sh.data as BlobPart, XLSX_MIME)
                            await tick(300)
                          }
                        }}
                      >
                        <Download className="h-4 w-4" /> Download all simple sheets
                      </Button>
                    )}
                  </Card>
                </>
              ) : (
                <RowsPanel key={savedTab} result={result} tab={savedTab} />
              )}
            </div>

            {blocked.length > 0 && (
              <p className="text-sm text-muted">{blocked.length.toLocaleString()} donation{blocked.length === 1 ? ' was' : 's were'} left out because they were already claimed.</p>
            )}
          </>
        )}

        <SubmitDialog
          claim={submitFor}
          onClose={() => setSubmitFor(null)}
          onDone={(updated) => setClaimState((m) => ({ ...m, [updated.id]: updated }))}
        />

        {phase === 'saved' && !saved && <Skeleton className="h-40 w-full" />}
        {phase === 'saved' && saved && fileCharityId && <RecoveryPanel charityId={fileCharityId} refreshKey={panelKey} />}
      </div>
    </>
  )
}

const HMRC_URL = 'https://www.gov.uk/claim-tax-back-on-donations-using-charities-online'

/** "Files read" chips: which platform each sheet was recognised as, and any copied rows skipped. */
function FilesRead({ result }: { result: ClaimResult }) {
  const label = (table: string) => table.replace(/:(Sheet1|CSV)$/, '').replace(':', ' · ')
  return (
    <div className="flex flex-wrap items-center gap-2.5 text-sm">
      <span className="font-semibold text-muted">Files read:</span>
      {result.sheets.map((sh) =>
        sh.platform.startsWith('Unknown layout') ? (
          <span key={sh.table} className="rounded-full bg-warn-pill px-3 py-1.5 font-medium text-warn">
            {label(sh.table)} · unknown layout
          </span>
        ) : (
          <span key={sh.table} className="rounded-full bg-good-pill px-3 py-1.5 font-medium text-good">
            ✓ {label(sh.table)} · {sh.platform} · {sh.rowsUsed.toLocaleString()} rows
          </span>
        )
      )}
      {result.sheets
        .filter((sh) => sh.copiesSkipped > 0)
        .map((sh) => (
          <span key={`skip-${sh.table}`} className="rounded-full bg-[#f1eff6] px-3 py-1.5 text-muted">
            {label(sh.table)}: {sh.copiesSkipped.toLocaleString()} exact copy row{sh.copiesSkipped === 1 ? '' : 's'} skipped
          </span>
        ))}
    </div>
  )
}

/** The four figures from the design: rows read, in the claim, Gift Aid to reclaim, held back for review. */
function ResultTiles({ result, onReview }: { result: ClaimResult; saved?: boolean; onReview: () => void }) {
  const inClaim = result.ready.length + result.overseas.length
  return (
    <div className="grid grid-cols-2 gap-4 lg:grid-cols-4">
      <StatTile
        label="Rows read"
        value={result.linesRead.toLocaleString()}
        caption={`${result.ineligible.length.toLocaleString()} not eligible (no Gift Aid, fees, refunds)`}
      />
      <StatTile label="In the claim" value={inClaim.toLocaleString()} caption={`${gbp(result.totals.ukDonations + result.totals.overseasDonations)} donated`} />
      <StatTile tone="good" label="Gift Aid to reclaim" value={gbp(result.totals.giftAid)} caption="25% of the claim" />
      <button onClick={onReview} className="block h-full w-full cursor-pointer rounded-2xl text-left transition hover:brightness-95">
        <StatTile
          tone="bad"
          label="Held back for review"
          value={result.review.length.toLocaleString()}
          caption={`${gbp(result.totals.reviewDonations)} · Review now →`}
        />
      </button>
    </div>
  )
}

/** The three steps after saving: download, upload to HMRC, mark submitted. */
function NextSteps({
  claimIds,
  downloaded,
  submittedCount,
  onDownloadAll,
  onGoToFiles,
}: {
  claimIds: string[]
  downloaded: boolean
  submittedCount: number
  onDownloadAll: () => void
  onGoToFiles: () => void
}) {
  const total = claimIds.length
  const allSubmitted = total > 0 && submittedCount === total
  const box = (done: boolean, active: boolean) =>
    done ? 'border border-good-line bg-good-soft' : active ? 'border-2 border-accent bg-brand-soft' : 'border border-line bg-canvas'
  const badge = (done: boolean) => `inline-flex h-7 w-7 items-center justify-center rounded-full text-sm font-bold text-white ${done ? 'bg-good' : 'bg-brand'}`
  const outlineLink = 'inline-flex h-10 items-center self-start rounded-full border border-line-strong bg-white px-4 text-sm font-medium text-ink hover:bg-wash'

  return (
    <section aria-label="Next steps" className="grid grid-cols-1 gap-4 rounded-2xl border border-line bg-white p-5 md:grid-cols-3">
      <div className={`flex flex-col gap-2.5 rounded-xl p-4 ${box(downloaded, !downloaded)}`}>
        <div className="flex items-center gap-2.5"><span className={badge(downloaded)}>{downloaded ? '✓' : '1'}</span><span className="font-semibold text-ink">Download the files</span></div>
        <span className="text-sm text-muted">{total} HMRC file{total === 1 ? '' : 's'} (.ods), ready exactly as HMRC needs them.</span>
        <Button size="sm" className="self-start" onClick={onDownloadAll}>{downloaded ? 'Downloaded (again)' : 'Download all (.zip)'}</Button>
      </div>
      <div className={`flex flex-col gap-2.5 rounded-xl p-4 ${box(allSubmitted, downloaded && !allSubmitted)}`}>
        <div className="flex items-center gap-2.5"><span className={badge(allSubmitted)}>{allSubmitted ? '✓' : '2'}</span><span className="font-semibold text-ink">Upload to Charities Online</span></div>
        <span className="text-sm text-muted">Sign in to HMRC Charities Online and upload each file. Note the HMRC reference.</span>
        <a href={HMRC_URL} target="_blank" rel="noreferrer" className={outlineLink}>HMRC guidance ↗</a>
      </div>
      <div className={`flex flex-col gap-2.5 rounded-xl p-4 ${box(allSubmitted, false)}`}>
        <div className="flex items-center gap-2.5"><span className={badge(allSubmitted)}>{allSubmitted ? '✓' : '3'}</span><span className="font-semibold text-ink">Mark as submitted</span></div>
        <span className="text-sm text-muted">
          {allSubmitted ? 'All files submitted. Mark them received in Claims when HMRC pays.' : `${submittedCount} of ${total} file${total === 1 ? '' : 's'} marked submitted. You will be asked for the HMRC reference.`}
        </span>
        <button onClick={onGoToFiles} className={`${outlineLink} cursor-pointer`}>Go to the file list ↓</button>
      </div>
    </section>
  )
}

/** Needs-review / not-eligible rows, paged. Re-mounted (key) when the tab changes. */
function RowsPanel({ result, tab }: { result: ClaimResult; tab: 'review' | 'ineligible' }) {
  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = useState(25)
  const total = tab === 'review' ? result.review.length : result.ineligible.length
  const start = (page - 1) * pageSize

  return (
    <div className="space-y-3">
      <p className="text-muted">
        {tab === 'review'
          ? 'These are left out of the HMRC files. Each has a reason, and the review workbook lists them all.'
          : 'These can never be claimed. Each has a reason.'}
      </p>
      <div className="overflow-hidden rounded-2xl border border-line bg-white">
        {total === 0 ? (
          <div className="py-12 text-center text-muted">{tab === 'review' ? 'Nothing to review. Every eligible donation passed the checks.' : 'Nothing here.'}</div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-base">
              <thead>
                <tr className="border-b border-line">
                  <th className={th}>{tab === 'review' ? "Why it's held back" : 'Reason'}</th>
                  <th className={th}>{tab === 'review' ? 'Donor (as uploaded)' : 'Donor'}</th>
                  {tab === 'review' && <th className={th}>Address (as uploaded)</th>}
                  {tab === 'review' && <th className={th}>Postcode</th>}
                  <th className={thr}>Amount</th>
                  <th className={th}>{tab === 'review' ? 'Platform' : 'Source'}</th>
                </tr>
              </thead>
              <tbody>
                {tab === 'review'
                  ? result.review.slice(start, start + pageSize).map((r, i) => (
                      <tr key={start + i} className="border-b border-line-soft last:border-0">
                        <td className={`${td} max-w-[300px]`}>
                          <div className="flex flex-wrap gap-1.5">
                            {r.reason.split('; ').map((x) => <Pill key={x} tone="bad" className="whitespace-normal">{x}</Pill>)}
                          </div>
                        </td>
                        <td className={`${td} text-ink`}>
                          {r.originalName || '-'}
                          {r.suggestedName && <div className="text-sm text-muted">Suggested: {r.suggestedName}</div>}
                        </td>
                        <td className={`${td} max-w-[200px] truncate text-ink`} title={r.originalAddress}>{r.originalAddress || '-'}</td>
                        <td className={`${td} whitespace-nowrap text-ink`}>{r.originalPostcode || '-'}</td>
                        <td className={`${tdr} text-ink`}>{/ GBP$/.test(r.originalAmount) ? gbp(r.amount) : r.originalAmount}</td>
                        <td className={`${td} text-ink`}>{r.platform}</td>
                      </tr>
                    ))
                  : result.ineligible.slice(start, start + pageSize).map((r, i) => (
                      <tr key={start + i} className="border-b border-line-soft last:border-0">
                        <td className={td}><Pill tone="neutral">{r.reason}</Pill></td>
                        <td className={`${td} text-ink`}>{r.name || '-'}</td>
                        <td className={`${tdr} text-ink`}>{gbp(r.amount)}</td>
                        <td className={`${td} max-w-[260px] truncate text-muted`} title={r.source}>{r.source}</td>
                      </tr>
                    ))}
              </tbody>
            </table>
          </div>
        )}
        <Pagination page={page} pageSize={pageSize} total={total} onPage={setPage} onPageSize={(n) => { setPageSize(n); setPage(1) }} />
      </div>
    </div>
  )
}
