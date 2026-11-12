'use client'

import { useRef, useState } from 'react'
import Link from 'next/link'
import { AlertCircle, CheckCircle2, ChevronRight, Columns, Download, Eye, FileSpreadsheet, FileText, FileUp, Plus, RotateCcw } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Card, PageHeader, Pill, StatTile, td, tdr, th, thr } from '@/components/ui/primitives'
import { Pagination } from '@/components/ui/pagination'
import { ReviewSkeleton, Skeleton } from '@/components/ui/skeleton'
import { Select } from '@/components/ui/select'
import NoCharityNotice from '@/components/charity/no-charity-notice'
import SummaryStrip from '@/components/claims/summary-strip'
import ColumnMapper, { type ColumnMapping, type GiftAidConfig } from '@/components/column-mapper'
import { useCharity } from '@/lib/store/charity'
import { apiJson, gbp } from '@/components/claims/shared'
import { buildClaim, JustGivingError, PROFILES, readTables, type ClaimResult } from '@/lib/giftaid'
import type { Table } from '@/lib/giftaid/read'
import { blankSources, mapToCashBook, tableKey, tableRecords, type Upload } from '@/lib/claim-flow'
import { assemblePack, buildSchedules, scheduleName, type NamedSchedule } from '@/lib/claim-pack'
import { buildSimpleSheets, simpleSheetName } from '@/lib/claim-extras'
import OptionalChecks from '@/components/claim-optional-checks'

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
  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = useState(25)
  const [showNotes, setShowNotes] = useState(false)
  const [message, setMessage] = useState('Reading your files...')
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

    let res = buildClaim(effective)
    let blockedRows: BlockedRow[] = []

    // Before anything is saved: donations already claimed for this charity are taken out of this claim
    const withIds = [...res.ready, ...res.overseas].filter((r) => r.donationId !== null && r.donationId !== undefined && String(r.donationId) !== '')
    if (withIds.length) {
      const found = new Map<string, string>()
      for (let i = 0; i < withIds.length; i += 20000) {
        const chunk = withIds.slice(i, i + 20000)
        const { claimed } = await apiJson<{ claimed: { platform: string; donationId: string; claimRef: string }[] }>('/api/batches/check', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ charityId, items: chunk.map((r) => ({ platform: r.platform, donationId: String(r.donationId) })) }),
        })
        for (const c of claimed) found.set(`${c.platform}\u0000${c.donationId}`, c.claimRef)
      }
      if (found.size) {
        const hit = withIds.filter((r) => found.has(`${r.platform}\u0000${String(r.donationId)}`))
        blockedRows = hit.map((r) => ({
          name: `${r.firstName} ${r.lastName}`.trim(),
          amount: r.amount,
          platform: r.platform,
          donationId: String(r.donationId),
          claimRef: found.get(`${r.platform}\u0000${String(r.donationId)}`) ?? '',
        }))
        res = buildClaim(blankSources(effective, hit.map((r) => r.source)))
      }
    }

    setUploads(ups)
    setResult(res)
    setBlocked(blockedRows)
    setTab(res.review.length ? 'review' : 'ineligible')
    setPage(1)
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

  const reset = () => {
    chooseUploadTarget('')
    setPhase('idle')
    setUploads([])
    setMappedTables({})
    setResult(null)
    setBlocked([])
    setSaved(null)
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
            donations: f.rows.map((r) => ({
              title: r.title,
              firstName: r.firstName,
              lastName: r.lastName,
              address: r.house,
              postcode: r.postcode,
              donationDate: r.date,
              amount: r.amount,
              platform: r.platform,
              donationId: r.donationId === null || r.donationId === undefined ? '' : String(r.donationId),
              hmrcRow: r.hmrcRow,
            })),
          })),
        }),
      })
      const data = await res.json().catch(() => null)
      if (res.status === 409) {
        // Something was claimed in the meantime: re-check and show the updated claim
        await analyse(uploads, mappedTables, fileCharityId)
        throw new Error('Some of these donations were claimed while you were reviewing. They have been taken out, so please check the numbers and save again.')
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
      setPhase('saved')
      refresh().catch(() => {})
    } catch (e) {
      setError((e as Error).message)
      setPhase(result ? 'review' : 'idle')
    }
  }

  // ---------------------------------------------------------------- derived
  const claimedCount = result ? result.ready.length + result.overseas.length : 0
  const claimedTotal = result ? result.totals.ukDonations + result.totals.overseasDonations : 0
  const unknownSheets = result ? result.sheets.filter((s) => s.platform.startsWith('Unknown layout')) : []
  const list = result ? (tab === 'review' ? result.review : result.ineligible) : []
  const start = (page - 1) * pageSize

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
                <SummaryStrip />
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
                <SummaryStrip />
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
            <div className="grid grid-cols-2 gap-5 lg:grid-cols-4">
              <StatTile label="Rows read" value={result.linesRead.toLocaleString()} caption={`${result.platforms.length} platform${result.platforms.length === 1 ? '' : 's'}`} />
              <StatTile label="Ready for HMRC" tone="good" value={claimedCount.toLocaleString()} caption={`${gbp(claimedTotal)} · ${gbp(result.totals.giftAid)} Gift Aid`} />
              <StatTile label="Needs review" tone="bad" value={result.review.length.toLocaleString()} caption={`${gbp(result.totals.reviewDonations)} held back`} />
              <StatTile label="Not eligible" value={result.ineligible.length.toLocaleString()} caption="No declaration, refunds, etc." />
            </div>

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
                  <label className="inline-flex h-12 cursor-pointer items-center gap-2 rounded-full border border-line bg-white px-6 text-base font-medium text-ink hover:bg-wash">
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
              <div className="rounded-2xl border border-warn/20 bg-warn-soft p-5">
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
              <div className="rounded-2xl border border-warn/20 bg-warn-soft p-5">
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
                        <tr key={s.table} className="border-b border-line last:border-0">
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

            <div>
              <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
                <div className="flex gap-1 rounded-full border border-line bg-wash p-1">
                  {([['review', `Needs review (${result.review.length.toLocaleString()})`], ['ineligible', `Not eligible (${result.ineligible.length.toLocaleString()})`]] as const).map(([id, label]) => (
                    <button
                      key={id}
                      onClick={() => { setTab(id); setPage(1) }}
                      className={`h-9 cursor-pointer rounded-full border px-4 text-sm font-medium transition-colors ${tab === id ? 'border-line bg-white text-ink shadow-sm' : 'border-transparent text-muted hover:text-ink'}`}
                    >
                      {label}
                    </button>
                  ))}
                </div>
                <p className="text-muted">
                  {tab === 'review'
                    ? 'These are left out of the HMRC files. Each has a reason, and the review workbook lists them all.'
                    : 'These can never be claimed. Each has a reason.'}
                </p>
              </div>

              <div className="overflow-hidden rounded-2xl border border-line bg-white">
                {list.length === 0 ? (
                  <div className="py-12 text-center text-muted">{tab === 'review' ? 'Nothing to review. Every eligible donation passed the checks.' : 'Nothing here.'}</div>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full text-base">
                      <thead>
                        <tr className="border-b border-line">
                          <th className={th}>Reason</th>
                          <th className={th}>Donor</th>
                          {tab === 'review' && <th className={th}>Address</th>}
                          {tab === 'review' && <th className={th}>Postcode</th>}
                          <th className={thr}>Amount</th>
                          <th className={th}>{tab === 'review' ? 'Platform' : 'Source'}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {tab === 'review'
                          ? result.review.slice(start, start + pageSize).map((r, i) => (
                              <tr key={start + i} className="border-b border-line last:border-0">
                                <td className={td}>
                                  <div className="flex flex-wrap gap-1.5">
                                    {r.reason.split('; ').map((x) => <Pill key={x} tone="bad">{x}</Pill>)}
                                  </div>
                                </td>
                                <td className={`${td} text-ink`}>
                                  {r.originalName || '-'}
                                  {r.suggestedName && <div className="text-sm text-muted">Suggested: {r.suggestedName}</div>}
                                </td>
                                <td className={`${td} max-w-[260px] truncate text-ink`} title={r.originalAddress}>{r.originalAddress || '-'}</td>
                                <td className={`${td} whitespace-nowrap text-ink`}>{r.originalPostcode || '-'}</td>
                                <td className={`${tdr} text-ink`}>{/ GBP$/.test(r.originalAmount) ? gbp(r.amount) : r.originalAmount}</td>
                                <td className={`${td} text-ink`}>{r.platform}</td>
                              </tr>
                            ))
                          : result.ineligible.slice(start, start + pageSize).map((r, i) => (
                              <tr key={start + i} className="border-b border-line last:border-0">
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
                <Pagination page={page} pageSize={pageSize} total={list.length} onPage={setPage} onPageSize={(n) => { setPageSize(n); setPage(1) }} />
              </div>
            </div>
          </>
        )}

        {/* ---------- Saved ---------- */}
        {phase === 'saved' && saved && result && (
          <>
            <div className="flex flex-wrap items-center justify-between gap-4 rounded-2xl border border-line bg-white p-6">
              <div className="flex items-center gap-4">
                <div className="flex h-12 w-12 items-center justify-center rounded-full bg-good-soft">
                  <CheckCircle2 className="h-6 w-6 text-good" />
                </div>
                <div>
                  <h2 className="font-display text-2xl font-semibold text-ink">Claim saved</h2>
                  <p className="text-muted">Saved for {saved.charityName}. Your files are ready to download.</p>
                </div>
              </div>
              <div className="flex flex-wrap gap-3">
                <Button onClick={() => download(`BoostAid_claim_pack_${saved.claims[0]?.ref ?? ''}.zip`, saved.zip as BlobPart, 'application/zip')}>
                  <Download className="h-4 w-4" /> Download claim pack (zip)
                </Button>
                <Button variant="outline" onClick={() => window.open(URL.createObjectURL(new Blob([saved.summaryHtml], { type: 'text/html' })), '_blank')}>
                  <FileText className="h-4 w-4" /> Summary (print or save as PDF)
                </Button>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-5 lg:grid-cols-4">
              <StatTile label="Rows read" value={result.linesRead.toLocaleString()} />
              <StatTile label="In the claim" tone="good" value={claimedCount.toLocaleString()} caption={`${gbp(claimedTotal)} donated`} />
              <StatTile label="Gift Aid to reclaim" tone="good" value={gbp(result.totals.giftAid)} caption="25% of the claim" />
              <StatTile label="Held back for review" tone="bad" value={result.review.length.toLocaleString()} caption={`${gbp(result.totals.reviewDonations)}`} />
            </div>

            {saved.possibleRepeats.count > 0 && (
              <div className="flex gap-3 rounded-2xl border border-warn/20 bg-warn-soft p-5 text-sm text-warn">
                <AlertCircle className="h-5 w-5 flex-shrink-0" />
                <div>
                  <span className="font-medium">{saved.possibleRepeats.count.toLocaleString()} donation{saved.possibleRepeats.count === 1 ? '' : 's'} without a platform ID look identical</span>{' '}
                  to ones already claimed ({saved.possibleRepeats.claimRefs.join(', ')}). Check you are not claiming them twice.
                </div>
              </div>
            )}

            <Card className="p-6">
              <div className="mb-4 flex flex-wrap items-baseline justify-between gap-3">
                <h3 className="font-display text-xl font-semibold text-ink">HMRC schedules</h3>
                <Link href="/claims" className="font-medium text-brand hover:underline">Open Claims &rarr;</Link>
              </div>
              <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
                {saved.schedules.map((s, i) => {
                  const c = saved.claims[i]
                  return (
                    <div key={s.name} className="rounded-xl border border-line p-4">
                      <div className="flex items-start justify-between gap-3">
                        <Link href={`/claims/${c.id}`} className="text-ink hover:text-brand">{c.ref}</Link>
                        <Pill tone={c.kind === 'Overseas' ? 'info' : 'brand'}>{c.kind}</Pill>
                      </div>
                      <div className="mt-1 text-sm text-muted">{s.plan.name.replace(/\.ods$/i, '')} · {c.donationCount.toLocaleString()} donations · {gbp(c.totalAmount)}</div>
                      <Button variant="outline" size="sm" className="mt-3 h-auto w-full justify-start whitespace-normal py-2 text-left" onClick={() => download(s.name, s.data as BlobPart, ODS)}>
                        <FileSpreadsheet className="h-4 w-4 flex-shrink-0 text-good" /> {s.name}
                      </Button>
                    </div>
                  )
                })}
              </div>
              <div className="mt-5 flex flex-wrap items-center gap-3 border-t border-line pt-5">
                <Button variant="outline" onClick={() => download('2_Claim_workbook.xlsx', saved.workbook as BlobPart, XLSX_MIME)}>
                  <FileSpreadsheet className="h-4 w-4 text-info" /> Review workbook (needs review, not eligible, changes)
                </Button>
              </div>
            </Card>

            <Card className="p-6">
              <div className="flex flex-wrap items-center gap-3">
                <h3 className="font-display text-xl font-semibold text-ink">Simple output sheets</h3>
                <Pill tone="neutral">Optional</Pill>
              </div>
              <p className="mb-4 mt-1 text-muted">
                Not needed for HMRC. Plain Excel copies of each schedule&apos;s clean data, without HMRC&apos;s template. Handy as a backup or for your own records.
              </p>
              <div className="grid grid-cols-1 gap-2 md:grid-cols-2">
                {saved.simpleSheets.map((s) => (
                  <Button key={s.name} variant="outline" className="h-auto justify-start whitespace-normal py-3 text-left" onClick={() => download(s.name, s.data as BlobPart, XLSX_MIME)}>
                    <FileSpreadsheet className="h-4 w-4 flex-shrink-0 text-info" /> {s.name}
                  </Button>
                ))}
              </div>
              {saved.simpleSheets.length > 1 && (
                <Button
                  variant="outline"
                  className="mt-3"
                  onClick={async () => {
                    for (const s of saved.simpleSheets) {
                      download(s.name, s.data as BlobPart, XLSX_MIME)
                      await tick(300)
                    }
                  }}
                >
                  <Download className="h-4 w-4" /> Download all simple sheets
                </Button>
              )}
            </Card>

            {blocked.length > 0 && (
              <p className="text-sm text-muted">{blocked.length.toLocaleString()} donation{blocked.length === 1 ? ' was' : 's were'} left out because they were already claimed.</p>
            )}
          </>
        )}

        {phase === 'saved' && !saved && <Skeleton className="h-40 w-full" />}
      </div>
    </>
  )
}
