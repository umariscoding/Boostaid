'use client'

import { useEffect, useState } from 'react'
import { AlertTriangle, Loader2, Plus, Trash2 } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Card, PageHeader, inputClass, labelClass, th, thr, td, tdr } from '@/components/ui/primitives'
import { Dialog } from '@/components/ui/dialog'
import { TableSkeleton } from '@/components/ui/skeleton'
import { apiJson, fmtDay, gbp } from '@/components/claims/shared'
import { useCharity, type CharityInfo } from '@/components/charity/charity-context'

export default function SettingsPage() {
  const { charities, loaded, selectedId, select, refresh } = useCharity()
  const [name, setName] = useState('')
  const [adding, setAdding] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [notice, setNotice] = useState<string | null>(null)
  const [toDelete, setToDelete] = useState<CharityInfo | null>(null)
  const [understood, setUnderstood] = useState(false)
  const [deleting, setDeleting] = useState(false)

  // Counts can be stale (uploads happen on another page), so refresh when this page opens
  useEffect(() => {
    refresh().catch(() => {})
  }, [refresh])

  // Keep the dialog's numbers current if the list refreshes while it is open
  const target = toDelete ? charities.find((c) => c.id === toDelete.id) ?? toDelete : null
  const hasRecords = !!target && (target.claimCount > 0 || target.donationCount > 0)

  const add = async (e: React.FormEvent) => {
    e.preventDefault()
    setError(null)
    setNotice(null)
    setAdding(true)
    try {
      const { charity } = await apiJson<{ charity: CharityInfo }>('/api/charities', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name }),
      })
      await refresh()
      // The first charity (or any new one) becomes the one you're working in
      select(charity.id)
      setName('')
      setNotice(`Added ${charity.name}. It's now selected in the sidebar.`)
    } catch (err) {
      setError((err as Error).message)
    } finally {
      setAdding(false)
    }
  }

  const openDelete = (c: CharityInfo) => {
    setError(null)
    setNotice(null)
    setUnderstood(false)
    setToDelete(c)
    refresh().catch(() => {})
  }

  const confirmDelete = async () => {
    if (!target) return
    setDeleting(true)
    try {
      const res = await apiJson<{ deleted: { name: string; claims: number; donations: number } }>(`/api/charities/${target.id}`, { method: 'DELETE' })
      if (selectedId === target.id) select('all')
      await refresh()
      setToDelete(null)
      const { name: n, claims, donations } = res.deleted
      setNotice(
        claims || donations
          ? `Deleted ${n} along with ${claims.toLocaleString()} claim${claims === 1 ? '' : 's'} and ${donations.toLocaleString()} donation${donations === 1 ? '' : 's'}.`
          : `Deleted ${n}.`
      )
    } catch (err) {
      setError((err as Error).message)
      setToDelete(null)
    } finally {
      setDeleting(false)
    }
  }

  return (
    <>
      <PageHeader title="Settings" subtitle="Manage the charities you claim Gift Aid for." />

      <Card className="overflow-hidden">
        <div className="p-6 border-b border-line">
          <h2 className="font-display text-2xl font-semibold text-ink">Manage charities</h2>
          <p className="text-muted mt-1">
            Each upload, claim and donor belongs to one charity. Use the sidebar to switch between them or view all together.
          </p>

          <form onSubmit={add} className="mt-5 flex flex-wrap items-end gap-3">
            <div className="grow max-w-md">
              <label className={labelClass} htmlFor="charity-name">Charity name</label>
              <input
                id="charity-name"
                className={inputClass}
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="e.g. Example Community Trust"
                maxLength={100}
                autoComplete="off"
              />
            </div>
            <Button type="submit" disabled={adding || !name.trim()}>
              {adding ? <Loader2 className="w-4 h-4 animate-spin" /> : <Plus className="w-4 h-4" />}
              Add charity
            </Button>
          </form>

          {error && <div className="mt-4 rounded-xl border border-bad/20 bg-bad-soft px-4 py-3 text-sm text-bad">{error}</div>}
          {notice && <div className="mt-4 rounded-xl border border-good/20 bg-good-soft px-4 py-3 text-sm text-good">{notice}</div>}
        </div>

        {!loaded ? (
          <TableSkeleton
            columns={[
              { label: 'Charity' },
              { label: 'Claims', right: true },
              { label: 'Donations', right: true },
              { label: 'Donated', right: true },
              { label: 'Added' },
              { label: '', skeleton: true },
            ]}
            rows={3}
          />
        ) : charities.length === 0 ? (
          <div className="px-6 py-14 text-center">
            <div className="font-display text-xl font-semibold text-ink">No charities yet</div>
            <p className="text-muted mt-1">Add your first charity above to start uploading donations.</p>
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-[15px]">
              <thead>
                <tr className="border-b border-line">
                  <th className={th}>Charity</th>
                  <th className={thr}>Claims</th>
                  <th className={thr}>Donations</th>
                  <th className={thr}>Donated</th>
                  <th className={th}>Added</th>
                  <th className={th} />
                </tr>
              </thead>
              <tbody>
                {charities.map((c) => (
                  <tr key={c.id} className="border-b border-line last:border-0">
                    <td className={`${td} text-ink`}>{c.name}</td>
                    <td className={`${tdr} text-ink`}>{c.claimCount.toLocaleString()}</td>
                    <td className={`${tdr} text-ink`}>{c.donationCount.toLocaleString()}</td>
                    <td className={`${tdr} text-ink`}>{gbp(c.donated)}</td>
                    <td className={`${td} text-ink whitespace-nowrap`}>{fmtDay(c.createdAt)}</td>
                    <td className={`${td} text-right`}>
                      <Button variant="danger" size="sm" onClick={() => openDelete(c)}>
                        <Trash2 className="w-4 h-4" /> Delete
                      </Button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      <Dialog open={!!target} onClose={() => !deleting && setToDelete(null)}>
        {target && (
          <>
            <div className="flex items-start gap-4">
              <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-full bg-bad-soft">
                <AlertTriangle className="h-5 w-5 text-bad" />
              </div>
              <div>
                <h2 className="font-display text-xl font-semibold text-ink">Delete {target.name}?</h2>
                {hasRecords ? (
                  <>
                    <p className="mt-2 text-ink">
                      This charity has <strong className="font-medium">{target.claimCount.toLocaleString()} claim{target.claimCount === 1 ? '' : 's'}</strong>{' '}
                      and <strong className="font-medium">{target.donationCount.toLocaleString()} donation{target.donationCount === 1 ? '' : 's'}</strong>{' '}
                      ({gbp(target.donated)} donated).
                    </p>
                    <p className="mt-2 rounded-xl bg-bad-soft px-4 py-3 text-bad">
                      If you delete this charity, all of its records will also be deleted: its uploads, claims, donors and status history.
                      This can&apos;t be undone.
                    </p>
                    <label className="mt-4 flex cursor-pointer items-start gap-3 text-ink">
                      <input
                        type="checkbox"
                        checked={understood}
                        onChange={(e) => setUnderstood(e.target.checked)}
                        className="mt-1 h-4 w-4 accent-[#9a2b1d]"
                      />
                      <span>I understand these records will be permanently deleted.</span>
                    </label>
                  </>
                ) : (
                  <p className="mt-2 text-muted">This charity has no records, so nothing else will be removed.</p>
                )}
              </div>
            </div>
            <div className="mt-6 flex justify-end gap-3">
              <Button variant="outline" onClick={() => setToDelete(null)} disabled={deleting}>
                Cancel
              </Button>
              <Button
                onClick={confirmDelete}
                disabled={deleting || (hasRecords && !understood)}
                className="bg-bad text-white hover:bg-[#832316]"
              >
                {deleting && <Loader2 className="w-4 h-4 animate-spin" />}
                {hasRecords ? 'Delete charity and records' : 'Delete charity'}
              </Button>
            </div>
          </>
        )}
      </Dialog>
    </>
  )
}
