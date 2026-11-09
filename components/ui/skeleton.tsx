import { cn } from '@/lib/utils'
import { Button } from '@/components/ui/button'
import { StatTile, card, th, thr, td } from '@/components/ui/primitives'

/** A pulsing placeholder block. Size it with className. */
export function Skeleton({ className }: { className?: string }) {
  return <div aria-hidden className={cn('animate-pulse rounded-md bg-line', className)} />
}

// Deterministic widths so rows look natural but never jump between renders
const WIDTHS = ['w-3/4', 'w-2/3', 'w-5/6', 'w-1/2', 'w-4/5', 'w-3/5']

export interface SkeletonColumn {
  label?: string
  right?: boolean
  /** Any CSS width, e.g. '20%' or '3rem'. Leave out to share the rest. */
  width?: string
  /** Show a placeholder in this column. Defaults to true unless the label is an empty string. */
  skeleton?: boolean
}

/** Real table headers with placeholder rows, so nothing shifts when the data arrives. */
export function TableSkeleton({ columns, rows = 6, minWidth }: { columns: SkeletonColumn[]; rows?: number; minWidth?: number }) {
  return (
    <div role="status" aria-busy="true" className="overflow-x-auto">
      <span className="sr-only">Loading</span>
      <table className="w-full table-fixed text-base" style={minWidth ? { minWidth } : undefined}>
        <colgroup>
          {columns.map((c, i) => (
            <col key={i} style={c.width ? { width: c.width } : undefined} />
          ))}
        </colgroup>
        <thead>
          <tr className="border-b border-line">
            {columns.map((c, i) => (
              <th key={i} className={c.right ? thr : th}>
                {c.label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {Array.from({ length: rows }).map((_, r) => (
            <tr key={r} className="border-b border-line last:border-0">
              {columns.map((c, i) => (
                <td key={i} className={cn(td, 'py-[1.15rem]')}>
                  {(c.skeleton ?? c.label !== '') ? (
                    <Skeleton className={cn('h-4', c.right ? 'ml-auto w-3/4' : WIDTHS[(r + i * 2) % WIDTHS.length])} />
                  ) : null}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

/**
 * Placeholder for the Clean review screen while the numbers are worked out. Labels, captions, the buttons and
 * table headers are real; only the values (counts, file name, rows) are placeholders.
 */
export function ReviewSkeleton({ message }: { message?: string }) {
  const value = <Skeleton className="h-12 w-24" />
  return (
    <div className="space-y-6" role="status" aria-busy="true">
      <span className="sr-only">{message ?? 'Loading'}</span>
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-5">
        <StatTile label="Rows uploaded" value={value} caption={<Skeleton className="h-4 w-20" />} />
        <StatTile label="Ready for HMRC" tone="good" value={value} caption="All checks passed" />
        <StatTile label="Auto-fixed" tone="info" value={value} caption="House numbers, names, characters" />
        <StatTile label="To fix" tone="bad" value={value} caption="Need a decision" />
      </div>
      <div className="border-2 border-dashed border-[#cfc9e8] bg-[#faf9ff] rounded-2xl p-6 flex flex-wrap items-center justify-between gap-4">
        <div>
          <Skeleton className="h-7 w-64" />
          <Skeleton className="h-4 w-80 mt-3" />
        </div>
        <div className="flex gap-3">
          <Button variant="outline" size="lg" disabled>Change column mapping</Button>
          <Button size="lg" disabled>Save claim &amp; make files</Button>
        </div>
      </div>
      <div>
        <h2 className="font-display text-2xl font-semibold text-ink mb-3">Rows to fix</h2>
        <div className={cn(card, 'overflow-hidden')}>
          <TableSkeleton
            columns={[{ label: 'Reason' }, { label: 'Donor' }, { label: 'House' }, { label: 'Postcode' }, { label: 'Amount', right: true }]}
            rows={6}
          />
        </div>
      </div>
      {message && <p className="text-center text-muted">{message}</p>}
    </div>
  )
}
