import { ChevronLeft, ChevronRight } from 'lucide-react'
import { Select } from '@/components/ui/select'

export const DEFAULT_PAGE_SIZES = [10, 25, 50, 100]

/** Card footer: "Showing 1-25 of 490", a page-size picker and previous/next. */
export function Pagination({
  page,
  pageSize,
  total,
  onPage,
  onPageSize,
  sizes = DEFAULT_PAGE_SIZES,
}: {
  page: number
  pageSize: number
  total: number
  onPage: (page: number) => void
  onPageSize?: (size: number) => void
  sizes?: number[]
}) {
  if (total <= 0) return null
  const pages = Math.max(1, Math.ceil(total / pageSize))
  const from = (page - 1) * pageSize + 1
  const to = Math.min(page * pageSize, total)
  // Nothing to paginate when everything fits on one page at the smallest size
  if (total <= Math.min(...sizes)) return null

  const nav = 'inline-flex h-9 w-9 items-center justify-center rounded-full border border-line bg-white text-ink hover:bg-wash disabled:opacity-40 disabled:pointer-events-none cursor-pointer'

  return (
    <div className="flex flex-wrap items-center justify-between gap-3 border-t border-line px-5 py-3 text-sm text-muted">
      <span>
        Showing {from.toLocaleString()}-{to.toLocaleString()} of {total.toLocaleString()}
      </span>
      <div className="flex items-center gap-4">
        {onPageSize && (
          <label className="flex items-center gap-2">
            Rows per page
            <Select variant="compact" value={pageSize} onChange={(e) => onPageSize(Number(e.target.value))}>
              {sizes.map((n) => (
                <option key={n} value={n}>
                  {n}
                </option>
              ))}
            </Select>
          </label>
        )}
        <div className="flex items-center gap-2">
          <button className={nav} onClick={() => onPage(page - 1)} disabled={page <= 1} aria-label="Previous page">
            <ChevronLeft className="w-4 h-4" />
          </button>
          <span className="min-w-[84px] text-center text-ink">
            Page {page} of {pages.toLocaleString()}
          </span>
          <button className={nav} onClick={() => onPage(page + 1)} disabled={page >= pages} aria-label="Next page">
            <ChevronRight className="w-4 h-4" />
          </button>
        </div>
      </div>
    </div>
  )
}
