import Link from 'next/link'
import { Building2 } from 'lucide-react'

/** Shown wherever charity data is needed but no charity exists yet. */
export default function NoCharityNotice({ className }: { className?: string }) {
  return (
    <div className={`bg-white border border-dashed border-[#cfc9e8] rounded-2xl p-10 text-center ${className ?? ''}`}>
      <div className="w-14 h-14 mx-auto mb-4 rounded-full bg-brand-soft flex items-center justify-center">
        <Building2 className="w-6 h-6 text-brand" />
      </div>
      <div className="font-display text-xl font-semibold text-ink">Add a charity to get started</div>
      <p className="text-muted mt-1 max-w-md mx-auto">
        Every upload, claim and donor belongs to a charity. Add your first charity, then clean a spreadsheet for it.
      </p>
      <Link
        href="/settings"
        className="mt-5 inline-flex h-11 items-center rounded-full bg-brand px-6 text-sm font-medium text-white hover:bg-brand-dark"
      >
        Add a charity
      </Link>
    </div>
  )
}
