'use client'

import Link from 'next/link'
import { usePathname } from 'next/navigation'
import { Gift, PieChart, Plus, Receipt, Search, Settings, Sparkles, type LucideIcon } from 'lucide-react'
import { Select } from '@/components/ui/select'
import { useCharity } from '@/lib/store/charity'
import { Skeleton } from '@/components/ui/skeleton'

const NAV: { href: string; label: string; icon: LucideIcon }[] = [
  { href: '/', label: 'Clean', icon: Sparkles },
  { href: '/donors', label: 'Search', icon: Search },
  { href: '/claims', label: 'Claims', icon: Receipt },
  { href: '/gasds', label: 'GASDS', icon: PieChart },
  { href: '/settings', label: 'Settings', icon: Settings },
]

/** UK tax year runs 6 April to 5 April. */
function taxYear(now = new Date()): string {
  const start = now.getMonth() > 3 || (now.getMonth() === 3 && now.getDate() >= 6) ? now.getFullYear() : now.getFullYear() - 1
  return `${start}/${String(start + 1).slice(2)}`
}

async function logout() {
  await fetch('/api/auth/logout', { method: 'POST' })
  window.location.href = '/login'
}

function CharityPicker() {
  const { charities, loaded, selectedId, select } = useCharity()
  return (
    <div className="px-5 pb-3">
      <div className="text-sm text-side-muted mb-1.5">Charity</div>
      {!loaded ? (
        <Skeleton className="h-11 w-full rounded-xl bg-side-active" />
      ) : charities.length === 0 ? (
        <Link
          href="/settings"
          className="flex h-11 items-center gap-2 rounded-xl border border-dashed border-side-muted/50 px-4 text-[15px] text-white hover:bg-side-active"
        >
          <Plus className="w-4 h-4" /> Add a charity
        </Link>
      ) : (
        <Select variant="dark" aria-label="Charity" value={selectedId} onChange={(e) => select(e.target.value)}>
          <option value="all">All charities</option>
          {charities.map((c) => (
            <option key={c.id} value={c.id}>
              {c.name}
            </option>
          ))}
        </Select>
      )}
    </div>
  )
}

export default function Sidebar() {
  const pathname = usePathname()
  const isActive = (href: string) => (href === '/' ? pathname === '/' : pathname.startsWith(href))

  return (
    <aside className="bg-side text-white md:w-60 md:shrink-0 md:min-h-screen md:sticky md:top-0 md:h-screen flex flex-col">
      <div className="px-5 pt-6 md:pt-7 pb-3 md:pb-4 flex items-center gap-3">
        <div className="w-11 h-11 rounded-full bg-[#7c5cf0] flex items-center justify-center">
          <Gift className="w-5 h-5 text-white" />
        </div>
        <span className="font-display text-xl font-bold tracking-tight">BoostAid</span>
        <button onClick={logout} className="md:hidden ml-auto text-sm text-side-muted hover:text-white cursor-pointer">
          Log out
        </button>
      </div>

      <CharityPicker />

      <nav className="px-3 pb-3 md:pt-3 flex md:flex-col gap-1 overflow-x-auto">
        {NAV.map(({ href, label, icon: Icon }) => (
          <Link
            key={href}
            href={href}
            className={`flex shrink-0 items-center gap-2 md:gap-3 rounded-xl px-3 md:px-4 py-2.5 md:py-3 text-[15px] font-medium transition-colors ${
              isActive(href) ? 'bg-side-active text-white' : 'text-[#c9c9d8] hover:bg-side-active/60 hover:text-white'
            }`}
          >
            <Icon className="w-[18px] h-[18px]" />
            {label}
          </Link>
        ))}
      </nav>

      <div className="mt-auto px-5 py-5 text-sm text-side-muted space-y-1 hidden md:block">
        <div>Tax year {taxYear()}</div>
        <button
          onClick={logout}
          className="text-side-muted hover:text-white underline-offset-2 hover:underline cursor-pointer"
        >
          Log out
        </button>
      </div>
    </aside>
  )
}
