'use client'

import Link from 'next/link'
import { usePathname } from 'next/navigation'
import { FileSpreadsheet, ListChecks, LogOut, Search, Upload } from 'lucide-react'
import { Button } from '@/components/ui/button'

const NAV = [
  { href: '/', label: 'Process', icon: Upload },
  { href: '/claims', label: 'Claims', icon: ListChecks },
  { href: '/donors', label: 'Donor Search', icon: Search },
]

export default function AppHeader() {
  const pathname = usePathname()
  const isActive = (href: string) => (href === '/' ? pathname === '/' : pathname.startsWith(href))

  return (
    <div className="bg-slate-800 border-b border-slate-700/50 text-white">
      <div className="max-w-7xl mx-auto px-6 py-4">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="w-11 h-11 bg-gradient-to-br from-violet-500 to-purple-600 rounded-xl flex items-center justify-center shadow-lg shadow-purple-500/20">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <h1 className="text-xl font-semibold tracking-tight">BoostAid</h1>
              <p className="text-slate-400 text-sm">HMRC Gift Aid Submission Tool</p>
            </div>
          </div>

          <nav className="flex items-center gap-1 order-3 w-full md:order-none md:w-auto">
            {NAV.map(({ href, label, icon: Icon }) => (
              <Link
                key={href}
                href={href}
                className={`flex items-center gap-2 px-3.5 py-2 rounded-lg text-sm font-medium transition-colors ${
                  isActive(href)
                    ? 'bg-purple-500/20 text-purple-200 border border-purple-500/30'
                    : 'text-slate-400 hover:text-white hover:bg-slate-700/50 border border-transparent'
                }`}
              >
                <Icon className="w-4 h-4" />
                {label}
              </Link>
            ))}
          </nav>

          <Button
            onClick={async () => {
              await fetch('/api/auth/logout', { method: 'POST' })
              window.location.href = '/login'
            }}
            variant="outline"
            size="sm"
            className="border-slate-600 bg-slate-700/50 text-slate-300 hover:bg-slate-600 hover:text-white"
          >
            <LogOut className="w-4 h-4 mr-2" />
            Logout
          </Button>
        </div>
      </div>
    </div>
  )
}
