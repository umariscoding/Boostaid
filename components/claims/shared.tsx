import type { ClaimDto } from '@/lib/claims'

export type Status = ClaimDto['status']

export const gbp = (n: number) =>
  '£' + n.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })

export const fmtDateTime = (iso: string | null) =>
  iso
    ? new Date(iso).toLocaleString('en-GB', { timeZone: 'Europe/London', dateStyle: 'medium', timeStyle: 'short' })
    : '-'

export const fmtDay = (iso: string | null) =>
  iso ? new Date(iso).toLocaleDateString('en-GB', { timeZone: 'Europe/London', dateStyle: 'medium' }) : '-'

/** yyyy-MM-dd (a plain date, no timezone) -> dd/MM/yyyy */
export const fmtPlainDate = (d: string | null) => (d ? d.split('-').reverse().join('/') : '-')

const STYLES: Record<Status, string> = {
  pending: 'bg-amber-500/15 text-amber-200 border-amber-500/30',
  submitted: 'bg-sky-500/15 text-sky-200 border-sky-500/30',
  received: 'bg-emerald-500/15 text-emerald-200 border-emerald-500/30',
}
const LABELS: Record<Status, string> = { pending: 'Pending', submitted: 'Submitted', received: 'Received' }

export function StatusBadge({ status }: { status: Status }) {
  return (
    <span className={`inline-flex items-center rounded-full border px-2.5 py-0.5 text-xs font-medium ${STYLES[status]}`}>
      {LABELS[status]}
    </span>
  )
}

export const NEXT_STATUS: Record<Status, { to: Status; label: string } | null> = {
  pending: { to: 'submitted', label: 'Mark submitted' },
  submitted: { to: 'received', label: 'Mark received' },
  received: null,
}

export const PREV_STATUS: Record<Status, Status | null> = {
  pending: null,
  submitted: 'pending',
  received: 'submitted',
}

export async function apiJson<T>(url: string, init?: RequestInit): Promise<T> {
  const res = await fetch(url, init)
  if (res.status === 401) {
    window.location.href = '/login'
    throw new Error('Session expired')
  }
  const data = await res.json().catch(() => null)
  if (!res.ok) throw new Error(data?.error || `Request failed (${res.status})`)
  return data as T
}

export const cardClass =
  'bg-slate-800/50 backdrop-blur-sm rounded-2xl shadow-xl border border-slate-700/50'
export const inputClass =
  'w-full h-10 px-3 rounded-lg bg-slate-700/50 border border-slate-600/50 text-white placeholder-slate-500 text-sm focus:outline-none focus:ring-2 focus:ring-purple-500/50 focus:border-purple-500/50'
