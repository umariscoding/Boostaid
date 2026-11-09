import type { ClaimDto } from '@/lib/claims'
import { Pill } from '@/components/ui/primitives'

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

/** "Oct - Nov 2026", "Jul 2026", or "Dec 2025 - Jan 2026" for a claim's donation dates. */
export function fmtPeriod(from?: string | null, to?: string | null): string {
  if (!from) return '-'
  const f = new Date(from + 'T00:00:00Z')
  const t = new Date((to ?? from) + 'T00:00:00Z')
  const month = (d: Date) => d.toLocaleDateString('en-GB', { month: 'short', timeZone: 'UTC' })
  const year = (d: Date) => d.getUTCFullYear()
  if (f.getUTCFullYear() === t.getUTCFullYear() && f.getUTCMonth() === t.getUTCMonth()) return `${month(f)} ${year(f)}`
  if (year(f) === year(t)) return `${month(f)} – ${month(t)} ${year(f)}`
  return `${month(f)} ${year(f)} – ${month(t)} ${year(t)}`
}

export const STATUS_LABEL: Record<Status, string> = {
  pending: 'Pending',
  submitted: 'Submitted',
  received: 'Received',
}

const STATUS_TONE: Record<Status, 'warn' | 'info' | 'good'> = {
  pending: 'warn',
  submitted: 'info',
  received: 'good',
}

export function StatusBadge({ status }: { status: Status }) {
  return <Pill tone={STATUS_TONE[status]}>{STATUS_LABEL[status]}</Pill>
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
