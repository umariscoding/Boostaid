import Link from 'next/link'
import type { ReactNode } from 'react'
import { cn } from '@/lib/utils'

export const card = 'bg-white border border-line rounded-2xl'
export const inputClass =
  'w-full h-11 px-4 rounded-xl bg-white border border-line-strong text-ink placeholder:text-muted/70 text-sm focus:outline-none focus:ring-2 focus:ring-brand/30 focus:border-brand/50'
export const labelClass = 'block text-sm font-medium text-ink mb-1.5'
export const th = 'px-5 py-3.5 text-left text-sm font-medium text-muted whitespace-nowrap'
/** Right-aligned header for number columns. */
export const thr = 'px-5 py-3.5 text-right text-sm font-medium text-muted whitespace-nowrap'
export const td = 'px-5 py-4 align-middle'
/** Right-aligned, fixed-width digits so amounts line up down the column. */
export const tdr = 'px-5 py-4 align-middle text-right tabular-nums whitespace-nowrap'

export function Card({ className, children }: { className?: string; children: ReactNode }) {
  return <div className={cn(card, className)}>{children}</div>
}

type Tone = 'ink' | 'good' | 'info' | 'bad' | 'warn' | 'brand'

// Plain white card for neutral figures; a tinted card (like the design) for status figures.
const TILE: Record<Tone, { box: string; label: string; value: string; caption: string }> = {
  ink: { box: 'bg-white border-line', label: 'text-muted', value: 'text-ink', caption: 'text-muted' },
  brand: { box: 'bg-white border-line', label: 'text-muted', value: 'text-brand', caption: 'text-muted' },
  good: { box: 'bg-good-soft border-good-line', label: 'text-good', value: 'text-good', caption: 'text-good' },
  info: { box: 'bg-info-soft border-info-line', label: 'text-info', value: 'text-info', caption: 'text-info' },
  warn: { box: 'bg-warn-soft border-warn-line', label: 'text-warn', value: 'text-warn', caption: 'text-warn' },
  bad: { box: 'bg-bad-soft border-bad-line', label: 'text-bad', value: 'text-bad', caption: 'text-bad' },
}

export function StatTile({
  label,
  value,
  caption,
  tone = 'ink',
}: {
  label: string
  value: ReactNode
  caption?: ReactNode
  tone?: Tone
}) {
  const t = TILE[tone]
  return (
    <div className={cn('h-full rounded-2xl border p-5', t.box)}>
      <div className={cn('text-base font-medium', t.label)}>{label}</div>
      <div className={cn('mt-1.5 font-display text-4xl font-bold tracking-tight', t.value)}>{value}</div>
      {caption && <div className={cn('mt-1.5 text-base', t.caption)}>{caption}</div>}
    </div>
  )
}

const PILL: Record<'neutral' | 'warn' | 'good' | 'bad' | 'info' | 'brand', string> = {
  neutral: 'bg-[#f1eff6] text-muted',
  warn: 'bg-warn-pill text-warn',
  good: 'bg-good-pill text-good',
  bad: 'bg-bad-pill text-bad',
  info: 'bg-info-pill text-info',
  brand: 'bg-brand-soft text-brand',
}

export function Pill({
  tone = 'neutral',
  children,
  className,
}: {
  tone?: keyof typeof PILL
  children: ReactNode
  className?: string
}) {
  return (
    <span className={cn('inline-flex items-center rounded-full px-3 py-1 text-sm font-medium whitespace-nowrap', PILL[tone], className)}>
      {children}
    </span>
  )
}

/** Page title row. The right-hand action defaults to the "Upload new file" shortcut. */
export function PageHeader({ title, subtitle, action }: { title: string; subtitle?: ReactNode; action?: ReactNode }) {
  return (
    <div className="flex flex-wrap items-start justify-between gap-4 mb-8">
      <div>
        <h1 className="font-display text-4xl font-bold tracking-tight text-ink">{title}</h1>
        {subtitle && <p className="text-muted mt-2">{subtitle}</p>}
      </div>
      {action ?? (
        <Link
          href="/"
          className="inline-flex h-12 items-center rounded-full bg-ink px-6 text-base font-medium text-white hover:bg-black"
        >
          Upload new file
        </Link>
      )}
    </div>
  )
}

/** Underlined tabs, as in the design. */
export function Tabs<T extends string>({
  tabs,
  value,
  onChange,
}: {
  tabs: { id: T; label: ReactNode }[]
  value: T
  onChange: (id: T) => void
}) {
  return (
    <div role="tablist" className="flex gap-2 border-b border-line">
      {tabs.map((t) => (
        <button
          key={t.id}
          role="tab"
          aria-selected={value === t.id}
          onClick={() => onChange(t.id)}
          className={cn(
            '-mb-px h-11 cursor-pointer border-b-[3px] px-4 text-base transition-colors',
            value === t.id ? 'border-brand font-semibold text-brand' : 'border-transparent font-medium text-muted hover:text-ink'
          )}
        >
          {t.label}
        </button>
      ))}
    </div>
  )
}
