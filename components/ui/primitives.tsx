import Link from 'next/link'
import type { ReactNode } from 'react'
import { cn } from '@/lib/utils'

export const card = 'bg-white border border-line rounded-2xl'
export const inputClass =
  'w-full h-11 px-4 rounded-xl bg-white border border-line text-ink placeholder:text-muted/70 text-sm focus:outline-none focus:ring-2 focus:ring-brand/30 focus:border-brand/50'
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
const TONE_TEXT: Record<Tone, string> = {
  ink: 'text-ink',
  good: 'text-good',
  info: 'text-info',
  bad: 'text-bad',
  warn: 'text-warn',
  brand: 'text-brand',
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
  return (
    <div className={cn(card, 'p-6')}>
      <div className="text-base text-muted">{label}</div>
      <div className={cn('font-display text-4xl font-bold tracking-tight mt-1.5', TONE_TEXT[tone])}>{value}</div>
      {caption && <div className="text-base text-muted mt-2">{caption}</div>}
    </div>
  )
}

const PILL: Record<'neutral' | 'warn' | 'good' | 'bad' | 'info' | 'brand', string> = {
  neutral: 'bg-[#eef0f3] text-[#3b4252]',
  warn: 'bg-warn-soft text-warn',
  good: 'bg-good-soft text-good',
  bad: 'bg-bad-soft text-bad',
  info: 'bg-info-soft text-info',
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
