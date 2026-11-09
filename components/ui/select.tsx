import * as React from 'react'
import { ChevronDown } from 'lucide-react'
import { cn } from '@/lib/utils'

type Variant = 'light' | 'compact' | 'dark'

const SELECT: Record<Variant, string> = {
  light: 'h-11 rounded-full border border-line bg-white pl-4 pr-10 text-[15px] text-ink focus:ring-brand/30',
  compact: 'h-9 rounded-full border border-line bg-white pl-4 pr-10 text-sm text-ink focus:ring-brand/30',
  dark: 'h-11 w-full rounded-xl bg-side-active pl-4 pr-10 text-[15px] font-medium text-white focus:ring-[#7c5cf0]/60',
}
// The chevron is centred on the control's height and sits the same distance from the edge in every variant
const CHEVRON: Record<Variant, string> = {
  light: 'right-4 text-muted',
  compact: 'right-4 text-muted',
  dark: 'right-4 text-side-muted',
}

/** A native select with one consistent chevron, so every dropdown in the app lines up. */
export const Select = React.forwardRef<
  HTMLSelectElement,
  React.SelectHTMLAttributes<HTMLSelectElement> & { variant?: Variant; wrapperClassName?: string }
>(({ variant = 'light', className, wrapperClassName, children, ...props }, ref) => (
  <div className={cn('relative', variant === 'dark' ? 'block' : 'inline-block', wrapperClassName)}>
    <select
      ref={ref}
      className={cn(
        'block w-full appearance-none truncate focus:outline-none focus:ring-2 cursor-pointer disabled:cursor-not-allowed disabled:opacity-60',
        SELECT[variant],
        className
      )}
      {...props}
    >
      {children}
    </select>
    <ChevronDown
      aria-hidden
      className={cn('pointer-events-none absolute top-1/2 h-4 w-4 -translate-y-1/2', CHEVRON[variant])}
    />
  </div>
))
Select.displayName = 'Select'
