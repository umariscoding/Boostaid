'use client'

import { use, useEffect, useState } from 'react'

interface Open {
  state: 'open'
  charity: string
  declaration: string
  twoPeople: boolean
  donations: { date: string | null; amount: number; giftAid: number }[]
}
type Info = Open | { state: 'expired' | 'used' | 'done' | 'invalid'; charity?: string; contactEmail?: string | null }

const gbp = (n: number) => '£' + n.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
const ukDate = (iso: string | null) => (iso ? iso.split('-').reverse().join('/') : '')
const field = 'h-12 w-full rounded-xl border border-[#cfcadf] bg-white px-3 text-base text-[#17152B] outline-none focus:border-[#4C1D95] focus:ring-2 focus:ring-[#4C1D95]/20'

/** What the donor sees (mobile first). No login: the link's secret code is the key. */
export default function ConfirmPage({ params }: { params: Promise<{ code: string }> }) {
  const { code } = use(params)
  const [info, setInfo] = useState<Info | null>(null)
  const [f, setF] = useState({ title: '', first: '', last: '', house: '', postcode: '', declaration: false, all: false })
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)
  const [thanks, setThanks] = useState<{ first: string } | null>(null)

  useEffect(() => {
    fetch(`/api/confirm/${encodeURIComponent(code)}`).then(async (r) => setInfo((await r.json().catch(() => ({ state: 'invalid' }))) as Info), () => setInfo({ state: 'invalid' }))
  }, [code])

  const submit = async (e: React.FormEvent) => {
    e.preventDefault()
    setError('')
    setBusy(true)
    try {
      const r = await fetch(`/api/confirm/${encodeURIComponent(code)}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ title: f.title, first: f.first, last: f.last, house: f.house, postcode: f.postcode, declaration: f.declaration, all: f.all }),
      })
      const d = await r.json().catch(() => ({}))
      if (!r.ok) {
        if (d.state) setInfo({ state: d.state })
        setError(d.error ?? 'Something went wrong. Please try again.')
      } else setThanks({ first: d.first })
    } finally {
      setBusy(false)
    }
  }

  const charity = info && 'charity' in info && info.charity ? info.charity : ''
  const total = info?.state === 'open' ? info.donations.reduce((s, d) => s + d.giftAid, 0) : 0

  return (
    <div className="min-h-screen bg-[#F7F6FB] px-4 py-8 text-[#17152B]">
      <div className="mx-auto w-full max-w-md overflow-hidden rounded-3xl border border-[#E3E0EC] bg-white shadow-sm">
        <div className="flex items-center gap-3 bg-[#4C1D95] px-6 py-5 text-white">
          <div className="grid h-10 w-10 place-items-center rounded-full bg-white/20 text-lg font-bold">{(charity || 'B')[0]}</div>
          <div>
            <div className="font-bold">{charity || 'Confirm your Gift Aid'}</div>
            <div className="text-sm opacity-85">Confirm your Gift Aid</div>
          </div>
        </div>
        <div className="space-y-4 p-6">
          {!info && <p className="text-[#3D3A55]">Loading…</p>}

          {thanks && (
            <div className="space-y-2 py-4 text-center">
              <div className="mx-auto grid h-14 w-14 place-items-center rounded-full bg-[#E7F6EC] text-2xl text-[#0F5A30]">✓</div>
              <h1 className="text-xl font-bold">Thank you, {thanks.first}</h1>
              <p className="text-[#3D3A55]">{charity} can now claim Gift Aid on your donation. You don’t need to do anything else. This link has now been used.</p>
            </div>
          )}

          {!thanks && info && info.state !== 'open' && (
            <div className="space-y-2 py-4 text-center">
              <h1 className="text-xl font-bold">
                {info.state === 'used' ? 'This link has already been used' : info.state === 'done' ? 'Thank you, this is already sorted' : info.state === 'expired' ? 'This link has expired' : 'This link is not valid'}
              </h1>
              <p className="text-[#3D3A55]">
                {info.state === 'done'
                  ? `${charity} already has everything it needs for this donation. Nothing more to do.`
                  : info.contactEmail
                    ? `Please contact ${charity} at ${info.contactEmail}.`
                    : charity
                      ? `Please contact ${charity}.`
                      : 'Please check the link in your email, or contact the charity.'}
              </p>
            </div>
          )}

          {!thanks && info?.state === 'open' && (
            <form onSubmit={submit} className="space-y-4" noValidate>
              <div className="rounded-xl bg-[#F1EFF6] px-4 py-3 text-sm">
                {info.donations.map((d, i) => (
                  <div key={i}>Your donation: <b>{gbp(d.amount)}</b> on {ukDate(d.date)}</div>
                ))}
                <div className="mt-1 text-[#3D3A55]">Gift Aid adds {gbp(total)} at no cost to you.</div>
              </div>
              {info.twoPeople && <div className="rounded-xl bg-[#FFF1DB] px-4 py-3 text-sm text-[#7A4300]">Gift Aid needs one taxpayer. Enter the name of the person who pays the tax.</div>}
              <div className="grid grid-cols-[110px_1fr] gap-3">
                <label className="space-y-1 text-sm font-semibold">Title
                  <select className={field} value={f.title} onChange={(e) => setF({ ...f, title: e.target.value })}>
                    {['', 'Mr', 'Mrs', 'Ms', 'Miss', 'Dr'].map((t) => <option key={t}>{t}</option>)}
                  </select>
                </label>
                <label className="space-y-1 text-sm font-semibold">First name
                  <input className={field} value={f.first} onChange={(e) => setF({ ...f, first: e.target.value })} autoComplete="given-name" maxLength={100} />
                </label>
              </div>
              <label className="block space-y-1 text-sm font-semibold">Last name
                <input className={field} value={f.last} onChange={(e) => setF({ ...f, last: e.target.value })} autoComplete="family-name" maxLength={100} />
              </label>
              <div className="grid grid-cols-2 gap-3">
                <label className="space-y-1 text-sm font-semibold">House name or number
                  <input className={field} value={f.house} onChange={(e) => setF({ ...f, house: e.target.value })} autoComplete="off" maxLength={100} />
                </label>
                <label className="space-y-1 text-sm font-semibold">Postcode
                  <input className={field} value={f.postcode} onChange={(e) => setF({ ...f, postcode: e.target.value })} autoComplete="postal-code" maxLength={12} />
                </label>
              </div>
              <label className="flex gap-3 rounded-xl border border-[#E3E0EC] p-3 text-sm">
                <input type="checkbox" className="mt-1 h-5 w-5 shrink-0" checked={f.declaration} onChange={(e) => setF({ ...f, declaration: e.target.checked })} />
                <span>{info.declaration}</span>
              </label>
              <label className="flex gap-3 rounded-xl border border-[#E3E0EC] p-3 text-sm">
                <input type="checkbox" className="mt-1 h-5 w-5 shrink-0" checked={f.all} onChange={(e) => setF({ ...f, all: e.target.checked })} />
                <span>Also cover my past and future donations. (Optional)</span>
              </label>
              {error && <div role="alert" className="text-sm font-semibold text-[#9B1C13]">{error}</div>}
              <button type="submit" disabled={busy} className="h-12 w-full rounded-full bg-[#4C1D95] text-base font-semibold text-white disabled:opacity-60">
                {busy ? 'Confirming…' : 'Confirm'}
              </button>
              <p className="text-center text-xs text-[#5B5874]">Your details are only used to claim Gift Aid on your donation.</p>
            </form>
          )}
        </div>
      </div>
    </div>
  )
}
