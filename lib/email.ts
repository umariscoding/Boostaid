/**
 * Sending email. Version 1 sends through BoostAid with Postmark (the charity's name as the sender, its own address as reply-to).
 * With no provider configured nothing leaves BoostAid: `delivered` is false, and callers say so.
 *
 *   POSTMARK_SERVER_TOKEN   Postmark server token
 *   EMAIL_FROM_ADDRESS      a verified sender address, e.g. no-reply@boostaid.co.uk
 */
export interface Mail {
  to: string
  /** Shown as the sender name, e.g. "Lonely Orphans (sent by BoostAid)". */
  fromName: string
  replyTo?: string | null
  subject: string
  html: string
  text: string
  attachments?: { name: string; contentType: string; data: Uint8Array }[]
}

export const emailConfigured = () => !!process.env.POSTMARK_SERVER_TOKEN && !!process.env.EMAIL_FROM_ADDRESS

const clean = (s: string) => s.replace(/[\r\n"<>]/g, ' ').trim()

export async function sendMail(m: Mail): Promise<{ delivered: boolean }> {
  if (!emailConfigured()) return { delivered: false }
  const res = await fetch('https://api.postmarkapp.com/email', {
    method: 'POST',
    headers: { Accept: 'application/json', 'Content-Type': 'application/json', 'X-Postmark-Server-Token': process.env.POSTMARK_SERVER_TOKEN! },
    body: JSON.stringify({
      From: `"${clean(m.fromName)}" <${process.env.EMAIL_FROM_ADDRESS}>`,
      To: m.to,
      ReplyTo: m.replyTo || undefined,
      Subject: m.subject.replace(/[\r\n]/g, ' '),
      HtmlBody: m.html,
      TextBody: m.text,
      MessageStream: 'outbound',
      Attachments: m.attachments?.map((a) => ({ Name: a.name, Content: Buffer.from(a.data).toString('base64'), ContentType: a.contentType })),
    }),
  })
  if (!res.ok) throw new Error(`The email service refused the message (${res.status})`)
  return { delivered: true }
}

/** The site's own address, for links in emails. Set APP_URL in production; otherwise the request's address is used. */
export function appBase(requestOrigin?: string): string {
  const env = process.env.APP_URL?.replace(/\/$/, '')
  if (env) return env
  if (requestOrigin) return requestOrigin.replace(/\/$/, '')
  if (process.env.VERCEL_PROJECT_PRODUCTION_URL) return `https://${process.env.VERCEL_PROJECT_PRODUCTION_URL}`
  return 'http://localhost:3000'
}

export const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
