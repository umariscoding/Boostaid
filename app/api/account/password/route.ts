import { NextResponse } from 'next/server'
import { eq } from 'drizzle-orm'
import { getDb, users } from '@/lib/db'
import { withAuth, currentUserId } from '@/lib/api'
import { hashPassword, passwordProblem, verifyPassword } from '@/lib/password'

/** Change your own password. Body: { current, next } */
export const POST = withAuth(async (req) => {
  const body = await req.json().catch(() => null)
  const current = String(body?.current ?? '')
  const next = String(body?.next ?? '')

  const id = currentUserId(req)
  const db = getDb()
  const [user] = id ? await db.select().from(users).where(eq(users.id, id)) : []
  if (!user) return NextResponse.json({ error: 'Unauthorised' }, { status: 401 })

  if (!(await verifyPassword(current, user.passwordHash))) {
    return NextResponse.json({ error: 'Your current password is not right' }, { status: 400 })
  }
  const problem = passwordProblem(next, user.email)
  if (problem) throw new RangeError(problem)
  if (next === current) throw new RangeError('Choose a different password from the current one')

  await db.update(users).set({ passwordHash: await hashPassword(next), failedAttempts: 0, lockedUntil: null }).where(eq(users.id, user.id))
  return NextResponse.json({ success: true })
})
