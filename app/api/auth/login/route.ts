import { NextRequest, NextResponse } from 'next/server'
import { eq, sql } from 'drizzle-orm'
import { getDb, users } from '@/lib/db'
import { burnPasswordCheck, verifyPassword } from '@/lib/password'
import { SESSION_COOKIE, SESSION_MAX_AGE, createSessionToken } from '@/lib/session'

const MAX_ATTEMPTS = 5
const LOCK_MINUTES = 15
const BAD = { success: false, error: 'Invalid email or password' }

export async function POST(request: NextRequest) {
  try {
    const body = await request.json().catch(() => null)
    const email = String(body?.email ?? '').trim().toLowerCase()
    const password = String(body?.password ?? '')
    if (!email || !password || password.length > 200) return NextResponse.json(BAD, { status: 401 })

    const db = getDb()
    const [user] = await db.select().from(users).where(sql`lower(${users.email}) = ${email}`)

    if (!user) {
      await burnPasswordCheck(password) // same work as a real check, so timing doesn't reveal unknown emails
      return NextResponse.json(BAD, { status: 401 })
    }

    if (user.lockedUntil && user.lockedUntil > new Date()) {
      const mins = Math.ceil((user.lockedUntil.getTime() - Date.now()) / 60000)
      return NextResponse.json(
        { success: false, error: `Too many wrong attempts. Try again in ${mins} minute${mins === 1 ? '' : 's'}.` },
        { status: 429 }
      )
    }

    if (!(await verifyPassword(password, user.passwordHash))) {
      const attempts = user.failedAttempts + 1
      await db
        .update(users)
        .set(attempts >= MAX_ATTEMPTS ? { failedAttempts: 0, lockedUntil: new Date(Date.now() + LOCK_MINUTES * 60000) } : { failedAttempts: attempts })
        .where(eq(users.id, user.id))
      return NextResponse.json(BAD, { status: 401 })
    }

    await db.update(users).set({ failedAttempts: 0, lockedUntil: null, lastLoginAt: new Date() }).where(eq(users.id, user.id))

    const response = NextResponse.json({ success: true })
    response.cookies.set(SESSION_COOKIE, createSessionToken(user.id), {
      httpOnly: true,
      secure: process.env.NODE_ENV === 'production',
      sameSite: 'lax',
      maxAge: SESSION_MAX_AGE,
      path: '/',
    })
    return response
  } catch (err) {
    console.error(err)
    return NextResponse.json({ success: false, error: 'Something went wrong' }, { status: 500 })
  }
}
