import { NextRequest, NextResponse } from 'next/server'
import {
  SESSION_COOKIE,
  SESSION_MAX_AGE,
  createSessionToken,
  safeEqual,
} from '@/lib/session'

export async function POST(request: NextRequest) {
  try {
    const expectedEmail = process.env.ADMIN_EMAIL
    const expectedPassword = process.env.ADMIN_PASSWORD
    if (!expectedEmail || !expectedPassword) {
      console.error('ADMIN_EMAIL / ADMIN_PASSWORD are not configured')
      return NextResponse.json(
        { success: false, error: 'Login is not configured' },
        { status: 500 }
      )
    }

    const { email, password } = await request.json()
    const emailOk = safeEqual(String(email ?? '').trim().toLowerCase(), expectedEmail.toLowerCase())
    const passwordOk = safeEqual(String(password ?? ''), expectedPassword)

    if (emailOk && passwordOk) {
      const response = NextResponse.json({ success: true })
      response.cookies.set(SESSION_COOKIE, createSessionToken(), {
        httpOnly: true,
        secure: process.env.NODE_ENV === 'production',
        sameSite: 'lax',
        maxAge: SESSION_MAX_AGE,
        path: '/',
      })
      return response
    }

    return NextResponse.json(
      { success: false, error: 'Invalid email or password' },
      { status: 401 }
    )
  } catch {
    return NextResponse.json(
      { success: false, error: 'Something went wrong' },
      { status: 500 }
    )
  }
}
