import { NextRequest, NextResponse } from 'next/server'
import { SESSION_COOKIE, verifySessionToken } from '@/lib/session'

const PUBLIC_PREFIXES = ['/login', '/api/auth/login', '/_next/', '/favicon.ico', '/icon.svg']

export function proxy(request: NextRequest) {
  const { pathname } = request.nextUrl
  const authed = verifySessionToken(request.cookies.get(SESSION_COOKIE)?.value)

  if (PUBLIC_PREFIXES.some((p) => pathname.startsWith(p))) {
    // Already logged in and visiting the login page: go home
    if (pathname.startsWith('/login') && authed) {
      return NextResponse.redirect(new URL('/', request.url))
    }
    return NextResponse.next()
  }

  if (!authed) {
    if (pathname.startsWith('/api/')) {
      return NextResponse.json({ error: 'Unauthorised' }, { status: 401 })
    }
    return NextResponse.redirect(new URL('/login', request.url))
  }

  return NextResponse.next()
}

export const config = {
  matcher: ['/((?!_next/static|_next/image|favicon.ico|icon.svg).*)'],
}
