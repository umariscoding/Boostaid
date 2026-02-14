import { NextRequest, NextResponse } from 'next/server'

const VALID_EMAIL = 'admin@boostaid.com'
const VALID_PASSWORD = 'Faiz@123'

export async function POST(request: NextRequest) {
  try {
    const { email, password } = await request.json()

    if (email === VALID_EMAIL && password === VALID_PASSWORD) {
      const response = NextResponse.json({ success: true })
      
      // Set HTTP-only cookie for session (expires in 7 days)
      response.cookies.set('boostaid_session', 'authenticated', {
        httpOnly: true,
        secure: process.env.NODE_ENV === 'production',
        sameSite: 'lax',
        maxAge: 60 * 60 * 24 * 7, // 7 days
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
