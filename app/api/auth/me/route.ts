import { NextResponse } from 'next/server'
import { eq } from 'drizzle-orm'
import { getDb, users } from '@/lib/db'
import { withAuth, currentUserId } from '@/lib/api'

/** Who is signed in. Also fails (401) if the account has since been removed. */
export const GET = withAuth(async (req) => {
  const id = currentUserId(req)
  const [user] = id ? await getDb().select({ email: users.email, name: users.name }).from(users).where(eq(users.id, id)) : []
  if (!user) return NextResponse.json({ error: 'Unauthorised' }, { status: 401 })
  return NextResponse.json({ email: user.email, name: user.name })
})
