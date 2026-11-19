/**
 * Give someone the admin role (BoostAid's owner login, which can open the Admin page), or take it away.
 *
 *   npx tsx scripts/set-role.ts <email> admin
 *   npx tsx scripts/set-role.ts <email> user
 *
 * Uses DATABASE_URL (from the environment or .env.local).
 */
import { sql } from 'drizzle-orm'
import { getDb, users } from '../lib/db'

try {
  process.loadEnvFile('.env.local')
} catch {
  // not present: rely on the real environment
}

async function main() {
  const [email, role] = process.argv.slice(2)
  if (!email || !['admin', 'user'].includes(role)) {
    console.error('usage: npx tsx scripts/set-role.ts <email> admin|user')
    process.exit(2)
  }
  const done = await getDb().update(users).set({ role }).where(sql`lower(${users.email}) = ${email.trim().toLowerCase()}`).returning({ id: users.id })
  if (!done.length) {
    console.error(`No user with the email ${email}.`)
    process.exit(1)
  }
  console.log(`${email} is now ${role === 'admin' ? 'an admin' : 'a normal user'}.`)
  process.exit(0)
}
main()
