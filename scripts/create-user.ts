/**
 * Create a login, or reset an existing one's password.
 *
 *   npx tsx scripts/create-user.ts <email> [--name "Full Name"] [--password "..."]
 *
 * With no --password a strong random one is generated and printed ONCE. Pass it to the person, who should
 * change it under Settings -> Account. Uses DATABASE_URL (from the environment or .env.local).
 */
import { sql } from 'drizzle-orm'
import { getDb, users } from '../lib/db'
import { generatePassword, hashPassword, passwordProblem } from '../lib/password'

try {
  process.loadEnvFile('.env.local')
} catch {
  // not present: rely on the real environment
}

const args = process.argv.slice(2)
const flag = (f: string) => (args.includes(f) ? args[args.indexOf(f) + 1] : undefined)
const email = args.find((a) => !a.startsWith('--') && args[args.indexOf(a) - 1] !== '--name' && args[args.indexOf(a) - 1] !== '--password')?.trim().toLowerCase()

async function main() {
  if (!email || !email.includes('@')) {
    console.error('usage: npx tsx scripts/create-user.ts <email> [--name "Full Name"] [--password "..."]')
    process.exit(2)
  }
  const name = flag('--name') ?? ''
  const given = flag('--password')
  const password = given ?? generatePassword()
  const problem = passwordProblem(password, email)
  if (problem) {
    console.error(problem)
    process.exit(2)
  }

  const db = getDb()
  const passwordHash = await hashPassword(password)
  const [existing] = await db.select({ id: users.id }).from(users).where(sql`lower(${users.email}) = ${email}`)

  if (existing) {
    await db.update(users).set({ passwordHash, failedAttempts: 0, lockedUntil: null, ...(name ? { name } : {}) }).where(sql`${users.id} = ${existing.id}`)
    console.log(`Updated ${email}: password reset.`)
  } else {
    await db.insert(users).values({ email, name, passwordHash })
    console.log(`Created ${email}.`)
  }
  if (!given) console.log(`Password (shown once): ${password}`)
  process.exit(0)
}

main().catch((e) => {
  console.error(e)
  process.exit(1)
})
