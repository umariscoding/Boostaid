import { drizzle } from 'drizzle-orm/node-postgres'
import { Pool } from 'pg'
import * as schema from './schema'

const globalForDb = globalThis as unknown as { __boostaidPool?: Pool }

function getPool(): Pool {
  if (!globalForDb.__boostaidPool) {
    const connectionString = process.env.DATABASE_URL
    if (!connectionString) {
      throw new Error('DATABASE_URL is not set')
    }
    globalForDb.__boostaidPool = new Pool({ connectionString, max: 5 })
  }
  return globalForDb.__boostaidPool
}

type Db = ReturnType<typeof drizzle<typeof schema>>
let cached: Db | undefined

/** Lazily created so `next build` doesn't need a database. */
export function getDb(): Db {
  if (!cached) cached = drizzle(getPool(), { schema })
  return cached
}

export * from './schema'
