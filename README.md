# BoostAid

Cleans donation spreadsheets for HMRC Gift Aid, then tracks every claim from **pending → submitted → received**,
lets you look up any donor, and exports an audit file.

## Features
- **Process** - upload a spreadsheet, map columns, get cleaned data + HMRC submission files. Garbled characters
  (`â€™`, `ðŸ˜‡`) are repaired and emoji/symbols stripped on upload.
- **Claims** - each HMRC file becomes a claim with its own reference (`GA-2026-0001`). Mark it submitted, then
  received; received claims move to the Archive tab. Full status history is kept.
- **Donor search** - search by name, postcode, date range or amount to see if/when a donation was claimed.
- **Audit export** - download all claims (or a single claim) as an Excel file with Claims, Donations and Status History sheets.

## Setup
```bash
npm install
cp .env.example .env.local        # then edit the values
npm run db:migrate                # creates the tables
npm run dev
```

Local Postgres via Docker:
```bash
docker run -d --name boostaid-pg -e POSTGRES_PASSWORD=dev -e POSTGRES_DB=boostaid -p 54329:5432 postgres:16-alpine
```

## Environment variables
| Name | Purpose |
| --- | --- |
| `DATABASE_URL` | Postgres connection string (Neon, Supabase, ...) |
| `ADMIN_EMAIL`, `ADMIN_PASSWORD` | The single dashboard login |
| `SESSION_SECRET` | 16+ random chars used to sign the session cookie (`openssl rand -hex 32`) |

## Deploying (Vercel + Neon)
1. Create a Neon Postgres database and copy its connection string.
2. Set the three variables above in the Vercel project settings.
3. Run the migration once against the production database: `DATABASE_URL=... npm run db:migrate`.
4. Deploy.

Changing the schema: edit `lib/db/schema.ts`, run `npm run db:generate`, commit the new file in `drizzle/`, then `npm run db:migrate`.
