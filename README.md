# BoostAid

Cleans donation spreadsheets for HMRC Gift Aid, then tracks every claim from **pending → submitted → received**,
lets you look up any donor, and exports an audit file.

## Features
- **Process** - upload a spreadsheet (.xlsx, .xls or .csv), map columns, get cleaned data + HMRC submission files.
  - Address is reduced to the HMRC "house name or number" (max 40 chars): `7 Wordsworth Rd` -> `7`,
    `302 hadding lodge, 9 knapp road` -> `302 Hadding Lodge`, `Flat 2, 25 Waylen Street` -> `Flat 2`.
    A bare number is valid; no number or house name (`Shelley Ave`) goes to review.
  - Text is cleaned on upload: garbled characters repaired (`a-euro-tm`, emoji mojibake), emoji removed, accents
    converted to plain letters (`Erdogan` with a breve -> `Erdogan`), curly apostrophes straightened, hyphens and
    apostrophes kept, fields trimmed. Non-Latin scripts (e.g. Arabic) can't be converted and go to review.
  - CSV files are read as UTF-8 (falling back to Windows-1252 if the file isn't valid UTF-8).
  - Output files are named with the schedule number, the total donations and the claim reference, e.g.
    `Gift Aid Schedule 1 - £25,000 - GA-2026-0001.xlsx`.
- **Dashboard** - pending / submitted / received claim counts and totals on the main page.
- **Claims** - each HMRC file becomes a claim with its own reference (`GA-2026-0001`). Mark it submitted, then
  received; received claims move to the Archive tab. Full status history is kept.
- **Donor search** - search by name, postcode, date range or amount. One row per donor with donation count, total
  given and the claims their money is in; expand a row for each donation, its claim status and dates.
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
