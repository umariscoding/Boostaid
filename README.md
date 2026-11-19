# BoostAid

Cleans donation spreadsheets for HMRC Gift Aid, then tracks every claim from **pending → submitted → received**,
lets you look up any donor, and exports an audit file.

## Features
- **Charities** - add and remove charities under **Settings -> Manage charities**. The sidebar dropdown picks the
  charity you're working in, or **All charities** to see everything together. Every upload, claim and donor belongs
  to one charity; uploading with "All charities" selected asks you to choose which one. Duplicate-donation warnings
  only compare within the same charity. Deleting a charity deletes all of its records (uploads, claims, donors and
  status history) after a warning that shows how much will go. With no charities yet, every page prompts you to add one.
- **Clean** - upload one or more donation exports (.xlsx, .xls, .csv), even from different platforms, for one
  charity. The Gift Aid engine in `lib/giftaid` (the client's reference rules) does the cleaning:
  - recognises LaunchGood, Raisely, CAF Donate, MuslimGiving, Fundraise Up, Donations Manager and a cash book from
    their headings; any other layout can be matched by hand with the column mapper
  - works out which donations are eligible, fixes names/addresses/postcodes/dates, and holds back anything it can't
    fix with certainty, each with a reason ("Needs review"); non-GBP, refunds, organisations etc. are not eligible
  - HMRC schedules are saved as `.ods` from HMRC's own template, max 1,000 rows each, UK and overseas separate,
    named like `UK_schedule_1_of_2 - £60,446 - GA-2026-0009.ods`
  - a claim pack zip: the schedules, a review workbook, a summary (print to PDF) and the original files with fingerprints
  - donations already claimed for the charity (same platform + donation ID) are taken out before saving, and the
    server refuses them too
  - acceptance test: `npm run giftaid:verify -- <folder with the client's test exports>` (needs the exports, which
    are real donor data and must never be committed)
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
