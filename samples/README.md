# Sample data (fake: safe to share and commit)

Made-up donation exports for every platform BoostAid reads. **Every donor, address, ID and amount is invented.**
The column headings and value formats are copied from the real exports, so each file is detected and cleaned
exactly like the real thing. Use these for demos, for testing the app on Vercel, and for checking changes,
instead of the client's real test files (which contain real donor data and must never be committed).

| File | Platform | Rows | Ready | Overseas | Needs review | Not eligible |
|---|---|---|---|---|---|---|
| `LaunchGood_SAMPLE.xlsx` | LaunchGood (2 sheets, like the real export) | 20 + 4 repeats | 7 | 1 | 6 | 5 |
| `Raisely_SAMPLE.xlsx` | Raisely | 17 | 7 | 0 | 6 | 3 |
| `CAF_SAMPLE.xlsx` | CAF Donate | 17 | 6 | 1 | 5 | 4 |
| `MuslimGiving_SAMPLE.xlsx` | MuslimGiving | 15 | 6 | 0 | 5 | 3 |
| `FundraiseUp_SAMPLE.csv` | Fundraise Up | 19 | 6 | 1 | 6 | 5 |
| `DonationsManager_SAMPLE.csv` | Donations Manager | 17 | 6 | 1 | 6 | 3 |
| `CashBook_SAMPLE.csv` | Cash book (HMRC layout) | 16 | 8 | 0 | 6 | 1 |

Each file has one exact duplicate line, which is skipped, so "Rows" is one more than the four result columns add up to.

`EXPECTED_RESULTS.xlsx` is the answer sheet. It lists every line of every file: what it tests, where it should end up
(ready / overseas / review / not eligible / skipped), why, and, for ready lines, exactly what the HMRC schedule
should show (title, names, house, postcode, date, amount).

## What the files test

All platforms:
- clean donation
- messy capitals and lower-case postcode get tidied
- title typed into the first-name box
- flat plus building name
- the same donor giving twice
- no Gift Aid declaration
- £0 donation
- two people in one name
- half a postcode
- missing postcode
- "Anonymous"
- missing surname
- the same line twice in the export

Where the platform has the column, also:
- a payment in US dollars (never converted automatically)
- an overseas donor (overseas schedule, postcode X)

Platform-specific:
- **LaunchGood:**
  - a donation that is refunded later in the same export, plus the refund line itself
  - a payout line
  - a surname filled in from the billing name
  - US dates
  - a second sheet repeating rows
- **Raisely:**
  - a TEST-mode payment
  - a donor who covered the fees, where Donation Total is claimed, not Total
- **CAF:**
  - a Non Tax Effective donation
  - a cancelled payment
- **MuslimGiving:** Gift Aid ticked, but the donor is not a UK taxpayer
- **Fundraise Up:**
  - a failed payment
  - a refunded payment
  - a donation on behalf of a company
  - mixed US date-time styles
- **Donations Manager:** an organisation donor
- **Cash book:**
  - a blank date
  - the same donor on the same day with two different amounts (both are claimed)

## Commands

```bash
npm run samples:check   # cleans every sample and checks each line against the answer sheet
npm run samples:make    # rebuilds the samples (only needed after changing scripts/make-samples.ts)
```
