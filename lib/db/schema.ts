import { sql } from 'drizzle-orm'
import {
  date,
  index,
  integer,
  numeric,
  pgEnum,
  pgSequence,
  pgTable,
  text,
  timestamp,
  uniqueIndex,
  uuid,
} from 'drizzle-orm/pg-core'

export const claimStatus = pgEnum('claim_status', ['pending', 'submitted', 'received'])
export type ClaimStatus = (typeof claimStatus.enumValues)[number]

/** Monotonic counter used to build human-readable claim references (GA-2026-0001). */
export const claimRefSeq = pgSequence('claim_ref_seq', { startWith: 1 })

/** People who can log in. Passwords are stored only as scrypt hashes (see lib/password.ts). */
export const users = pgTable(
  'users',
  {
    id: uuid('id').primaryKey().defaultRandom(),
    email: text('email').notNull(),
    name: text('name').notNull().default(''),
    passwordHash: text('password_hash').notNull(),
    /** Wrong passwords in a row; after 5 the account is locked for 15 minutes. */
    failedAttempts: integer('failed_attempts').notNull().default(0),
    lockedUntil: timestamp('locked_until', { withTimezone: true }),
    lastLoginAt: timestamp('last_login_at', { withTimezone: true }),
    createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
  },
  (t) => [uniqueIndex('users_email_lower_unique').on(sql`lower(${t.email})`)]
)

/**
 * A charity whose donations are being claimed. Deleting a charity deletes everything that belongs to it
 * (its uploads, claims, donations and status history) through the foreign keys below.
 */
export const charities = pgTable(
  'charities',
  {
    id: uuid('id').primaryKey().defaultRandom(),
    name: text('name').notNull(),
    createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
  },
  (t) => [uniqueIndex('charities_name_lower_unique').on(sql`lower(${t.name})`)]
)

/** One upload that was processed through the cleaner, for one charity. */
export const batches = pgTable(
  'batches',
  {
    id: uuid('id').primaryKey().defaultRandom(),
    charityId: uuid('charity_id')
      .notNull()
      .references(() => charities.id, { onDelete: 'cascade' }),
    filename: text('filename').notNull(),
    totalRecords: integer('total_records').notNull().default(0),
    eligibleRecords: integer('eligible_records').notNull().default(0),
    validRecords: integer('valid_records').notNull().default(0),
    invalidRecords: integer('invalid_records').notNull().default(0),
    createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
  },
  (t) => [index('batches_charity_idx').on(t.charityId)]
)

/** One HMRC submission file (up to 1000 donations). This is what gets tracked. */
export const claims = pgTable(
  'claims',
  {
    id: uuid('id').primaryKey().defaultRandom(),
    ref: text('ref').notNull().unique(),
    batchId: uuid('batch_id')
      .notNull()
      .references(() => batches.id, { onDelete: 'cascade' }),
    chunkIndex: integer('chunk_index').notNull(),
    status: claimStatus('status').notNull().default('pending'),
    /** 'UK' or 'Overseas': HMRC wants overseas donors in their own schedule. */
    kind: text('kind').notNull().default('UK'),
    /** The schedule's name from the cleaning engine, e.g. UK_schedule_1_of_4.ods */
    fileName: text('file_name'),
    donationCount: integer('donation_count').notNull(),
    totalAmount: numeric('total_amount', { precision: 12, scale: 2 }).notNull(),
    /** Optional reference issued by HMRC once the claim is submitted. */
    hmrcRef: text('hmrc_ref'),
    notes: text('notes'),
    createdAt: timestamp('created_at', { withTimezone: true }).notNull().defaultNow(),
    submittedAt: timestamp('submitted_at', { withTimezone: true }),
    receivedAt: timestamp('received_at', { withTimezone: true }),
  },
  (t) => [index('claims_status_idx').on(t.status), index('claims_batch_idx').on(t.batchId)]
)

/** Append-only status history, for audit. */
export const claimEvents = pgTable(
  'claim_events',
  {
    id: uuid('id').primaryKey().defaultRandom(),
    claimId: uuid('claim_id')
      .notNull()
      .references(() => claims.id, { onDelete: 'cascade' }),
    status: claimStatus('status').notNull(),
    at: timestamp('at', { withTimezone: true }).notNull().defaultNow(),
    note: text('note'),
  },
  (t) => [index('claim_events_claim_idx').on(t.claimId)]
)

/** One donation row included in a claim. Powers the donor lookup. */
export const donations = pgTable(
  'donations',
  {
    id: uuid('id').primaryKey().defaultRandom(),
    claimId: uuid('claim_id')
      .notNull()
      .references(() => claims.id, { onDelete: 'cascade' }),
    title: text('title').notNull().default(''),
    firstName: text('first_name').notNull().default(''),
    lastName: text('last_name').notNull().default(''),
    address: text('address').notNull().default(''),
    postcode: text('postcode').notNull().default(''),
    /** Upper-cased, whitespace-free postcode for matching. */
    postcodeNorm: text('postcode_norm').notNull().default(''),
    donationDate: date('donation_date'),
    amount: numeric('amount', { precision: 12, scale: 2 }).notNull(),
    /** Which platform the donation came from, and its ID there. Used to stop the same donation being claimed twice. */
    platform: text('platform').notNull().default(''),
    platformDonationId: text('platform_donation_id'),
    /** Row in the HMRC schedule (data starts at row 25). */
    hmrcRow: integer('hmrc_row'),
  },
  (t) => [
    index('donations_claim_idx').on(t.claimId),
    index('donations_platform_id_idx').on(t.platformDonationId),
    index('donations_postcode_idx').on(t.postcodeNorm),
    index('donations_last_name_idx').on(t.lastName),
  ]
)
