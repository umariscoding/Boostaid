ALTER TABLE "claims" ADD COLUMN "kind" text DEFAULT 'UK' NOT NULL;--> statement-breakpoint
ALTER TABLE "claims" ADD COLUMN "file_name" text;--> statement-breakpoint
ALTER TABLE "donations" ADD COLUMN "platform" text DEFAULT '' NOT NULL;--> statement-breakpoint
ALTER TABLE "donations" ADD COLUMN "platform_donation_id" text;--> statement-breakpoint
ALTER TABLE "donations" ADD COLUMN "hmrc_row" integer;--> statement-breakpoint
CREATE INDEX "donations_platform_id_idx" ON "donations" USING btree ("platform_donation_id");