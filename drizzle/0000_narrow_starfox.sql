CREATE TYPE "public"."claim_status" AS ENUM('pending', 'submitted', 'received');--> statement-breakpoint
CREATE SEQUENCE "public"."claim_ref_seq" INCREMENT BY 1 MINVALUE 1 MAXVALUE 9223372036854775807 START WITH 1 CACHE 1;--> statement-breakpoint
CREATE TABLE "batches" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"filename" text NOT NULL,
	"total_records" integer DEFAULT 0 NOT NULL,
	"eligible_records" integer DEFAULT 0 NOT NULL,
	"valid_records" integer DEFAULT 0 NOT NULL,
	"invalid_records" integer DEFAULT 0 NOT NULL,
	"created_at" timestamp with time zone DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "claim_events" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"claim_id" uuid NOT NULL,
	"status" "claim_status" NOT NULL,
	"at" timestamp with time zone DEFAULT now() NOT NULL,
	"note" text
);
--> statement-breakpoint
CREATE TABLE "claims" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"ref" text NOT NULL,
	"batch_id" uuid NOT NULL,
	"chunk_index" integer NOT NULL,
	"status" "claim_status" DEFAULT 'pending' NOT NULL,
	"donation_count" integer NOT NULL,
	"total_amount" numeric(12, 2) NOT NULL,
	"hmrc_ref" text,
	"notes" text,
	"created_at" timestamp with time zone DEFAULT now() NOT NULL,
	"submitted_at" timestamp with time zone,
	"received_at" timestamp with time zone,
	CONSTRAINT "claims_ref_unique" UNIQUE("ref")
);
--> statement-breakpoint
CREATE TABLE "donations" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"claim_id" uuid NOT NULL,
	"title" text DEFAULT '' NOT NULL,
	"first_name" text DEFAULT '' NOT NULL,
	"last_name" text DEFAULT '' NOT NULL,
	"address" text DEFAULT '' NOT NULL,
	"postcode" text DEFAULT '' NOT NULL,
	"postcode_norm" text DEFAULT '' NOT NULL,
	"donation_date" date,
	"amount" numeric(12, 2) NOT NULL
);
--> statement-breakpoint
ALTER TABLE "claim_events" ADD CONSTRAINT "claim_events_claim_id_claims_id_fk" FOREIGN KEY ("claim_id") REFERENCES "public"."claims"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "claims" ADD CONSTRAINT "claims_batch_id_batches_id_fk" FOREIGN KEY ("batch_id") REFERENCES "public"."batches"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "donations" ADD CONSTRAINT "donations_claim_id_claims_id_fk" FOREIGN KEY ("claim_id") REFERENCES "public"."claims"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
CREATE INDEX "claim_events_claim_idx" ON "claim_events" USING btree ("claim_id");--> statement-breakpoint
CREATE INDEX "claims_status_idx" ON "claims" USING btree ("status");--> statement-breakpoint
CREATE INDEX "claims_batch_idx" ON "claims" USING btree ("batch_id");--> statement-breakpoint
CREATE INDEX "donations_claim_idx" ON "donations" USING btree ("claim_id");--> statement-breakpoint
CREATE INDEX "donations_postcode_idx" ON "donations" USING btree ("postcode_norm");--> statement-breakpoint
CREATE INDEX "donations_last_name_idx" ON "donations" USING btree ("last_name");