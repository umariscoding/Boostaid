CREATE TABLE "gasds_claims" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"charity_id" uuid NOT NULL,
	"tax_year" integer NOT NULL,
	"amount" numeric(12, 2) NOT NULL,
	"hmrc_ref" text NOT NULL,
	"sent_at" timestamp with time zone DEFAULT now() NOT NULL,
	"sent_by" text DEFAULT '' NOT NULL,
	"gift_aid_figure" numeric(12, 2) DEFAULT '0' NOT NULL,
	"status" "claim_status" DEFAULT 'submitted' NOT NULL,
	"received_amount" numeric(12, 2),
	"received_at" timestamp with time zone
);
--> statement-breakpoint
CREATE TABLE "gasds_photos" (
	"receipt_id" uuid PRIMARY KEY NOT NULL,
	"mime" text NOT NULL,
	"data" "bytea" NOT NULL
);
--> statement-breakpoint
CREATE TABLE "gasds_receipts" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"charity_id" uuid NOT NULL,
	"receipt_no" text NOT NULL,
	"receipt_date" date NOT NULL,
	"tax_year" integer NOT NULL,
	"amount" numeric(12, 2) NOT NULL,
	"card" boolean DEFAULT false NOT NULL,
	"not_cash" boolean DEFAULT false NOT NULL,
	"notes" text DEFAULT '' NOT NULL,
	"donor" text DEFAULT '' NOT NULL,
	"is_gift" boolean DEFAULT true NOT NULL,
	"gift_aid" boolean DEFAULT false NOT NULL,
	"banked" boolean DEFAULT false NOT NULL,
	"slip" text,
	"counted" numeric(12, 2),
	"count" jsonb,
	"photo_name" text,
	"claim_id" uuid,
	"used" numeric(12, 2),
	"created_at" timestamp with time zone DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "gasds_years" (
	"charity_id" uuid NOT NULL,
	"tax_year" integer NOT NULL,
	"gift_aid_elsewhere" numeric(12, 2) DEFAULT '0' NOT NULL
);
--> statement-breakpoint
ALTER TABLE "gasds_claims" ADD CONSTRAINT "gasds_claims_charity_id_charities_id_fk" FOREIGN KEY ("charity_id") REFERENCES "public"."charities"("id") ON DELETE restrict ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "gasds_photos" ADD CONSTRAINT "gasds_photos_receipt_id_gasds_receipts_id_fk" FOREIGN KEY ("receipt_id") REFERENCES "public"."gasds_receipts"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "gasds_receipts" ADD CONSTRAINT "gasds_receipts_charity_id_charities_id_fk" FOREIGN KEY ("charity_id") REFERENCES "public"."charities"("id") ON DELETE restrict ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "gasds_receipts" ADD CONSTRAINT "gasds_receipts_claim_id_gasds_claims_id_fk" FOREIGN KEY ("claim_id") REFERENCES "public"."gasds_claims"("id") ON DELETE restrict ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "gasds_years" ADD CONSTRAINT "gasds_years_charity_id_charities_id_fk" FOREIGN KEY ("charity_id") REFERENCES "public"."charities"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
CREATE INDEX "gasds_claims_charity_idx" ON "gasds_claims" USING btree ("charity_id","tax_year");--> statement-breakpoint
CREATE UNIQUE INDEX "gasds_receipts_charity_no_unique" ON "gasds_receipts" USING btree ("charity_id","receipt_no");--> statement-breakpoint
CREATE INDEX "gasds_receipts_year_idx" ON "gasds_receipts" USING btree ("charity_id","tax_year");--> statement-breakpoint
CREATE INDEX "gasds_receipts_claim_idx" ON "gasds_receipts" USING btree ("claim_id");--> statement-breakpoint
CREATE UNIQUE INDEX "gasds_years_unique" ON "gasds_years" USING btree ("charity_id","tax_year");