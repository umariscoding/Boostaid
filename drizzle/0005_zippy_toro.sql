CREATE TABLE "admin_notifications" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"type" text NOT NULL,
	"charity_id" uuid,
	"claim_id" uuid,
	"text" text NOT NULL,
	"created_at" timestamp with time zone DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "invoices" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"charity_id" uuid NOT NULL,
	"claim_id" uuid NOT NULL,
	"number" text NOT NULL,
	"recovered_paid" numeric(12, 2) NOT NULL,
	"fee" numeric(12, 2) NOT NULL,
	"status" text DEFAULT 'sent' NOT NULL,
	"sent_at" timestamp with time zone DEFAULT now() NOT NULL,
	"due_at" timestamp with time zone NOT NULL,
	"paid_at" timestamp with time zone,
	"sent_to" text DEFAULT '' NOT NULL,
	"delivered" boolean DEFAULT false NOT NULL
);
--> statement-breakpoint
CREATE TABLE "recovery_log" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"charity_id" uuid NOT NULL,
	"request_id" uuid,
	"type" text NOT NULL,
	"to_email" text,
	"detail" text DEFAULT '' NOT NULL,
	"delivered" boolean,
	"at" timestamp with time zone DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "recovery_requests" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"charity_id" uuid NOT NULL,
	"code_hash" text NOT NULL,
	"email" text NOT NULL,
	"expires_at" timestamp with time zone NOT NULL,
	"used_at" timestamp with time zone,
	"created_at" timestamp with time zone DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "review_rows" (
	"id" uuid PRIMARY KEY DEFAULT gen_random_uuid() NOT NULL,
	"charity_id" uuid NOT NULL,
	"platform" text DEFAULT '' NOT NULL,
	"donation_id" text NOT NULL,
	"donation_date" date,
	"amount" numeric(12, 2) NOT NULL,
	"title" text DEFAULT '' NOT NULL,
	"first_name" text DEFAULT '' NOT NULL,
	"last_name" text DEFAULT '' NOT NULL,
	"house" text DEFAULT '' NOT NULL,
	"postcode" text DEFAULT '' NOT NULL,
	"reason" text DEFAULT '' NOT NULL,
	"source_file" text DEFAULT '' NOT NULL,
	"email" text,
	"status" text DEFAULT 'none' NOT NULL,
	"request_id" uuid,
	"email_sent_at" timestamp with time zone,
	"reminder_at" timestamp with time zone,
	"replied_at" timestamp with time zone,
	"fixed_by" text,
	"fixed_title" text,
	"fixed_first" text,
	"fixed_last" text,
	"fixed_house" text,
	"fixed_postcode" text,
	"claim_id" uuid,
	"created_at" timestamp with time zone DEFAULT now() NOT NULL
);
--> statement-breakpoint
ALTER TABLE "charities" ADD COLUMN "contact_email" text;--> statement-breakpoint
ALTER TABLE "donations" ADD COLUMN "recovered" boolean DEFAULT false NOT NULL;--> statement-breakpoint
ALTER TABLE "donations" ADD COLUMN "email_sent_at" timestamp with time zone;--> statement-breakpoint
ALTER TABLE "donations" ADD COLUMN "replied_at" timestamp with time zone;--> statement-breakpoint
ALTER TABLE "users" ADD COLUMN "role" text DEFAULT 'user' NOT NULL;--> statement-breakpoint
ALTER TABLE "admin_notifications" ADD CONSTRAINT "admin_notifications_charity_id_charities_id_fk" FOREIGN KEY ("charity_id") REFERENCES "public"."charities"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "admin_notifications" ADD CONSTRAINT "admin_notifications_claim_id_claims_id_fk" FOREIGN KEY ("claim_id") REFERENCES "public"."claims"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "invoices" ADD CONSTRAINT "invoices_charity_id_charities_id_fk" FOREIGN KEY ("charity_id") REFERENCES "public"."charities"("id") ON DELETE restrict ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "invoices" ADD CONSTRAINT "invoices_claim_id_claims_id_fk" FOREIGN KEY ("claim_id") REFERENCES "public"."claims"("id") ON DELETE restrict ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "recovery_log" ADD CONSTRAINT "recovery_log_charity_id_charities_id_fk" FOREIGN KEY ("charity_id") REFERENCES "public"."charities"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "recovery_requests" ADD CONSTRAINT "recovery_requests_charity_id_charities_id_fk" FOREIGN KEY ("charity_id") REFERENCES "public"."charities"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "review_rows" ADD CONSTRAINT "review_rows_charity_id_charities_id_fk" FOREIGN KEY ("charity_id") REFERENCES "public"."charities"("id") ON DELETE cascade ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "review_rows" ADD CONSTRAINT "review_rows_claim_id_claims_id_fk" FOREIGN KEY ("claim_id") REFERENCES "public"."claims"("id") ON DELETE set null ON UPDATE no action;--> statement-breakpoint
CREATE INDEX "admin_notifications_created_idx" ON "admin_notifications" USING btree ("created_at");--> statement-breakpoint
CREATE UNIQUE INDEX "invoices_claim_unique" ON "invoices" USING btree ("claim_id");--> statement-breakpoint
CREATE UNIQUE INDEX "invoices_number_unique" ON "invoices" USING btree ("number");--> statement-breakpoint
CREATE INDEX "recovery_log_charity_idx" ON "recovery_log" USING btree ("charity_id","at");--> statement-breakpoint
CREATE UNIQUE INDEX "recovery_requests_code_unique" ON "recovery_requests" USING btree ("code_hash");--> statement-breakpoint
CREATE UNIQUE INDEX "review_rows_key_unique" ON "review_rows" USING btree ("charity_id","platform","donation_id");--> statement-breakpoint
CREATE INDEX "review_rows_charity_status_idx" ON "review_rows" USING btree ("charity_id","status");--> statement-breakpoint
CREATE INDEX "review_rows_request_idx" ON "review_rows" USING btree ("request_id");