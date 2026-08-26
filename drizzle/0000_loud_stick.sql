CREATE TABLE `work_records` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`app_id` text NOT NULL,
	`title` text NOT NULL,
	`activity_date` text NOT NULL,
	`activity_type` text NOT NULL,
	`description` text DEFAULT '' NOT NULL,
	`detailed_notes` text DEFAULT '' NOT NULL,
	`duration_minutes` integer NOT NULL,
	`status` text DEFAULT 'complete' NOT NULL,
	`project_ids_json` text DEFAULT '[]' NOT NULL,
	`organization_ids_json` text DEFAULT '[]' NOT NULL,
	`category_ids_json` text DEFAULT '[]' NOT NULL,
	`educator_reach` integer DEFAULT 0 NOT NULL,
	`student_reach` integer DEFAULT 0 NOT NULL,
	`partner_reach` integer DEFAULT 0 NOT NULL,
	`other_reach` integer DEFAULT 0 NOT NULL,
	`output` text DEFAULT '' NOT NULL,
	`outcome` text DEFAULT '' NOT NULL,
	`next_step` text DEFAULT '' NOT NULL,
	`follow_up_needed` integer DEFAULT false NOT NULL,
	`follow_up_date` text,
	`orbit_reportable` integer DEFAULT false NOT NULL,
	`orbit_primary_deliverable` text,
	`orbit_supporting_json` text DEFAULT '[]' NOT NULL,
	`stem_poc_minutes` integer DEFAULT 0 NOT NULL,
	`tac_minutes` integer DEFAULT 0 NOT NULL,
	`orbit_evidence` text DEFAULT '' NOT NULL,
	`is_sample` integer DEFAULT false NOT NULL,
	`created_at` text NOT NULL,
	`modified_at` text NOT NULL
);
--> statement-breakpoint
CREATE UNIQUE INDEX `work_records_app_id_unique` ON `work_records` (`app_id`);--> statement-breakpoint
CREATE INDEX `idx_work_records_activity_date` ON `work_records` (`activity_date`);--> statement-breakpoint
CREATE INDEX `idx_work_records_follow_up_date` ON `work_records` (`follow_up_date`);--> statement-breakpoint
CREATE INDEX `idx_work_records_orbit_date` ON `work_records` (`orbit_reportable`,`activity_date`);