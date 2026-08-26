ALTER TABLE `work_records` ADD `engagement_scope` text DEFAULT 'none' NOT NULL;--> statement-breakpoint
ALTER TABLE `work_records` ADD `contact_ids_json` text DEFAULT '[]' NOT NULL;--> statement-breakpoint
ALTER TABLE `work_records` ADD `evidence_summary` text DEFAULT '' NOT NULL;--> statement-breakpoint
ALTER TABLE `work_records` ADD `evidence_reference_ids_json` text DEFAULT '[]' NOT NULL;--> statement-breakpoint
ALTER TABLE `work_records` ADD `schema_version` integer DEFAULT 2 NOT NULL;--> statement-breakpoint
ALTER TABLE `work_records` ADD `record_version` integer DEFAULT 1 NOT NULL;--> statement-breakpoint
UPDATE `work_records`
SET `engagement_scope` = 'specific'
WHERE `organization_ids_json` IN ('["org-north-valley"]', '["org-riverbend"]');--> statement-breakpoint
UPDATE `work_records`
SET `engagement_scope` = 'regional', `organization_ids_json` = '[]'
WHERE `organization_ids_json` = '["org-regional"]';--> statement-breakpoint
UPDATE `work_records`
SET `evidence_summary` = `orbit_evidence`
WHERE `evidence_summary` = '' AND `orbit_evidence` <> '';--> statement-breakpoint
UPDATE `work_records`
SET `stem_poc_minutes` = 30
WHERE `app_id` = 'sample-ecosystem' AND `stem_poc_minutes` + `tac_minutes` > `duration_minutes`;
