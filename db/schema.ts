import { index, integer, sqliteTable, text } from "drizzle-orm/sqlite-core";

export const workRecords = sqliteTable("work_records", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  appId: text("app_id").notNull().unique(), title: text("title").notNull(), activityDate: text("activity_date").notNull(),
  activityType: text("activity_type").notNull(), description: text("description").notNull().default(""), detailedNotes: text("detailed_notes").notNull().default(""),
  durationMinutes: integer("duration_minutes").notNull(), status: text("status").notNull().default("complete"),
  projectIdsJson: text("project_ids_json").notNull().default("[]"), organizationIdsJson: text("organization_ids_json").notNull().default("[]"), categoryIdsJson: text("category_ids_json").notNull().default("[]"),
  educatorReach: integer("educator_reach").notNull().default(0), studentReach: integer("student_reach").notNull().default(0), partnerReach: integer("partner_reach").notNull().default(0), otherReach: integer("other_reach").notNull().default(0),
  output: text("output").notNull().default(""), outcome: text("outcome").notNull().default(""), nextStep: text("next_step").notNull().default(""),
  followUpNeeded: integer("follow_up_needed", { mode: "boolean" }).notNull().default(false), followUpDate: text("follow_up_date"),
  orbitReportable: integer("orbit_reportable", { mode: "boolean" }).notNull().default(false), orbitPrimaryDeliverable: text("orbit_primary_deliverable"),
  orbitSupportingJson: text("orbit_supporting_json").notNull().default("[]"), stemPocMinutes: integer("stem_poc_minutes").notNull().default(0), tacMinutes: integer("tac_minutes").notNull().default(0), orbitEvidence: text("orbit_evidence").notNull().default(""),
  isSample: integer("is_sample", { mode: "boolean" }).notNull().default(false), createdAt: text("created_at").notNull(), modifiedAt: text("modified_at").notNull(),
}, (table) => [
  index("idx_work_records_activity_date").on(table.activityDate),
  index("idx_work_records_follow_up_date").on(table.followUpDate),
  index("idx_work_records_orbit_date").on(table.orbitReportable, table.activityDate),
]);
