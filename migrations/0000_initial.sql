CREATE TABLE IF NOT EXISTS work_records (
  id INTEGER PRIMARY KEY AUTOINCREMENT, app_id TEXT NOT NULL UNIQUE, title TEXT NOT NULL, activity_date TEXT NOT NULL,
  activity_type TEXT NOT NULL, description TEXT NOT NULL DEFAULT '', detailed_notes TEXT NOT NULL DEFAULT '', duration_minutes INTEGER NOT NULL,
  status TEXT NOT NULL DEFAULT 'complete', project_ids_json TEXT NOT NULL DEFAULT '[]', organization_ids_json TEXT NOT NULL DEFAULT '[]', category_ids_json TEXT NOT NULL DEFAULT '[]',
  educator_reach INTEGER NOT NULL DEFAULT 0, student_reach INTEGER NOT NULL DEFAULT 0, partner_reach INTEGER NOT NULL DEFAULT 0, other_reach INTEGER NOT NULL DEFAULT 0,
  output TEXT NOT NULL DEFAULT '', outcome TEXT NOT NULL DEFAULT '', next_step TEXT NOT NULL DEFAULT '', follow_up_needed INTEGER NOT NULL DEFAULT 0, follow_up_date TEXT,
  orbit_reportable INTEGER NOT NULL DEFAULT 0, orbit_primary_deliverable TEXT, orbit_supporting_json TEXT NOT NULL DEFAULT '[]', stem_poc_minutes INTEGER NOT NULL DEFAULT 0,
  tac_minutes INTEGER NOT NULL DEFAULT 0, orbit_evidence TEXT NOT NULL DEFAULT '', is_sample INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL, modified_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_work_records_activity_date ON work_records(activity_date);
CREATE INDEX IF NOT EXISTS idx_work_records_follow_up_date ON work_records(follow_up_date) WHERE follow_up_needed = 1;
CREATE INDEX IF NOT EXISTS idx_work_records_orbit_date ON work_records(orbit_reportable, activity_date) WHERE orbit_reportable = 1;
