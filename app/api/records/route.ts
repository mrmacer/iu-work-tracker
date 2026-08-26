import { env } from "cloudflare:workers";
import { REFERENCE_DATA } from "../../../lib/reference-data";
import { SAMPLE_RECORDS } from "../../../lib/sample-data";
import type { EngagementScope, WorkRecord } from "../../../lib/models";
import { validateWorkRecord } from "../../../lib/validation";

export const dynamic = "force-dynamic";

function db() {
  if (!env.DB) throw new Error("Prototype data store unavailable");
  return env.DB;
}

function jsonArray(value: unknown): string[] {
  try {
    const parsed: unknown = JSON.parse(String(value ?? "[]"));
    return Array.isArray(parsed) && parsed.every((item) => typeof item === "string") ? parsed : [];
  } catch {
    return [];
  }
}

function fromRow(row: Record<string, unknown>): WorkRecord {
  return {
    appId: String(row.app_id),
    title: String(row.title),
    activityDate: String(row.activity_date),
    activityType: String(row.activity_type),
    description: String(row.description),
    detailedNotes: String(row.detailed_notes),
    durationMinutes: Number(row.duration_minutes),
    status: row.status === "draft" ? "draft" : "complete",
    engagementScope: String(row.engagement_scope) as EngagementScope,
    projectIds: jsonArray(row.project_ids_json),
    organizationIds: jsonArray(row.organization_ids_json),
    contactIds: jsonArray(row.contact_ids_json),
    categoryIds: jsonArray(row.category_ids_json),
    reach: {
      educatorsLeaders: Number(row.educator_reach),
      studentsFamilies: Number(row.student_reach),
      workforceCommunity: Number(row.partner_reach),
      other: Number(row.other_reach),
    },
    evidenceSummary: String(row.evidence_summary),
    evidenceReferenceIds: jsonArray(row.evidence_reference_ids_json),
    output: String(row.output),
    outcome: String(row.outcome),
    nextStep: String(row.next_step),
    followUpNeeded: Boolean(row.follow_up_needed),
    followUpDate: row.follow_up_date ? String(row.follow_up_date) : null,
    orbit: {
      reportable: Boolean(row.orbit_reportable),
      primaryDeliverable: row.orbit_primary_deliverable ? String(row.orbit_primary_deliverable) : null,
      supportingDeliverables: jsonArray(row.orbit_supporting_json),
      stemPocMinutes: Number(row.stem_poc_minutes),
      tacMinutes: Number(row.tac_minutes),
      evidence: String(row.orbit_evidence),
    },
    schemaVersion: 2,
    metadata: {
      providerId: String(row.id),
      version: Number(row.record_version),
      createdAt: String(row.created_at),
      modifiedAt: String(row.modified_at),
      syncState: "saved",
    },
    isSample: Boolean(row.is_sample),
  };
}

const insertSql = `INSERT OR IGNORE INTO work_records (
  app_id,title,activity_date,activity_type,description,detailed_notes,duration_minutes,status,engagement_scope,
  project_ids_json,organization_ids_json,contact_ids_json,category_ids_json,educator_reach,student_reach,partner_reach,other_reach,
  evidence_summary,evidence_reference_ids_json,output,outcome,next_step,follow_up_needed,follow_up_date,orbit_reportable,
  orbit_primary_deliverable,orbit_supporting_json,stem_poc_minutes,tac_minutes,orbit_evidence,schema_version,record_version,
  is_sample,created_at,modified_at
) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`;

function insertStatement(database: D1Database, record: WorkRecord) {
  return database.prepare(insertSql).bind(
    record.appId, record.title, record.activityDate, record.activityType, record.description, record.detailedNotes,
    record.durationMinutes, record.status, record.engagementScope, JSON.stringify(record.projectIds), JSON.stringify(record.organizationIds),
    JSON.stringify(record.contactIds), JSON.stringify(record.categoryIds), record.reach.educatorsLeaders, record.reach.studentsFamilies,
    record.reach.workforceCommunity, record.reach.other, record.evidenceSummary, JSON.stringify(record.evidenceReferenceIds), record.output,
    record.outcome, record.nextStep, record.followUpNeeded ? 1 : 0, record.followUpDate, record.orbit.reportable ? 1 : 0,
    record.orbit.primaryDeliverable, JSON.stringify(record.orbit.supportingDeliverables), record.orbit.stemPocMinutes, record.orbit.tacMinutes,
    record.orbit.evidence, record.schemaVersion, record.metadata.version, record.isSample ? 1 : 0, record.metadata.createdAt, record.metadata.modifiedAt,
  );
}

async function seedDevelopmentRecords() {
  const database = db();
  await database.batch(SAMPLE_RECORDS.map((record) => insertStatement(database, record)));
}

function validationResponse(record: unknown) {
  const validation = validateWorkRecord(record, REFERENCE_DATA);
  return validation.valid ? null : Response.json({ status: "validation_error", errors: validation.issues }, { status: 400 });
}

async function currentByAppId(appId: string) {
  return db().prepare("SELECT * FROM work_records WHERE app_id = ?").bind(appId).first<Record<string, unknown>>();
}

export async function GET() {
  try {
    await seedDevelopmentRecords();
    const result = await db().prepare("SELECT * FROM work_records ORDER BY activity_date DESC,created_at DESC").all();
    return Response.json({ status: "success", value: (result.results as Record<string, unknown>[]).map((row) => fromRow(row)) });
  } catch (error) {
    console.error("Work record load failed", error);
    return Response.json({ status: "persistence_error", message: "Records could not be loaded." }, { status: 500 });
  }
}

export async function POST(request: Request) {
  try {
    const input: unknown = await request.json();
    const invalid = validationResponse(input);
    if (invalid) return invalid;
    const record = input as WorkRecord;
    const existing = await currentByAppId(record.appId);
    if (existing) return Response.json({ status: "conflict", current: fromRow(existing), message: "A record with this application ID already exists." }, { status: 409 });
    const now = new Date().toISOString();
    const savedInput: WorkRecord = { ...record, metadata: { version: 1, createdAt: now, modifiedAt: now, syncState: "saved" } };
    const inserted = await insertStatement(db(), savedInput).run();
    if (!inserted.meta.changes) {
      const current = await currentByAppId(record.appId);
      if (current) return Response.json({ status: "conflict", current: fromRow(current), message: "A record with this application ID already exists." }, { status: 409 });
      throw new Error("Record insert was ignored without an existing row.");
    }
    const saved = await currentByAppId(record.appId);
    if (!saved) throw new Error("Created record could not be read back.");
    return Response.json({ status: "success", value: fromRow(saved) }, { status: 201 });
  } catch (error) {
    console.error("Work record create failed", error);
    return Response.json({ status: "persistence_error", message: "The record could not be created. Your draft remains open." }, { status: 500 });
  }
}

export async function PUT(request: Request) {
  try {
    const payload = await request.json() as { record?: unknown; expectedVersion?: unknown };
    const invalid = validationResponse(payload.record);
    if (invalid) return invalid;
    if (!Number.isInteger(payload.expectedVersion) || Number(payload.expectedVersion) < 1) {
      return Response.json({ status: "validation_error", errors: [{ path: "expectedVersion", code: "invalid_version", message: "expectedVersion must be a positive integer." }] }, { status: 400 });
    }
    const record = payload.record as WorkRecord;
    const expectedVersion = Number(payload.expectedVersion);
    const existingRow = await currentByAppId(record.appId);
    if (!existingRow) return Response.json({ status: "persistence_error", message: "The record no longer exists." }, { status: 404 });
    const existing = fromRow(existingRow);
    if (existing.metadata.version !== expectedVersion) return Response.json({ status: "conflict", current: existing, message: "This record changed after you opened it. Your draft remains open." }, { status: 409 });
    const modifiedAt = new Date().toISOString();
    const result = await db().prepare(`UPDATE work_records SET
      title=?,activity_date=?,activity_type=?,description=?,detailed_notes=?,duration_minutes=?,status=?,engagement_scope=?,
      project_ids_json=?,organization_ids_json=?,contact_ids_json=?,category_ids_json=?,educator_reach=?,student_reach=?,partner_reach=?,other_reach=?,
      evidence_summary=?,evidence_reference_ids_json=?,output=?,outcome=?,next_step=?,follow_up_needed=?,follow_up_date=?,orbit_reportable=?,
      orbit_primary_deliverable=?,orbit_supporting_json=?,stem_poc_minutes=?,tac_minutes=?,orbit_evidence=?,schema_version=?,is_sample=?,
      record_version=record_version+1,modified_at=? WHERE app_id=? AND record_version=?`).bind(
      record.title, record.activityDate, record.activityType, record.description, record.detailedNotes, record.durationMinutes, record.status,
      record.engagementScope, JSON.stringify(record.projectIds), JSON.stringify(record.organizationIds), JSON.stringify(record.contactIds),
      JSON.stringify(record.categoryIds), record.reach.educatorsLeaders, record.reach.studentsFamilies, record.reach.workforceCommunity,
      record.reach.other, record.evidenceSummary, JSON.stringify(record.evidenceReferenceIds), record.output, record.outcome, record.nextStep,
      record.followUpNeeded ? 1 : 0, record.followUpDate, record.orbit.reportable ? 1 : 0, record.orbit.primaryDeliverable,
      JSON.stringify(record.orbit.supportingDeliverables), record.orbit.stemPocMinutes, record.orbit.tacMinutes, record.orbit.evidence,
      record.schemaVersion, record.isSample ? 1 : 0, modifiedAt, record.appId, expectedVersion,
    ).run();
    if (!result.meta.changes) {
      const currentRow = await currentByAppId(record.appId);
      if (currentRow) return Response.json({ status: "conflict", current: fromRow(currentRow), message: "This record changed while it was being saved. Your draft remains open." }, { status: 409 });
      throw new Error("Record disappeared during update.");
    }
    const saved = await currentByAppId(record.appId);
    if (!saved) throw new Error("Updated record could not be read back.");
    return Response.json({ status: "success", value: fromRow(saved) });
  } catch (error) {
    console.error("Work record update failed", error);
    return Response.json({ status: "persistence_error", message: "The record could not be updated. Your draft remains open." }, { status: 500 });
  }
}
