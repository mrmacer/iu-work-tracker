import { WORK_RECORD_SCHEMA_VERSION, type ReferenceData, type WorkRecord } from "./models";

export type ValidationIssue = { path: string; code: string; message: string };
export type ValidationResult = { valid: true; record: WorkRecord } | { valid: false; issues: ValidationIssue[] };

const scopes = new Set(["none", "specific", "regional", "allDistricts"]);
const statuses = new Set(["complete", "draft"]);
const datePattern = /^\d{4}-\d{2}-\d{2}$/;

function issue(issues: ValidationIssue[], path: string, code: string, message: string) {
  issues.push({ path, code, message });
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((item) => typeof item === "string" && item.trim().length > 0);
}

function validateIdArray(
  issues: ValidationIssue[],
  path: string,
  value: unknown,
  allowed: Set<string> | null,
) {
  if (!isStringArray(value)) {
    issue(issues, path, "invalid_array", `${path} must be an array of non-empty IDs.`);
    return [];
  }
  if (new Set(value).size !== value.length) issue(issues, path, "duplicate_id", `${path} cannot contain duplicate IDs.`);
  if (allowed) for (const id of value) if (!allowed.has(id)) issue(issues, path, "invalid_id", `${id} is not a canonical ${path} value.`);
  return value;
}

function nonNegativeInteger(issues: ValidationIssue[], path: string, value: unknown) {
  if (!Number.isInteger(value) || Number(value) < 0) issue(issues, path, "invalid_number", `${path} must be a non-negative whole number.`);
}

export function validateWorkRecord(value: unknown, references: ReferenceData): ValidationResult {
  const issues: ValidationIssue[] = [];
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return { valid: false, issues: [{ path: "record", code: "invalid_record", message: "Record must be an object." }] };
  }
  const record = value as WorkRecord;
  for (const field of ["appId", "title", "activityDate", "activityType"] as const) {
    if (typeof record[field] !== "string" || !record[field].trim()) issue(issues, field, "required", `${field} is required.`);
  }
  if (typeof record.activityDate === "string" && !datePattern.test(record.activityDate)) issue(issues, "activityDate", "invalid_date", "activityDate must use YYYY-MM-DD.");
  if (!references.settings.activityTypes.includes(record.activityType)) issue(issues, "activityType", "invalid_value", "activityType must be selected from the configured vocabulary.");
  for (const field of ["description", "detailedNotes", "evidenceSummary", "output", "outcome", "nextStep"] as const) {
    if (typeof record[field] !== "string") issue(issues, field, "invalid_string", `${field} must be a string.`);
  }
  if (!Number.isInteger(record.durationMinutes) || record.durationMinutes <= 0) issue(issues, "durationMinutes", "invalid_duration", "durationMinutes must be a positive whole number.");
  if (!statuses.has(record.status)) issue(issues, "status", "invalid_value", "status must be complete or draft.");
  if (!scopes.has(record.engagementScope)) issue(issues, "engagementScope", "invalid_scope", "engagementScope is invalid.");

  const projectIds = validateIdArray(issues, "projectIds", record.projectIds, new Set(references.projects.map((item) => item.appId)));
  const organizationIds = validateIdArray(issues, "organizationIds", record.organizationIds, new Set(references.organizations.map((item) => item.appId)));
  validateIdArray(issues, "contactIds", record.contactIds, new Set(references.contacts.map((item) => item.appId)));
  validateIdArray(issues, "categoryIds", record.categoryIds, new Set(references.categories.map((item) => item.appId)));
  validateIdArray(issues, "evidenceReferenceIds", record.evidenceReferenceIds, null);
  void projectIds;

  const districtIds = new Set(references.organizations.filter((item) => item.type === "district").map((item) => item.appId));
  const selectedDistricts = organizationIds.filter((id) => districtIds.has(id));
  if (record.engagementScope === "specific" && selectedDistricts.length === 0) issue(issues, "engagementScope", "scope_requires_district", "Specific scope requires at least one canonical district.");
  if (record.engagementScope !== "specific" && selectedDistricts.length > 0) issue(issues, "organizationIds", "scope_district_mismatch", "District IDs are only valid with specific scope.");

  if (!record.reach || typeof record.reach !== "object" || Array.isArray(record.reach)) {
    issue(issues, "reach", "invalid_object", "reach must be an object.");
  } else {
    nonNegativeInteger(issues, "reach.educatorsLeaders", record.reach.educatorsLeaders);
    nonNegativeInteger(issues, "reach.studentsFamilies", record.reach.studentsFamilies);
    nonNegativeInteger(issues, "reach.workforceCommunity", record.reach.workforceCommunity);
    nonNegativeInteger(issues, "reach.other", record.reach.other);
  }

  if (typeof record.followUpNeeded !== "boolean") issue(issues, "followUpNeeded", "invalid_boolean", "followUpNeeded must be boolean.");
  if (record.followUpDate !== null && (typeof record.followUpDate !== "string" || !datePattern.test(record.followUpDate))) issue(issues, "followUpDate", "invalid_date", "followUpDate must be null or YYYY-MM-DD.");

  if (!record.orbit || typeof record.orbit !== "object" || Array.isArray(record.orbit)) {
    issue(issues, "orbit", "invalid_object", "orbit must be a valid object.");
  } else {
    if (typeof record.orbit.reportable !== "boolean") issue(issues, "orbit.reportable", "invalid_boolean", "orbit.reportable must be boolean.");
    const deliverableIds = new Set(references.deliverables.map((item) => item.code));
    const supporting = validateIdArray(issues, "orbit.supportingDeliverables", record.orbit.supportingDeliverables, deliverableIds);
    if (record.orbit.primaryDeliverable !== null && (typeof record.orbit.primaryDeliverable !== "string" || !deliverableIds.has(record.orbit.primaryDeliverable))) issue(issues, "orbit.primaryDeliverable", "invalid_id", "Primary deliverable is invalid.");
    if (record.orbit.reportable && !record.orbit.primaryDeliverable) issue(issues, "orbit.primaryDeliverable", "required", "A reportable record requires one primary deliverable.");
    if (record.orbit.primaryDeliverable && supporting.includes(record.orbit.primaryDeliverable)) issue(issues, "orbit.supportingDeliverables", "duplicate_primary", "The primary deliverable cannot also be supporting.");
    nonNegativeInteger(issues, "orbit.stemPocMinutes", record.orbit.stemPocMinutes);
    nonNegativeInteger(issues, "orbit.tacMinutes", record.orbit.tacMinutes);
    if (Number.isFinite(record.durationMinutes) && Number(record.orbit.stemPocMinutes) + Number(record.orbit.tacMinutes) > record.durationMinutes) issue(issues, "orbit", "reporting_time_exceeds_duration", "PoC and TaC minutes are allocations and cannot together exceed activity duration.");
    if (typeof record.orbit.evidence !== "string") issue(issues, "orbit.evidence", "invalid_string", "ORBIT evidence must be a string.");
    if (!record.orbit.reportable && (record.orbit.primaryDeliverable !== null || supporting.length > 0 || record.orbit.stemPocMinutes !== 0 || record.orbit.tacMinutes !== 0)) issue(issues, "orbit", "non_reportable_has_reporting_data", "Non-reportable records cannot retain ORBIT classifications or reporting time.");
  }

  if (record.schemaVersion !== WORK_RECORD_SCHEMA_VERSION) issue(issues, "schemaVersion", "unsupported_schema", `schemaVersion must be ${WORK_RECORD_SCHEMA_VERSION}.`);
  if (!record.metadata || typeof record.metadata !== "object" || Array.isArray(record.metadata)) {
    issue(issues, "metadata", "invalid_object", "metadata must be a provider metadata object.");
  } else {
    if (!Number.isInteger(record.metadata.version) || record.metadata.version < 0) issue(issues, "metadata.version", "invalid_version", "metadata.version must be a non-negative integer.");
    if (typeof record.metadata.createdAt !== "string" || typeof record.metadata.modifiedAt !== "string") issue(issues, "metadata", "invalid_timestamp", "Provider timestamps must be strings.");
  }
  if (typeof record.isSample !== "boolean") issue(issues, "isSample", "invalid_boolean", "isSample must be boolean.");

  return issues.length ? { valid: false, issues } : { valid: true, record };
}

export function validationMessage(issues: ValidationIssue[]) {
  return issues[0]?.message ?? "The record is invalid.";
}
