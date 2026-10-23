import type { Category, Contact, Deliverable, Organization, Project, ReferenceData, ReportingConfig, SystemSettings } from "./models";

export const PROJECTS: Project[] = [
  { appId: "project-steels", name: "STEELS Implementation", description: "District planning, professional learning, and implementation support.", status: "active", color: "blue" },
  { appId: "project-ai", name: "AI in Education", description: "Responsible AI learning and instructional support.", status: "active", color: "coral" },
  { appId: "project-keystone", name: "Keystone STEM Competition", description: "Regional student competition planning and delivery.", status: "planning", color: "lime" },
  { appId: "project-ecosystem", name: "STEM Ecosystem", description: "Cross-sector partnership and regional ecosystem development.", status: "active", color: "purple" },
  { appId: "project-makerspace", name: "Makerspace", description: "Hands-on student and educator experiences.", status: "active", color: "yellow" },
];

// Patch 9B — the real IU29 directory. Seeds live in code (they are NOT rows in IU_Organizations),
// are read-only in the UI, and keep their appIds forever. District remains Organization.type ===
// "district": there is no separate District model, provider, or list.
//
// The twelve districts are exactly the school districts IU29 itself lists on iu29.org (District
// Announcements and Guest Teacher Program pages), spelled as IU29 and each district's own site
// spell them (the district's site confirms "Saint Clair", not "St. Clair"). `org-iu` keeps its
// original appId — existing records that link to it now read "Schuylkill Intermediate Unit 29".
// PAIMS is a job-alike group within the Pennsylvania Association of Intermediate Units
// (instructional media, educational technology, and IT), so it is a partner, not a district.
export const ORGANIZATIONS: Organization[] = [
  { appId: "org-iu", name: "Schuylkill Intermediate Unit 29", type: "iu" },
  { appId: "org-blue-mountain", name: "Blue Mountain School District", type: "district" },
  { appId: "org-mahanoy-area", name: "Mahanoy Area School District", type: "district" },
  { appId: "org-minersville-area", name: "Minersville Area School District", type: "district" },
  { appId: "org-north-schuylkill", name: "North Schuylkill School District", type: "district" },
  { appId: "org-pine-grove-area", name: "Pine Grove Area School District", type: "district" },
  { appId: "org-pottsville-area", name: "Pottsville Area School District", type: "district" },
  { appId: "org-saint-clair-area", name: "Saint Clair Area School District", type: "district" },
  { appId: "org-schuylkill-haven-area", name: "Schuylkill Haven Area School District", type: "district" },
  { appId: "org-shenandoah-valley", name: "Shenandoah Valley School District", type: "district" },
  { appId: "org-tamaqua-area", name: "Tamaqua Area School District", type: "district" },
  { appId: "org-tri-valley", name: "Tri-Valley School District", type: "district" },
  { appId: "org-williams-valley", name: "Williams Valley School District", type: "district" },
  { appId: "org-paims", name: "PAIMS", type: "partner" },
];

/**
 * Patch 9B — NON-SELECTABLE compatibility collection. These are the fictional development
 * Organizations the app shipped with. They are retired from every picker, matcher, importer, and
 * directory, but their appIds (and original names) are preserved so any Work Record, Contact, or
 * Inbox Intelligence record that already references one still resolves, still reads correctly, and
 * can still be edited. They are never silently remapped to a real district, and a NEW record can
 * never select one (see lib/organization-directory.ts and lib/validation.ts).
 */
export const LEGACY_ORGANIZATIONS: Organization[] = [
  { appId: "org-north-valley", name: "North Valley SD", type: "district" },
  { appId: "org-riverbend", name: "Riverbend Area SD", type: "district" },
  { appId: "org-futureworks", name: "FutureWorks Partnership", type: "partner" },
];

export const CONTACTS: Contact[] = [
  { appId: "contact-north-valley-lead", displayName: "Development District Lead", role: "District curriculum lead (sample)", organizationId: "org-north-valley", status: "active" },
  { appId: "contact-futureworks", displayName: "Development Partner Contact", role: "Workforce partner (sample)", organizationId: "org-futureworks", status: "active" },
  { appId: "contact-iu-colleague", displayName: "Development IU Colleague", role: "IU team member (sample)", organizationId: "org-iu", status: "active" },
];

export const CATEGORIES: Category[] = [
  { appId: "cat-district", name: "District Support", group: "work-area" },
  { appId: "cat-pl", name: "Professional Learning", group: "work-area" },
  { appId: "cat-stem", name: "STEM / Science", group: "topic" },
  { appId: "cat-steels", name: "STEELS", group: "topic" },
  { appId: "cat-ai", name: "Artificial Intelligence", group: "topic" },
  { appId: "cat-students", name: "Student Programs", group: "work-area" },
  { appId: "cat-partnerships", name: "Partnerships", group: "work-area" },
  { appId: "cat-internal", name: "Internal IU Work", group: "work-area" },
  { appId: "cat-planning", name: "Meetings / Planning", group: "work-area" },
];

export const DELIVERABLES: Deliverable[] = [
  { code: "A", label: "Statewide STEM & CS systems" },
  { code: "B", label: "PA STEELS implementation" },
  { code: "C", label: "CS, AI & computational thinking" },
  { code: "D", label: "Educational leadership" },
  { code: "E", label: "Workforce & ecosystem development" },
  { code: "F", label: "Student competitions & experiences" },
  { code: "G", label: "Math instruction & data literacy" },
];

export const REPORTING_CONFIG: ReportingConfig = {
  minutesPerReportingDay: 420,
  schoolYearStartMonth: 7,
  quarters: [
    { code: "Q1", startMonth: 7, endMonth: 9 },
    { code: "Q2", startMonth: 10, endMonth: 12 },
    { code: "Q3", startMonth: 1, endMonth: 3 },
    { code: "Q4", startMonth: 4, endMonth: 6 },
  ],
};

export const SYSTEM_SETTINGS: SystemSettings = {
  activityTypes: ["District meeting", "Professional learning", "Classroom support", "Project planning", "Partner meeting", "Student program", "Internal planning", "Resource development", "Follow-up communication", "Other"],
};

export const REFERENCE_DATA: ReferenceData = {
  projects: PROJECTS,
  organizations: ORGANIZATIONS,
  legacyOrganizations: LEGACY_ORGANIZATIONS,
  contacts: CONTACTS,
  categories: CATEGORIES,
  deliverables: DELIVERABLES,
  reportingConfig: REPORTING_CONFIG,
  settings: SYSTEM_SETTINGS,
};
