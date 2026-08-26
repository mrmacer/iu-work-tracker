import type { Category, Contact, Deliverable, Organization, Project, ReferenceData, ReportingConfig, SystemSettings } from "./models";

export const PROJECTS: Project[] = [
  { appId: "project-steels", name: "STEELS Implementation", description: "District planning, professional learning, and implementation support.", status: "active", color: "blue" },
  { appId: "project-ai", name: "AI in Education", description: "Responsible AI learning and instructional support.", status: "active", color: "coral" },
  { appId: "project-keystone", name: "Keystone STEM Competition", description: "Regional student competition planning and delivery.", status: "planning", color: "lime" },
  { appId: "project-ecosystem", name: "STEM Ecosystem", description: "Cross-sector partnership and regional ecosystem development.", status: "active", color: "purple" },
  { appId: "project-makerspace", name: "Makerspace", description: "Hands-on student and educator experiences.", status: "active", color: "yellow" },
];

export const ORGANIZATIONS: Organization[] = [
  { appId: "org-north-valley", name: "North Valley SD", type: "district" },
  { appId: "org-riverbend", name: "Riverbend Area SD", type: "district" },
  { appId: "org-iu", name: "Intermediate Unit", type: "iu" },
  { appId: "org-futureworks", name: "FutureWorks Partnership", type: "partner" },
];

export const CONTACTS: Contact[] = [
  { appId: "contact-north-valley-lead", displayName: "Development District Lead", role: "District curriculum lead (sample)", organizationId: "org-north-valley" },
  { appId: "contact-futureworks", displayName: "Development Partner Contact", role: "Workforce partner (sample)", organizationId: "org-futureworks" },
  { appId: "contact-iu-colleague", displayName: "Development IU Colleague", role: "IU team member (sample)", organizationId: "org-iu" },
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
  contacts: CONTACTS,
  categories: CATEGORIES,
  deliverables: DELIVERABLES,
  reportingConfig: REPORTING_CONFIG,
  settings: SYSTEM_SETTINGS,
};
