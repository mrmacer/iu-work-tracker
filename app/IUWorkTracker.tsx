"use client";

import { useEffect, useRef, useState } from "react";
import {
  PrototypeFallbackProvider,
  selectDataProvider,
  type ActiveProviderKind,
  type DataProvider,
} from "../lib/data-provider";
import { computeInboxIntelligenceSummary, type InboxIntelligenceRecord } from "../lib/inbox-intelligence-models";
import {
  selectInboxIntelligenceProvider,
  type InboxIntelligenceProvider,
  type InboxIntelligenceResult,
} from "../lib/inbox-intelligence-provider";
import {
  filterNeedsAttention,
  formatDueLabel,
  earliestDueDate,
  primaryLabel,
  selectNeedsAttention,
  selectWaiting,
} from "../lib/inbox-action-center";
import type { MeetingRecord } from "../lib/meeting-intelligence-models";
import {
  selectMeetingRecordProvider,
  type MeetingRecordProvider,
  type MeetingRecordResult,
} from "../lib/meeting-record-provider";
import { WORK_RECORD_SCHEMA_VERSION, type Contact, type Organization, type Project, type ReferenceData, type WorkRecord } from "../lib/models";
import { selectContactProvider, type ContactProvider, type ContactResult } from "../lib/contact-provider";
import { buildContactRelationshipSummary } from "../lib/contact-relationships";
import {
  buildOrganizationDraft,
  normalizeOrganizationName,
  ORGANIZATION_TYPES,
  selectOrganizationProvider,
  validateOrganizationShape,
  type OrganizationProvider,
  type OrganizationResult,
  type OrganizationTypeValue,
} from "../lib/organization-provider";
import {
  buildProjectDraft,
  PROJECT_STATUSES,
  selectProjectProvider,
  validateProjectShape,
  type ProjectProvider,
  type ProjectResult,
  type ProjectStatus,
} from "../lib/project-provider";
import { deriveReportingDays } from "../lib/reporting";
import ContactFormModal, { CONTACT_STATUS_LABELS, emptyContactDraft } from "./ContactFormModal";
import ContactImport from "./ContactImport";
import DevMicrosoftConnection from "./DevMicrosoftConnection";
import InboxIntelligence from "./InboxIntelligence";
import MeetingNotes from "./MeetingNotes";
import VoiceIntelligence from "./VoiceIntelligence";

type View = "home" | "today" | "history" | "projects" | "contacts" | "organizations" | "orbit" | "inbox" | "voice" | "meeting";
const navItems: [View, string, string][] = [
  ["home", "⌂", "Home"],
  ["today", "◷", "Today"],
  ["history", "≡", "History"],
  ["projects", "▤", "Projects"],
  ["contacts", "☺", "Contacts"],
  ["organizations", "🏛", "Organizations"],
  ["orbit", "◎", "STEM / ORBIT"],
  ["inbox", "✉", "Inbox Intelligence"],
  ["voice", "🎙", "Voice Intelligence"],
  ["meeting", "📝", "Meeting Notes"],
];
const todayIso = () =>
  new Intl.DateTimeFormat("en-CA", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(new Date());
const hours = (minutes: number) =>
  minutes < 60
    ? `${minutes}m`
    : `${Math.floor(minutes / 60)}h${minutes % 60 ? ` ${minutes % 60}m` : ""}`;
const niceDate = (date: string) =>
  new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
  }).format(new Date(`${date}T12:00:00`));
const entityName = (items: { appId: string; name: string }[], id: string) =>
  items.find((item) => item.appId === id)?.name ?? "Unassigned";

function emptyRecord(): WorkRecord {
  return {
    appId: crypto.randomUUID(),
    title: "",
    activityDate: todayIso(),
    activityType: "",
    description: "",
    detailedNotes: "",
    durationMinutes: 60,
    status: "complete",
    engagementScope: "none",
    projectIds: [],
    organizationIds: [],
    contactIds: [],
    categoryIds: [],
    reach: {
      educatorsLeaders: 0,
      studentsFamilies: 0,
      workforceCommunity: 0,
      other: 0,
    },
    evidenceSummary: "",
    evidenceReferenceIds: [],
    output: "",
    outcome: "",
    nextStep: "",
    followUpNeeded: false,
    followUpDate: null,
    orbit: {
      reportable: false,
      primaryDeliverable: null,
      supportingDeliverables: [],
      stemPocMinutes: 0,
      tacMinutes: 0,
      evidence: "",
    },
    schemaVersion: WORK_RECORD_SCHEMA_VERSION,
    metadata: { version: 0, createdAt: "", modifiedAt: "", syncState: "saved" },
    isSample: false,
  };
}

export default function IUWorkTracker({
  dataProvider,
  inboxDataProvider,
  meetingDataProvider,
  projectDataProvider,
  contactDataProvider,
  organizationDataProvider,
}: {
  dataProvider?: DataProvider;
  inboxDataProvider?: InboxIntelligenceProvider;
  meetingDataProvider?: MeetingRecordProvider;
  projectDataProvider?: ProjectProvider;
  contactDataProvider?: ContactProvider;
  organizationDataProvider?: OrganizationProvider;
} = {}) {
  const [view, setView] = useState<View>("home");
  const [records, setRecords] = useState<WorkRecord[]>([]);
  const [references, setReferences] = useState<ReferenceData | null>(null);
  const [loading, setLoading] = useState(true);
  const [storageMode, setStorageMode] = useState<ActiveProviderKind>("memory");
  const [logging, setLogging] = useState(false);
  const [draft, setDraft] = useState<WorkRecord>(emptyRecord);
  const [step, setStep] = useState(1);
  const [saveError, setSaveError] = useState("");
  const [saving, setSaving] = useState(false);
  const [search, setSearch] = useState("");
  const [draftBaseline, setDraftBaseline] = useState("");
  const opener = useRef<HTMLElement | null>(null);
  const provider = useRef<DataProvider | null>(dataProvider ?? null);
  const onSavedRef = useRef<((saved: WorkRecord) => void) | null>(null);
  const inboxProvider = useRef<InboxIntelligenceProvider | null>(inboxDataProvider ?? null);
  const [inboxRecords, setInboxRecords] = useState<InboxIntelligenceRecord[]>([]);
  const [inboxLoadFailed, setInboxLoadFailed] = useState(false);
  useEffect(() => {
    // Independent of the Work Record load above: a failure here never affects Work Records,
    // and vice versa. This is a read-only list() call — zero Anthropic calls, matching
    // docs/INBOX_INTELLIGENCE_SHAREPOINT_REPORT.md "AI cost control".
    let live = true;
    (async () => {
      if (!inboxProvider.current) {
        const selected = await selectInboxIntelligenceProvider();
        inboxProvider.current = selected.provider;
      }
      const result = await inboxProvider.current.list();
      if (!live) return;
      if (result.status === "success") setInboxRecords(result.value);
      else setInboxLoadFailed(true);
    })();
    return () => {
      live = false;
    };
  }, []);
  const meetingProvider = useRef<MeetingRecordProvider | null>(meetingDataProvider ?? null);
  const [meetingRecords, setMeetingRecords] = useState<MeetingRecord[]>([]);
  const [meetingLoadFailed, setMeetingLoadFailed] = useState(false);
  useEffect(() => {
    // Independent of the Work Record and Inbox Intelligence loads: a failure here never
    // affects them, and vice versa. Read-only list() call — zero Anthropic calls, matching
    // docs/AI_HANDOFF.md "Meeting Notes V1" cost discipline.
    let live = true;
    (async () => {
      if (!meetingProvider.current) {
        const selected = await selectMeetingRecordProvider();
        meetingProvider.current = selected.provider;
      }
      const result = await meetingProvider.current.list();
      if (!live) return;
      if (result.status === "success") setMeetingRecords(result.value);
      else setMeetingLoadFailed(true);
    })();
    return () => {
      live = false;
    };
  }, []);
  const saveMeetingRecord = async (record: MeetingRecord): Promise<MeetingRecordResult<MeetingRecord>> => {
    const active = meetingProvider.current;
    if (!active) return { status: "network_error", message: "The meeting store is still starting up. Try again in a moment." };
    const result = await active.create(record);
    if (result.status === "success") {
      setMeetingRecords((current) => [result.value, ...current.filter((item) => item.appId !== result.value.appId)]);
    }
    return result;
  };
  const updateMeetingRecord = async (
    record: MeetingRecord,
    expectedVersion: number,
  ): Promise<MeetingRecordResult<MeetingRecord>> => {
    const active = meetingProvider.current;
    if (!active) return { status: "network_error", message: "The meeting store is still starting up. Try again in a moment." };
    const result = await active.update(record, expectedVersion);
    if (result.status === "success") {
      setMeetingRecords((current) => current.map((item) => (item.appId === result.value.appId ? result.value : item)));
    }
    return result;
  };
  const projectProvider = useRef<ProjectProvider | null>(projectDataProvider ?? null);
  const [projectRecords, setProjectRecords] = useState<Project[]>([]);
  const [projectLoadFailed, setProjectLoadFailed] = useState(false);
  useEffect(() => {
    // Independent of every other load: a failure here never affects Work Records, Inbox
    // Intelligence, or Meeting Records, and vice versa. Read-only list() call — zero Anthropic
    // calls, matching docs/AI_HANDOFF.md "Durable Projects (Patch 7)" cost discipline.
    let live = true;
    (async () => {
      if (!projectProvider.current) {
        const selected = await selectProjectProvider();
        projectProvider.current = selected.provider;
      }
      const result = await projectProvider.current.list();
      if (!live) return;
      if (result.status === "success") setProjectRecords(result.value);
      else setProjectLoadFailed(true);
    })();
    return () => {
      live = false;
    };
  }, []);
  const saveProject = async (project: Project): Promise<ProjectResult<Project>> => {
    const active = projectProvider.current;
    if (!active) return { status: "network_error", message: "The project store is still starting up. Try again in a moment." };
    const result = await active.create(project);
    if (result.status === "success") {
      setProjectRecords((current) => [result.value, ...current.filter((item) => item.appId !== result.value.appId)]);
    }
    return result;
  };
  const updateProject = async (project: Project, expectedVersion: number): Promise<ProjectResult<Project>> => {
    const active = projectProvider.current;
    if (!active) return { status: "network_error", message: "The project store is still starting up. Try again in a moment." };
    const result = await active.update(project, expectedVersion);
    if (result.status === "success") {
      setProjectRecords((current) => current.map((item) => (item.appId === result.value.appId ? result.value : item)));
    }
    return result;
  };
  const projectRecordsRef = useRef<Project[]>(projectRecords);
  useEffect(() => {
    // Keeps the Work Record DataProvider's internal reference-project set in sync with the
    // currently-known durable Projects (see lib/data-provider.ts "setDurableProjects"), so
    // createWorkRecord/updateWorkRecord validation never rejects a project the UI itself offers
    // as selectable. Also updates the ref the Work Record load effect above reads from — an
    // effect (not render) access, so neither ordering between the two provider-load effects can
    // leave the durable-project set stale.
    projectRecordsRef.current = projectRecords;
    provider.current?.setDurableProjects(projectRecords);
  }, [projectRecords]);
  const contactProvider = useRef<ContactProvider | null>(contactDataProvider ?? null);
  const [contactRecords, setContactRecords] = useState<Contact[]>([]);
  const [contactLoadFailed, setContactLoadFailed] = useState(false);
  useEffect(() => {
    // Independent of every other load: a failure here never affects Work Records, Inbox
    // Intelligence, Meeting Records, or Projects, and vice versa. Read-only list() call — zero
    // Anthropic calls, matching docs/AI_HANDOFF.md "Durable Contacts (Patch 8B)" cost discipline.
    let live = true;
    (async () => {
      if (!contactProvider.current) {
        const selected = await selectContactProvider();
        contactProvider.current = selected.provider;
      }
      const result = await contactProvider.current.list();
      if (!live) return;
      if (result.status === "success") setContactRecords(result.value);
      else setContactLoadFailed(true);
    })();
    return () => {
      live = false;
    };
  }, []);
  const saveContact = async (contact: Contact): Promise<ContactResult<Contact>> => {
    const active = contactProvider.current;
    if (!active) return { status: "network_error", message: "The contact store is still starting up. Try again in a moment." };
    const result = await active.create(contact);
    if (result.status === "success") {
      setContactRecords((current) => [result.value, ...current.filter((item) => item.appId !== result.value.appId)]);
    }
    return result;
  };
  const updateContact = async (contact: Contact, expectedVersion: number): Promise<ContactResult<Contact>> => {
    const active = contactProvider.current;
    if (!active) return { status: "network_error", message: "The contact store is still starting up. Try again in a moment." };
    const result = await active.update(contact, expectedVersion);
    if (result.status === "success") {
      setContactRecords((current) => current.map((item) => (item.appId === result.value.appId ? result.value : item)));
    }
    return result;
  };
  const contactRecordsRef = useRef<Contact[]>(contactRecords);
  useEffect(() => {
    // Keeps the Work Record DataProvider's internal reference-contact set in sync with the
    // currently-known durable Contacts (see lib/data-provider.ts "setDurableContacts"), so
    // createWorkRecord/updateWorkRecord validation never rejects a contact the UI itself offers
    // as selectable. Mirrors the Project mechanism exactly — see its comment above.
    contactRecordsRef.current = contactRecords;
    provider.current?.setDurableContacts(contactRecords);
  }, [contactRecords]);
  const organizationProvider = useRef<OrganizationProvider | null>(organizationDataProvider ?? null);
  const [organizationRecords, setOrganizationRecords] = useState<Organization[]>([]);
  const [organizationLoadFailed, setOrganizationLoadFailed] = useState(false);
  useEffect(() => {
    // Independent of every other load: a failure here never affects Work Records, Inbox
    // Intelligence, Meeting Records, Projects, or Contacts, and vice versa. Read-only list()
    // call — zero Anthropic calls, matching docs/AI_HANDOFF.md "Durable Organizations (Patch
    // 8E)" cost discipline.
    let live = true;
    (async () => {
      if (!organizationProvider.current) {
        const selected = await selectOrganizationProvider();
        organizationProvider.current = selected.provider;
      }
      const result = await organizationProvider.current.list();
      if (!live) return;
      if (result.status === "success") setOrganizationRecords(result.value);
      else setOrganizationLoadFailed(true);
    })();
    return () => {
      live = false;
    };
  }, []);
  const saveOrganization = async (organization: Organization): Promise<OrganizationResult<Organization>> => {
    const active = organizationProvider.current;
    if (!active) return { status: "network_error", message: "The organization store is still starting up. Try again in a moment." };
    const result = await active.create(organization);
    if (result.status === "success") {
      setOrganizationRecords((current) => [result.value, ...current.filter((item) => item.appId !== result.value.appId)]);
    }
    return result;
  };
  const updateOrganization = async (organization: Organization, expectedVersion: number): Promise<OrganizationResult<Organization>> => {
    const active = organizationProvider.current;
    if (!active) return { status: "network_error", message: "The organization store is still starting up. Try again in a moment." };
    const result = await active.update(organization, expectedVersion);
    if (result.status === "success") {
      setOrganizationRecords((current) => current.map((item) => (item.appId === result.value.appId ? result.value : item)));
    }
    return result;
  };
  const organizationRecordsRef = useRef<Organization[]>(organizationRecords);
  useEffect(() => {
    // Keeps the Work Record DataProvider's internal reference-organization set in sync with the
    // currently-known durable Organizations (see lib/data-provider.ts "setDurableOrganizations"),
    // so createWorkRecord/updateWorkRecord validation never rejects an organization the UI
    // itself offers as selectable. Mirrors the Project/Contact mechanism exactly.
    organizationRecordsRef.current = organizationRecords;
    provider.current?.setDurableOrganizations(organizationRecords);
  }, [organizationRecords]);
  const allProjects: Project[] = references ? [...references.projects, ...projectRecords] : projectRecords;
  const allContacts: Contact[] = references ? [...references.contacts, ...contactRecords] : contactRecords;
  const allOrganizations: Organization[] = references ? [...references.organizations, ...organizationRecords] : organizationRecords;
  const effectiveReferences: ReferenceData | null = references
    ? { ...references, projects: allProjects, contacts: allContacts, organizations: allOrganizations }
    : null;
  const saveInboxRecord = async (record: InboxIntelligenceRecord): Promise<InboxIntelligenceResult<InboxIntelligenceRecord>> => {
    const active = inboxProvider.current;
    if (!active) return { status: "network_error", message: "The inbox store is still starting up. Try again in a moment." };
    const result = await active.create(record);
    if (result.status === "success") {
      setInboxRecords((current) => [result.value, ...current.filter((item) => item.appId !== result.value.appId)]);
    }
    return result;
  };
  const updateInboxRecord = async (
    record: InboxIntelligenceRecord,
    expectedVersion: number,
  ): Promise<InboxIntelligenceResult<InboxIntelligenceRecord>> => {
    const active = inboxProvider.current;
    if (!active) return { status: "network_error", message: "The inbox store is still starting up. Try again in a moment." };
    const result = await active.update(record, expectedVersion);
    if (result.status === "success") {
      setInboxRecords((current) => current.map((item) => (item.appId === result.value.appId ? result.value : item)));
    }
    return result;
  };
  useEffect(() => {
    let live = true;
    const load = async () => {
      const loadFrom = async (source: DataProvider) => {
        const [result, projects, organizations, contacts, categories, deliverables, reportingConfig, settings] = await Promise.all([
          source.getWorkRecords(), source.getProjects(), source.getOrganizations(), source.getContacts(), source.getCategories(),
          source.getDeliverables(), source.getReportingConfig(), source.getSystemSettings(),
        ]);
        if (result.status !== "success") throw new Error(result.status);
        if (live) {
          setRecords(result.value);
          setReferences({ projects, organizations, contacts, categories, deliverables, reportingConfig, settings });
        }
      };
      let kind: ActiveProviderKind = "memory";
      try {
        if (!provider.current) {
          const selected = await selectDataProvider();
          provider.current = selected.provider;
          kind = selected.kind;
        }
        // Patch 7 / 8B / 8E: sync whatever durable Projects/Contacts/Organizations are already
        // known (via the refs, not the possibly-stale state closures — see the
        // projectRecordsRef/contactRecordsRef/organizationRecordsRef effects below) the moment
        // this provider becomes available, so createWorkRecord/updateWorkRecord validation is
        // correct regardless of which provider-load effect resolves first.
        provider.current.setDurableProjects(projectRecordsRef.current);
        provider.current.setDurableContacts(contactRecordsRef.current);
        provider.current.setDurableOrganizations(organizationRecordsRef.current);
        await loadFrom(provider.current);
      } catch {
        // Selecting or loading from the active provider failed (including a SharePoint
        // initialization failure). Fall back to the safe in-memory prototype without
        // touching whatever the failed provider holds. This is the same non-durable
        // "memory" kind selectDataProvider() itself returns when there is no signed-in
        // Microsoft session — there is no separate durable fallback to distinguish it from.
        const fallback = new PrototypeFallbackProvider();
        provider.current = fallback;
        kind = "memory";
        await loadFrom(fallback);
      } finally {
        if (live) {
          setStorageMode(kind);
          setLoading(false);
        }
      }
    };
    void load();
    return () => {
      live = false;
    };
  }, []);
  const today = todayIso();
  const todayRecords = records.filter(
    (record) => record.activityDate === today,
  );
  const totalToday = todayRecords.reduce(
    (sum, record) => sum + record.durationMinutes,
    0,
  );
  const followups = records
    .filter((record) => record.followUpNeeded && record.followUpDate)
    .sort((a, b) =>
      String(a.followUpDate).localeCompare(String(b.followUpDate)),
    );
  const openLog = (record?: WorkRecord, onSaved?: (saved: WorkRecord) => void) => {
    const next = record ? structuredClone(record) : emptyRecord();
    opener.current = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    onSavedRef.current = onSaved ?? null;
    setDraft(next);
    setDraftBaseline(JSON.stringify(next));
    setStep(1);
    setSaveError("");
    setLogging(true);
  };
  const patchDraft = (patch: Partial<WorkRecord>) =>
    setDraft((current) => ({ ...current, ...patch }));
  const toggle = (
    key: "projectIds" | "organizationIds" | "contactIds" | "categoryIds",
    id: string,
  ) =>
    setDraft((current) => ({
      ...current,
      [key]: current[key].includes(id)
        ? current[key].filter((value) => value !== id)
        : [...current[key], id],
    }));
  const save = async (another = false) => {
    if (
      !draft.title.trim() ||
      !draft.activityType ||
      draft.durationMinutes <= 0
    ) {
      setSaveError("Add a title, activity type, and duration before saving.");
      return;
    }
    const activeProvider = provider.current;
    if (!activeProvider) {
      setSaveError("The data store is still starting up. Try again in a moment.");
      return;
    }
    setSaving(true);
    setSaveError("");
    const pending: WorkRecord = { ...draft, title: draft.title.trim(), metadata: { ...draft.metadata, syncState: "saving" } };
    const result = draft.metadata.version > 0
      ? await activeProvider.updateWorkRecord(pending, draft.metadata.version)
      : await activeProvider.createWorkRecord(pending);
    try {
      if (result.status !== "success") {
        setSaveError(result.status === "validation_error" ? result.errors[0]?.message ?? "Check the record and try again." : result.message);
        return;
      }
      const saved = result.value;
      setRecords((current) =>
        [saved, ...current.filter((item) => item.appId !== saved.appId)].sort(
          (a, b) => b.activityDate.localeCompare(a.activityDate),
        ),
      );
      onSavedRef.current?.(saved);
      onSavedRef.current = null;
      if (another) {
        const next = emptyRecord();
        setDraft(next);
        setDraftBaseline(JSON.stringify(next));
        setStep(1);
      } else {
        setLogging(false);
        setView("today");
      }
    } catch {
      setSaveError(
        "Sync failed. Your entry is still open—check the connection and try again.",
      );
    } finally {
      setSaving(false);
    }
  };
  const requestClose = () => {
    const dirty = JSON.stringify(draft) !== draftBaseline;
    if (dirty && !window.confirm("Discard the changes in this work entry?")) return;
    onSavedRef.current = null;
    setLogging(false);
  };
  return (
    <>
    <main className="os-shell" inert={logging ? true : undefined} aria-hidden={logging || undefined}>
      <Header view={view} setView={setView} onLog={() => openLog()} />
      <div className="os-body">
        <SideNav view={view} setView={setView} onLog={() => openLog()} />
        <section className="screen" aria-live="polite">
          {storageMode === "memory" && (
            <div className="storage-banner">
              <span>Preview mode</span> Changes in this session are not saved. Sign in
              with Microsoft (top right) to connect SharePoint and save durably.
            </div>
          )}
          {loading || !effectiveReferences ? (
            <div className="loading-card">Preparing your workspace…</div>
          ) : view === "home" ? (
            <Home
              records={records}
              todayRecords={todayRecords}
              totalToday={totalToday}
              followups={followups}
              openLog={openLog}
              setView={setView}
              references={effectiveReferences}
              inboxSummary={computeInboxIntelligenceSummary(inboxRecords)}
              needsAttentionCount={filterNeedsAttention(inboxRecords).length}
              needsAttentionItems={selectNeedsAttention(inboxRecords)}
              waitingItems={selectWaiting(inboxRecords)}
              inboxLoadFailed={inboxLoadFailed}
            />
          ) : view === "today" ? (
            <Today
              records={todayRecords}
              total={totalToday}
              followups={followups}
              openLog={openLog}
              references={effectiveReferences}
            />
          ) : view === "history" ? (
            <History
              records={records}
              search={search}
              setSearch={setSearch}
              openLog={openLog}
              references={effectiveReferences}
            />
          ) : view === "projects" ? (
            <Projects
              records={records}
              setView={setView}
              projects={allProjects}
              loadFailed={projectLoadFailed}
              saveProject={saveProject}
              updateProject={updateProject}
            />
          ) : view === "contacts" ? (
            <Contacts
              contacts={allContacts}
              organizations={effectiveReferences.organizations}
              projects={allProjects}
              records={records}
              loadFailed={contactLoadFailed}
              saveContact={saveContact}
              updateContact={updateContact}
              openLog={openLog}
            />
          ) : view === "organizations" ? (
            <Organizations
              organizations={allOrganizations}
              loadFailed={organizationLoadFailed}
              saveOrganization={saveOrganization}
              updateOrganization={updateOrganization}
            />
          ) : view === "orbit" ? (
            <Orbit records={records} references={effectiveReferences} />
          ) : view === "inbox" ? (
            <InboxIntelligence
              references={effectiveReferences}
              openLog={openLog}
              createDraftRecord={emptyRecord}
              records={inboxRecords}
              saveRecord={saveInboxRecord}
              updateRecord={updateInboxRecord}
              saveContact={saveContact}
              updateContact={updateContact}
            />
          ) : view === "voice" ? (
            <VoiceIntelligence
              openLog={openLog}
              createDraftRecord={emptyRecord}
              references={effectiveReferences}
              saveContact={saveContact}
              updateContact={updateContact}
            />
          ) : (
            <MeetingNotes
              openLog={openLog}
              createDraftRecord={emptyRecord}
              records={meetingRecords}
              saveRecord={saveMeetingRecord}
              updateRecord={updateMeetingRecord}
              loadFailed={meetingLoadFailed}
              storageMode={storageMode}
            />
          )}
        </section>
      </div>
      <footer className="os-status">
        <span>
          <i />{" "}
          {storageMode === "sharepoint" ? "SharePoint DEV connected" : "Preview session active"}
        </span>
        <span>Log it once. Use it everywhere.</span>
      </footer>
    </main>
      {logging && effectiveReferences && (
        <LogWizard
          record={draft}
          step={step}
          setStep={setStep}
          patch={patchDraft}
          toggle={toggle}
          close={requestClose}
          save={save}
          saving={saving}
          error={saveError}
          references={effectiveReferences}
          returnFocusRef={opener}
        />
      )}
    </>
  );
}

function Header({
  view,
  setView,
  onLog,
}: {
  view: View;
  setView: (view: View) => void;
  onLog: () => void;
}) {
  return (
    <header className="os-header">
      <button
        className="brand-button"
        onClick={() => setView("home")}
        aria-label="IU Work Tracker home"
      >
        <span className="brand-symbol">IU</span>
        <span>
          <strong>Work Tracker</strong>
          <small>IU Operations</small>
        </span>
      </button>
      <div className="header-title">
        {navItems.find(([key]) => key === view)?.[2]}
      </div>
      <div className="header-actions">
        <span className="dev-chip">Development sample</span>
        <button className="header-log" onClick={onLog}>
          + Log work
        </button>
        <DevMicrosoftConnection />
      </div>
    </header>
  );
}
function SideNav({
  view,
  setView,
  onLog,
}: {
  view: View;
  setView: (view: View) => void;
  onLog: () => void;
}) {
  return (
    <aside className="side-nav">
      <button className="side-log" onClick={onLog}>
        <span>＋</span>Log work
      </button>
      <nav aria-label="Workspaces">
        {navItems.map(([key, icon, label]) => (
          <button
            key={key}
            className={view === key ? "active" : ""}
            onClick={() => setView(key)}
            aria-current={view === key ? "page" : undefined}
          >
            <span>{icon}</span>
            {label}
          </button>
        ))}
      </nav>
      <div className="side-note">
        <strong>Built for the work</strong>
        <span>Not just the report.</span>
      </div>
    </aside>
  );
}

function Home({
  records,
  todayRecords,
  totalToday,
  followups,
  openLog,
  setView,
  references,
  inboxSummary,
  needsAttentionCount,
  needsAttentionItems,
  waitingItems,
  inboxLoadFailed,
}: {
  records: WorkRecord[];
  todayRecords: WorkRecord[];
  totalToday: number;
  followups: WorkRecord[];
  openLog: (r?: WorkRecord) => void;
  setView: (v: View) => void;
  references: ReferenceData;
  inboxSummary: { openCount: number; waitingCount: number; resolvedCount: number };
  needsAttentionCount: number;
  needsAttentionItems: InboxIntelligenceRecord[];
  waitingItems: InboxIntelligenceRecord[];
  inboxLoadFailed: boolean;
}) {
  const latest = todayRecords.slice(0, 3);
  return (
    <div className="screen-inner home-screen">
      <div className="welcome-row">
        <div>
          <p className="eyebrow">
            {new Intl.DateTimeFormat("en-US", {
              weekday: "long",
              month: "long",
              day: "numeric",
            }).format(new Date())}
          </p>
          <h1>Good morning. What are you working on?</h1>
          <p className="subhead">
            Capture it once. Keep your day, projects, and reporting connected.
          </p>
        </div>
        <button className="primary-action" onClick={() => openLog()}>
          <span>＋</span> Log work
        </button>
      </div>
      <div className="home-grid">
        <section className="panel today-panel">
          <div className="panel-heading">
            <div>
              <p className="eyebrow">Your day at a glance</p>
              <h2>Today</h2>
            </div>
            <button className="text-action" onClick={() => setView("today")}>
              Open today →
            </button>
          </div>
          <div className="metric-strip">
            <Metric value={String(todayRecords.length)} label="activities" />
            <Metric value={hours(totalToday)} label="recorded" />
            <Metric value={String(followups.length)} label="follow-ups" />
          </div>
          <div className="activity-list">
            {latest.length ? (
              latest.map((record) => (
                <RecordRow
                  key={record.appId}
                  record={record}
                  onClick={() => openLog(record)}
                  organizations={references.organizations}
                />
              ))
            ) : (
              <Empty
                title="Nothing logged yet"
                copy="Start with one quick work record."
                action="Log work"
                onAction={() => openLog()}
              />
            )}
          </div>
        </section>
        <aside className="panel next-panel">
          <div className="panel-heading">
            <div>
              <p className="eyebrow">Keep momentum</p>
              <h2>Next up</h2>
            </div>
            <span className="count-badge">{followups.length}</span>
          </div>
          {followups.slice(0, 3).map((record) => (
            <button
              className="followup-row"
              key={record.appId}
              onClick={() => openLog(record)}
            >
              <span className="date-tile">
                <b>{record.followUpDate?.slice(-2)}</b>
                {new Intl.DateTimeFormat("en-US", { month: "short" })
                  .format(new Date(`${record.followUpDate}T12:00:00`))
                  .toUpperCase()}
              </span>
              <span>
                <strong>{record.nextStep || "Follow up"}</strong>
                <small>{record.title}</small>
              </span>
            </button>
          ))}
          {!followups.length && (
            <p className="muted-copy">No open follow-ups.</p>
          )}
          <button className="panel-button" onClick={() => setView("today")}>
            View today
          </button>
        </aside>
      </div>
      <ActionCenter
        summary={inboxSummary}
        needsAttentionCount={needsAttentionCount}
        needsAttentionItems={needsAttentionItems}
        waitingItems={waitingItems}
        loadFailed={inboxLoadFailed}
        setView={setView}
      />
      <div className="command-grid">
        <Command
          accent
          icon="＋"
          title="Log work"
          copy="Capture an activity"
          onClick={() => openLog()}
        />
        <Command
          icon="≡"
          title="Work history"
          copy={`${records.length} records to explore`}
          onClick={() => setView("history")}
        />
        <Command
          icon="▤"
          title="Projects"
          copy="See connected work"
          onClick={() => setView("projects")}
        />
        <Command
          icon="◎"
          title="STEM / ORBIT"
          copy="Review reporting progress"
          onClick={() => setView("orbit")}
        />
        <Command
          icon="✉"
          title="Inbox Intelligence"
          copy={needsAttentionCount > 0 ? `${needsAttentionCount} need attention` : "Turn a pasted email into a work record"}
          onClick={() => setView("inbox")}
        />
        <Command
          icon="🎙"
          title="Voice Intelligence"
          copy="Turn a transcript into actionable intelligence"
          onClick={() => setView("voice")}
        />
        <Command
          icon="📝"
          title="Meeting Notes"
          copy="Agenda, notes, and reviewable meeting intelligence"
          onClick={() => setView("meeting")}
        />
      </div>
    </div>
  );
}
function Today({
  records,
  total,
  followups,
  openLog,
  references,
}: {
  records: WorkRecord[];
  total: number;
  followups: WorkRecord[];
  openLog: (r?: WorkRecord) => void;
  references: ReferenceData;
}) {
  return (
    <div className="screen-inner">
      <PageHeading
        eyebrow="Daily workspace"
        title="Today"
        copy="A clean answer to: What did I actually do today?"
        action="+ Log work"
        onAction={() => openLog()}
      />
      <div className="metric-cards">
        <MetricCard value={hours(total)} label="Time recorded" tone="teal" />
        <MetricCard
          value={String(records.length)}
          label="Activities"
          tone="blue"
        />
        <MetricCard
          value={String(followups.length)}
          label="Open follow-ups"
          tone="coral"
        />
      </div>
      <div className="two-column">
        <section className="panel list-panel">
          <div className="panel-heading">
            <h2>Today’s work</h2>
            <span className="sample-label">
              Development sample + your entries
            </span>
          </div>
          {records.length ? (
            records.map((record) => (
              <RecordRow
                key={record.appId}
                record={record}
                onClick={() => openLog(record)}
                organizations={references.organizations}
              />
            ))
          ) : (
            <Empty
              title="Your day is still open"
              copy="Log the first thing you worked on."
              action="Log work"
              onAction={() => openLog()}
            />
          )}
        </section>
        <aside className="panel focus-panel">
          <p className="eyebrow">Quick reflection</p>
          <h2>Your work, in one place</h2>
          <p>
            {records.length
              ? `${records.length} activities account for ${hours(total)} today. Add outcomes while they are fresh, then let the same record support later reporting.`
              : "A 30–60 second entry now saves reconstruction later."}
          </p>
          <button className="panel-button" onClick={() => openLog()}>
            Add another activity
          </button>
        </aside>
      </div>
    </div>
  );
}
function History({
  records,
  search,
  setSearch,
  openLog,
  references,
}: {
  records: WorkRecord[];
  search: string;
  setSearch: (v: string) => void;
  openLog: (r: WorkRecord) => void;
  references: ReferenceData;
}) {
  const filtered = records.filter((record) =>
    `${record.title} ${record.description} ${record.activityType}`
      .toLowerCase()
      .includes(search.toLowerCase()),
  );
  return (
    <div className="screen-inner">
      <PageHeading
        eyebrow="Universal work records"
        title="Work history"
        copy="Find, review, and improve the records that feed every workspace."
      />
      <label className="search-box">
        <span>⌕</span>
        <input
          value={search}
          onChange={(event) => setSearch(event.target.value)}
          placeholder="Search activities, types, or notes"
          aria-label="Search work history"
        />
      </label>
      <section className="panel history-panel">
        <div className="history-head">
          <span>Activity</span>
          <span>Connections</span>
          <span>Time</span>
          <span>Reporting</span>
          <span />
        </div>
        {filtered.map((record) => (
          <button
            className="history-row"
            key={record.appId}
            onClick={() => openLog(record)}
          >
            <span>
              <strong>{record.title}</strong>
              <small>
                {niceDate(record.activityDate)} · {record.activityType}
                {record.isSample && <em>Sample</em>}
              </small>
            </span>
            <span>
              <strong>
                {record.projectIds[0]
                  ? entityName(references.projects, record.projectIds[0])
                  : "No project"}
              </strong>
              <small>
                {record.organizationIds[0]
                  ? entityName(references.organizations, record.organizationIds[0])
                  : record.engagementScope === "regional"
                    ? "Regional scope"
                    : record.engagementScope === "allDistricts"
                      ? "All districts"
                      : "No organization"}
              </small>
            </span>
            <span>
              <strong>{hours(record.durationMinutes)}</strong>
            </span>
            <span>
              {record.orbit.reportable ? (
                <b className="orbit-tag">
                  Deliverable {record.orbit.primaryDeliverable}
                </b>
              ) : (
                <small>IU work</small>
              )}
            </span>
            <span className="row-arrow">→</span>
          </button>
        ))}
        {!filtered.length && (
          <Empty
            title="No matching work"
            copy="Try a different search phrase."
          />
        )}
      </section>
    </div>
  );
}
// Cycled deterministically for new durable projects (see globals.css .project-mark.<color>) —
// the create form deliberately has no color picker (not part of the Patch 7 spec), so every
// new project keeps the exact same visual card design as the five seeded ones.
const PROJECT_COLORS = ["blue", "coral", "lime", "purple", "yellow"] as const;

function emptyProjectDraft(existingCount: number): Project {
  return buildProjectDraft({
    appId: crypto.randomUUID(),
    name: "",
    description: "",
    status: "planning",
    color: PROJECT_COLORS[existingCount % PROJECT_COLORS.length],
    startDate: null,
    targetDate: null,
    stemOrbit: false,
  });
}

function Projects({
  records,
  setView,
  projects,
  loadFailed,
  saveProject,
  updateProject,
}: {
  records: WorkRecord[];
  setView: (v: View) => void;
  projects: Project[];
  loadFailed: boolean;
  saveProject: (project: Project) => Promise<ProjectResult<Project>>;
  updateProject: (project: Project, expectedVersion: number) => Promise<ProjectResult<Project>>;
}) {
  const [modalProject, setModalProject] = useState<Project | null>(null);

  return (
    <div className="screen-inner">
      <PageHeading
        eyebrow="Connected containers"
        title="Projects"
        copy="See the activity, time, and outcomes accumulating around each initiative."
      />
      {loadFailed && <p className="muted-copy">Projects could not be fully loaded. Try reloading the page.</p>}
      <div className="project-grid">
        {projects.map((project) => {
          const related = records.filter((record) =>
            record.projectIds.includes(project.appId),
          );
          const mins = related.reduce(
            (sum, record) => sum + record.durationMinutes,
            0,
          );
          // Only a durable project (created/loaded through the Project provider — see
          // docs/AI_HANDOFF.md "Durable Projects (Patch 7)") is editable. The five seeded
          // reference-data projects have no `metadata` and remain read-only in V1, since they
          // were never approved for migration into the durable store — see the Patch 7
          // approval gate's "Seeded-project migration strategy."
          const isDurable = Boolean(project.metadata);
          return (
            <article className="project-card" key={project.appId}>
              <span className={`project-mark ${project.color}`} />
              <div className="project-title">
                <span className="status-chip">{project.status}</span>
                <h2>{project.name}</h2>
                <p>{project.description}</p>
              </div>
              <div className="project-metrics">
                <span>
                  <b>{related.length}</b> records
                </span>
                <span>
                  <b>{hours(mins)}</b> invested
                </span>
              </div>
              <div className="project-card-actions">
                <button onClick={() => setView("history")}>
                  View connected work →
                </button>
                {isDurable && (
                  <button type="button" className="ghost-button" onClick={() => setModalProject(project)}>
                    Edit
                  </button>
                )}
              </div>
            </article>
          );
        })}
        <button type="button" className="planned-note" aria-label="Create Project" onClick={() => setModalProject(emptyProjectDraft(projects.length))}>
          <span>＋</span>
          <div>
            <strong>Create Project</strong>
            <p>It becomes selectable in Log Work immediately, with totals derived from Work Records.</p>
          </div>
        </button>
      </div>
      {modalProject && (
        <ProjectFormModal
          project={modalProject}
          onCancel={() => setModalProject(null)}
          onSaved={() => setModalProject(null)}
          saveProject={saveProject}
          updateProject={updateProject}
        />
      )}
    </div>
  );
}

/**
 * A single compact overlay handles both Create and Edit — see docs/AI_HANDOFF.md "Durable
 * Projects (Patch 7)". Create vs. update is decided the same way every other durable resource
 * in this codebase decides it: `project.metadata.version > 0` means the project already has a
 * durable identity, so submitting routes through updateProject(); otherwise saveProject().
 * Updating a project never touches its connected Work Records — this component only ever
 * calls the Project provider.
 */
function ProjectFormModal({
  project,
  onCancel,
  onSaved,
  saveProject,
  updateProject,
}: {
  project: Project;
  onCancel: () => void;
  onSaved: () => void;
  saveProject: (project: Project) => Promise<ProjectResult<Project>>;
  updateProject: (project: Project, expectedVersion: number) => Promise<ProjectResult<Project>>;
}) {
  const [draft, setDraft] = useState<Project>(project);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const nameRef = useRef<HTMLInputElement>(null);
  const isEditing = (draft.metadata?.version ?? 0) > 0;

  useEffect(() => {
    nameRef.current?.focus();
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onCancel();
      }
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, [onCancel]);

  const patch = (change: Partial<Project>) => setDraft((current) => ({ ...current, ...change }));

  const submit = async () => {
    if (saving) return;
    const shapeIssues = validateProjectShape(draft);
    if (shapeIssues.length) {
      setError(shapeIssues[0].message);
      return;
    }
    setSaving(true);
    setError("");
    const result = isEditing ? await updateProject(draft, draft.metadata!.version) : await saveProject(draft);
    setSaving(false);
    if (result.status !== "success") {
      setError(result.status === "validation_error" ? (result.errors[0]?.message ?? "Check the project and try again.") : result.message);
      return;
    }
    onSaved();
  };

  return (
    <div className="modal-backdrop" role="presentation">
      <section className="project-modal" role="dialog" aria-modal="true" aria-labelledby="project-modal-title">
        <header className="log-header">
          <h2 id="project-modal-title">{isEditing ? "Edit Project" : "Create Project"}</h2>
          <button onClick={onCancel} aria-label="Close">
            ×
          </button>
        </header>
        <div className="log-content">
          <div className="form-stack">
            <label>
              <span>
                Project name <b>*</b>
              </span>
              <input ref={nameRef} value={draft.name} onChange={(event) => patch({ name: event.target.value })} placeholder="e.g. STEM Ecosystem" />
            </label>
            <label>
              <span>Description</span>
              <textarea value={draft.description} onChange={(event) => patch({ description: event.target.value })} placeholder="A sentence is enough." rows={2} />
            </label>
            <label>
              <span>Status</span>
              <select value={draft.status} onChange={(event) => patch({ status: event.target.value as ProjectStatus })}>
                {PROJECT_STATUSES.map((status) => (
                  <option key={status} value={status}>
                    {status[0].toUpperCase() + status.slice(1)}
                  </option>
                ))}
              </select>
            </label>
            <div className="form-two">
              <label>
                <span>Start date</span>
                <input type="date" value={draft.startDate ?? ""} onChange={(event) => patch({ startDate: event.target.value || null })} />
              </label>
              <label>
                <span>Target date</span>
                <input type="date" value={draft.targetDate ?? ""} onChange={(event) => patch({ targetDate: event.target.value || null })} />
              </label>
            </div>
            <div className="toggle-line">
              <input
                aria-label="STEM / ORBIT connection"
                type="checkbox"
                checked={draft.stemOrbit ?? false}
                onChange={(event) => patch({ stemOrbit: event.target.checked })}
              />
              <span>
                <strong>STEM / ORBIT connection</strong>
                <small>Optional — does not classify any Work Record automatically.</small>
              </span>
            </div>
          </div>
        </div>
        {error && (
          <div className="form-error" role="alert">
            {error}
          </div>
        )}
        <footer className="log-footer">
          <button className="ghost-button" onClick={onCancel}>
            Cancel
          </button>
          <button className="primary-action" onClick={() => void submit()} disabled={saving || !draft.name.trim()}>
            {saving ? "Saving…" : isEditing ? "Save Changes" : "Create Project"}
          </button>
        </footer>
      </section>
    </div>
  );
}

const ORGANIZATION_TYPE_LABELS: Record<OrganizationTypeValue, string> = {
  district: "District",
  partner: "Partner",
  iu: "IU",
};

function emptyOrganizationDraft(): Organization {
  return buildOrganizationDraft({ appId: crypto.randomUUID(), name: "", type: "partner" });
}

/**
 * Patch 8E — Durable Organizations, identity/basic metadata only: directory + Create/Edit, no
 * connected-work metrics, no detail screen. District remains Organization.type === "district" —
 * there is no separate District model, provider, or list. Seed/reference Organizations (no
 * `metadata`) stay read-only, mirroring the Project/Contact precedent exactly.
 */
function Organizations({
  organizations,
  loadFailed,
  saveOrganization,
  updateOrganization,
}: {
  organizations: Organization[];
  loadFailed: boolean;
  saveOrganization: (organization: Organization) => Promise<OrganizationResult<Organization>>;
  updateOrganization: (organization: Organization, expectedVersion: number) => Promise<OrganizationResult<Organization>>;
}) {
  const [search, setSearch] = useState("");
  const [modalOrganization, setModalOrganization] = useState<Organization | null>(null);

  // Case-insensitive substring match only — the same philosophy already used for
  // Contacts/History search. No fuzzy/semantic matching.
  const filtered = organizations.filter((organization) =>
    `${organization.name} ${ORGANIZATION_TYPE_LABELS[organization.type]}`.toLowerCase().includes(search.toLowerCase()),
  );

  return (
    <div className="screen-inner">
      <PageHeading
        eyebrow="Districts and partners"
        title="Organizations"
        copy="The districts, partners, and IU entities your work connects to."
      />
      <div className="log-footer" style={{ padding: 0, marginBottom: 13 }}>
        <label className="search-box">
          <span>⌕</span>
          <input
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            placeholder="Search name or type"
            aria-label="Search organizations"
          />
        </label>
      </div>
      {loadFailed && <p className="muted-copy">Organizations could not be fully loaded. Try reloading the page.</p>}
      <div className="project-grid">
        {filtered.map((organization) => {
          // Only a durable organization (created/loaded through the Organization provider) is
          // editable. The four seeded reference-data organizations have no `metadata` and
          // remain read-only in V1 — mirrors the Project/Contact precedent exactly.
          const isDurable = Boolean(organization.metadata);
          return (
            <article className="project-card" key={organization.appId}>
              <div className="project-title">
                <span className="status-chip">{ORGANIZATION_TYPE_LABELS[organization.type]}</span>
                <h2>{organization.name}</h2>
              </div>
              {isDurable && (
                <div className="project-card-actions">
                  <span />
                  <button type="button" className="ghost-button" onClick={() => setModalOrganization(organization)}>
                    Edit
                  </button>
                </div>
              )}
            </article>
          );
        })}
        <button type="button" className="planned-note" aria-label="Add Organization" onClick={() => setModalOrganization(emptyOrganizationDraft())}>
          <span>＋</span>
          <div>
            <strong>Add Organization</strong>
            <p>Becomes selectable in Log Work and Contacts immediately.</p>
          </div>
        </button>
      </div>
      {!filtered.length && <Empty title="No matching organizations" copy="Try a different search phrase." />}
      {modalOrganization && (
        <OrganizationFormModal
          organization={modalOrganization}
          organizations={organizations}
          onCancel={() => setModalOrganization(null)}
          onSaved={() => setModalOrganization(null)}
          saveOrganization={saveOrganization}
          updateOrganization={updateOrganization}
        />
      )}
    </div>
  );
}

/**
 * A single compact overlay handles both Create and Edit — mirrors ContactFormModal/
 * ProjectFormModal exactly. Duplicate detection is conservative and deterministic (trim +
 * collapse whitespace + lowercase, exact match only — never fuzzy): a name-only match is an
 * informational warning that never blocks Save, since Organization has no email field to serve
 * as a stronger signal.
 */
function OrganizationFormModal({
  organization,
  organizations,
  onCancel,
  onSaved,
  saveOrganization,
  updateOrganization,
}: {
  organization: Organization;
  organizations: Organization[];
  onCancel: () => void;
  onSaved: () => void;
  saveOrganization: (organization: Organization) => Promise<OrganizationResult<Organization>>;
  updateOrganization: (organization: Organization, expectedVersion: number) => Promise<OrganizationResult<Organization>>;
}) {
  const [draft, setDraft] = useState<Organization>(organization);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const nameRef = useRef<HTMLInputElement>(null);
  const isEditing = (draft.metadata?.version ?? 0) > 0;

  useEffect(() => {
    nameRef.current?.focus();
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onCancel();
      }
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, [onCancel]);

  const patch = (change: Partial<Organization>) => setDraft((current) => ({ ...current, ...change }));

  const otherOrganizations = organizations.filter((item) => item.appId !== draft.appId);
  const nameDuplicate = draft.name.trim()
    ? otherOrganizations.find((item) => normalizeOrganizationName(item.name) === normalizeOrganizationName(draft.name))
    : undefined;

  const submit = async () => {
    if (saving) return;
    const shapeIssues = validateOrganizationShape(draft);
    if (shapeIssues.length) {
      setError(shapeIssues[0].message);
      return;
    }
    setSaving(true);
    setError("");
    const result = isEditing ? await updateOrganization(draft, draft.metadata!.version) : await saveOrganization(draft);
    setSaving(false);
    if (result.status !== "success") {
      setError(result.status === "validation_error" ? (result.errors[0]?.message ?? "Check the organization and try again.") : result.message);
      return;
    }
    onSaved();
  };

  return (
    <div className="modal-backdrop" role="presentation">
      <section className="project-modal" role="dialog" aria-modal="true" aria-labelledby="organization-modal-title">
        <header className="log-header">
          <h2 id="organization-modal-title">{isEditing ? "Edit Organization" : "Add Organization"}</h2>
          <button onClick={onCancel} aria-label="Close">
            ×
          </button>
        </header>
        <div className="log-content">
          <div className="form-stack">
            <label>
              <span>
                Name <b>*</b>
              </span>
              <input
                ref={nameRef}
                value={draft.name}
                onChange={(event) => patch({ name: event.target.value })}
                placeholder="e.g. North Valley SD"
              />
            </label>
            {nameDuplicate && (
              <p className="muted-copy" role="status">
                Another organization is already named &ldquo;{nameDuplicate.name}&rdquo;. This is just a heads up — Save is not blocked.
              </p>
            )}
            <label>
              <span>Type</span>
              <select value={draft.type} onChange={(event) => patch({ type: event.target.value as OrganizationTypeValue })}>
                {ORGANIZATION_TYPES.map((type) => (
                  <option key={type} value={type}>
                    {ORGANIZATION_TYPE_LABELS[type]}
                  </option>
                ))}
              </select>
            </label>
          </div>
        </div>
        {error && (
          <div className="form-error" role="alert">
            {error}
          </div>
        )}
        <footer className="log-footer">
          <button className="ghost-button" onClick={onCancel}>
            Cancel
          </button>
          <button className="primary-action" onClick={() => void submit()} disabled={saving || !draft.name.trim()}>
            {saving ? "Saving…" : isEditing ? "Save Changes" : "Add Organization"}
          </button>
        </footer>
      </section>
    </div>
  );
}

/**
 * Patch 8B — Durable Contacts, identity/basic metadata. Patch 8C adds Contact Detail: opening a
 * card derives a relationship summary from records already stored elsewhere (WorkRecord.
 * contactIds → WorkRecord.projectIds → Project) — see lib/contact-relationships.ts and
 * docs/AI_HANDOFF.md "Contact connected work (Patch 8C)". No relationship data is duplicated
 * onto Contact.
 */
function Contacts({
  contacts,
  organizations,
  projects,
  records,
  loadFailed,
  saveContact,
  updateContact,
  openLog,
}: {
  contacts: Contact[];
  organizations: Organization[];
  projects: Project[];
  records: WorkRecord[];
  loadFailed: boolean;
  saveContact: (contact: Contact) => Promise<ContactResult<Contact>>;
  updateContact: (contact: Contact, expectedVersion: number) => Promise<ContactResult<Contact>>;
  openLog: (record?: WorkRecord) => void;
}) {
  const [search, setSearch] = useState("");
  const [showArchived, setShowArchived] = useState(false);
  const [modalContact, setModalContact] = useState<Contact | null>(null);
  const [openContactId, setOpenContactId] = useState<string | null>(null);
  const [showImport, setShowImport] = useState(false);

  const organizationName = (id: string | null) =>
    id ? (organizations.find((org) => org.appId === id)?.name ?? "Unknown organization") : "";

  if (showImport) {
    return (
      <ContactImport
        contacts={contacts}
        organizations={organizations}
        saveContact={saveContact}
        updateContact={updateContact}
        onBack={() => setShowImport(false)}
      />
    );
  }

  const openContact = openContactId ? contacts.find((contact) => contact.appId === openContactId) : undefined;
  if (openContact) {
    return (
      <ContactDetail
        contact={openContact}
        organizationName={organizationName(openContact.organizationId)}
        records={records}
        projects={projects}
        onBack={() => setOpenContactId(null)}
        onEdit={() => setModalContact(openContact)}
        openLog={openLog}
      />
    );
  }

  // Case-insensitive substring match only — the same philosophy already used for Work
  // History's search (see the `History` component above). No fuzzy/semantic matching.
  const filtered = contacts.filter((contact) => {
    if (!showArchived && contact.status === "archived") return false;
    const haystack = `${contact.displayName} ${contact.role ?? ""} ${contact.email ?? ""} ${organizationName(contact.organizationId)}`.toLowerCase();
    return haystack.includes(search.toLowerCase());
  });

  return (
    <div className="screen-inner">
      <PageHeading
        eyebrow="People and organizations"
        title="Contacts"
        copy="Who you work with — matched to work you've already logged, not a CRM."
        action="Import Contacts"
        onAction={() => setShowImport(true)}
      />
      <div className="log-footer" style={{ padding: 0, marginBottom: 13 }}>
        <label className="search-box">
          <span>⌕</span>
          <input
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            placeholder="Search name, role, email, or organization"
            aria-label="Search contacts"
          />
        </label>
        <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 10, color: "var(--muted)" }}>
          <input type="checkbox" checked={showArchived} onChange={(event) => setShowArchived(event.target.checked)} />
          Show archived
        </label>
      </div>
      {loadFailed && <p className="muted-copy">Contacts could not be fully loaded. Try reloading the page.</p>}
      <div className="project-grid">
        {filtered.map((contact) => {
          // Only a durable contact (created/loaded through the Contact provider) is editable.
          // The three seeded reference-data contacts have no `metadata` and remain read-only
          // in V1 — mirrors the Project precedent exactly. Detail is available for both.
          const isDurable = Boolean(contact.metadata);
          const subtitle = [contact.role, organizationName(contact.organizationId)].filter(Boolean).join(" · ");
          return (
            <article className="project-card" key={contact.appId}>
              <button
                type="button"
                className="project-title"
                onClick={() => setOpenContactId(contact.appId)}
                style={{ all: "unset", cursor: "pointer", display: "block", width: "100%" }}
              >
                <span className="status-chip">{CONTACT_STATUS_LABELS[contact.status]}</span>
                <h2>{contact.displayName}</h2>
                <p>{subtitle || "No role or organization on file"}</p>
                {contact.email && <p className="muted-copy">{contact.email}</p>}
              </button>
              {isDurable && (
                <div className="project-card-actions">
                  <span />
                  <button type="button" className="ghost-button" onClick={() => setModalContact(contact)}>
                    Edit
                  </button>
                </div>
              )}
            </article>
          );
        })}
        <button type="button" className="planned-note" aria-label="Add Contact" onClick={() => setModalContact(emptyContactDraft())}>
          <span>＋</span>
          <div>
            <strong>Add Contact</strong>
            <p>Becomes selectable in Log Work immediately.</p>
          </div>
        </button>
      </div>
      {!filtered.length && <Empty title="No matching contacts" copy="Try a different search phrase." />}
      {modalContact && (
        <ContactFormModal
          contact={modalContact}
          contacts={contacts}
          organizations={organizations}
          onCancel={() => setModalContact(null)}
          onSaved={() => setModalContact(null)}
          saveContact={saveContact}
          updateContact={updateContact}
        />
      )}
    </div>
  );
}

/**
 * Patch 8C — every value below the identity header is derived at render time via
 * buildContactRelationshipSummary() (lib/contact-relationships.ts). Nothing here is persisted
 * on Contact — reopening this screen recomputes it fresh from the currently-loaded Work
 * Records and Projects, exactly like the Projects screen's own derived totals.
 */
function ContactDetail({
  contact,
  organizationName,
  records,
  projects,
  onBack,
  onEdit,
  openLog,
}: {
  contact: Contact;
  organizationName: string;
  records: WorkRecord[];
  projects: Project[];
  onBack: () => void;
  onEdit: () => void;
  openLog: (record?: WorkRecord) => void;
}) {
  const summary = buildContactRelationshipSummary(records, projects, contact.appId);
  const isDurable = Boolean(contact.metadata);
  const subtitle = [CONTACT_STATUS_LABELS[contact.status], contact.role].filter(Boolean).join(" · ");

  return (
    <div className="screen-inner">
      <button type="button" className="ghost-button" onClick={onBack} style={{ marginLeft: 0, marginBottom: 13 }}>
        ← Back to Contacts
      </button>
      <div className="page-heading">
        <div>
          <p className="eyebrow">{organizationName || "No organization on file"}</p>
          <h1>{contact.displayName}</h1>
          <p>{subtitle}</p>
          {contact.email && <p className="muted-copy">{contact.email}</p>}
        </div>
        {isDurable && (
          <button type="button" className="primary-action" onClick={onEdit}>
            Edit Contact
          </button>
        )}
      </div>

      <section className="panel">
        <p className="eyebrow">Relationship snapshot</p>
        <div className="project-metrics" style={{ gap: 28, paddingTop: 0, borderTop: "none" }}>
          <span>
            <b>{summary.lastInteractionDate ? niceDate(summary.lastInteractionDate) : "—"}</b>
            Last interaction
          </span>
          <span>
            <b>{hours(summary.totalMinutes)}</b>
            Connected work
          </span>
          <span>
            <b>{summary.workRecordCount}</b>
            Work record{summary.workRecordCount === 1 ? "" : "s"}
          </span>
        </div>
        {summary.workRecordCount === 0 && <p className="muted-copy">No recorded interaction yet.</p>}
      </section>

      <section className="panel">
        <p className="eyebrow">Connected projects</p>
        {summary.connectedProjects.length ? (
          <div className="chip-grid">
            {summary.connectedProjects.map((project) => (
              <span key={project.appId} className="status-chip" style={{ fontSize: 9, padding: "6px 10px" }}>
                {project.name}
              </span>
            ))}
          </div>
        ) : (
          <p className="muted-copy">No connected projects yet.</p>
        )}
      </section>

      <section className="panel">
        <p className="eyebrow">Recent work</p>
        {summary.recentWorkRecords.length ? (
          <div>
            {summary.recentWorkRecords.map((record) => (
              <button className="record-row" key={record.appId} onClick={() => openLog(record)}>
                <span className={`record-dot ${record.orbit.reportable ? "orbit" : "iu"}`} />
                <span>
                  <strong>{record.title}</strong>
                  <small>
                    {niceDate(record.activityDate)} · {record.activityType}
                  </small>
                </span>
                <b>{hours(record.durationMinutes)}</b>
                <i>→</i>
              </button>
            ))}
          </div>
        ) : (
          <p className="muted-copy">No work has been logged with this contact yet.</p>
        )}
      </section>

      {contact.notes && (
        <section className="panel">
          <p className="eyebrow">Notes</p>
          <p>{contact.notes}</p>
        </section>
      )}
    </div>
  );
}

function Orbit({ records, references }: { records: WorkRecord[]; references: ReferenceData }) {
  const reportable = records.filter((record) => record.orbit.reportable);
  const poc = reportable.reduce(
    (sum, record) => sum + record.orbit.stemPocMinutes,
    0,
  );
  const tac = reportable.reduce(
    (sum, record) => sum + record.orbit.tacMinutes,
    0,
  );
  const educators = reportable.reduce(
    (sum, record) => sum + record.reach.educatorsLeaders,
    0,
  );
  const students = reportable.reduce(
    (sum, record) => sum + record.reach.studentsFamilies,
    0,
  );
  const partners = reportable.reduce(
    (sum, record) => sum + record.reach.workforceCommunity,
    0,
  );
  return (
    <div className="screen-inner">
      <PageHeading
        eyebrow="Reporting layer"
        title="STEM / ORBIT"
        copy="Reporting insight derived from normal work records—without turning the app into a reporting portal."
      />
      <div className="sample-callout">
        <strong>Development sample view</strong>
        <span>
          These illustrative totals are not IU29 performance data and are not an
          official ORBIT submission.
        </span>
      </div>
      <div className="orbit-stats">
        <MetricCard
          value={deriveReportingDays(poc, references.reportingConfig).toFixed(2)}
          label="STEM PoC days used"
          detail={`${hours(poc)} ÷ 7-hour day`}
          tone="teal"
        />
        <MetricCard
          value={hours(tac)}
          label="TaC collaboration"
          detail="Precise time retained"
          tone="purple"
        />
        <MetricCard
          value={String(educators)}
          label="Educators / leaders"
          detail="Engagement count"
          tone="blue"
        />
        <MetricCard
          value={String(students + partners)}
          label="Students + partners"
          detail={`${students} students · ${partners} partners`}
          tone="coral"
        />
      </div>
      <div className="two-column orbit-columns">
        <section className="panel deliverable-panel">
          <div className="panel-heading">
            <div>
              <p className="eyebrow">Primary classification</p>
              <h2>Deliverables A–G</h2>
            </div>
            <span>{reportable.length} records</span>
          </div>
          {references.deliverables.map(({ code, label }) => {
            const related = reportable.filter(
              (record) => record.orbit.primaryDeliverable === code,
            );
            const mins = related.reduce(
              (sum, record) => sum + record.orbit.stemPocMinutes,
              0,
            );
            const max = Math.max(
              1,
              ...references.deliverables.map(({ code: c }) =>
                reportable
                  .filter((r) => r.orbit.primaryDeliverable === c)
                  .reduce((s, r) => s + r.orbit.stemPocMinutes, 0),
              ),
            );
            return (
              <div className="deliverable-row" key={code}>
                <b>{code}</b>
                <span>
                  <strong>{label}</strong>
                  <i>
                    <i
                      style={{ width: `${Math.max(3, (mins / max) * 100)}%` }}
                    />
                  </i>
                </span>
                <em>{hours(mins)}</em>
              </div>
            );
          })}
        </section>
        <aside className="panel rule-panel">
          <p className="eyebrow">Calculation rule</p>
          <h2>Minutes first. Days derived.</h2>
          <div className="formula">
            <b>{poc}</b>
            <span>minutes</span>
            <i>÷</i>
            <b>{references.reportingConfig.minutesPerReportingDay}</b>
            <span>minutes/day</span>
            <i>=</i>
            <b>{deriveReportingDays(poc, references.reportingConfig).toFixed(2)}</b>
            <span>days</span>
          </div>
          <p>
            The source record keeps precise time. Reporting settings can change
            later without destroying the underlying history.
          </p>
          <div className="rule-note">
            Quarter is derived automatically from the activity date and
            configurable school-year boundaries.
          </div>
        </aside>
      </div>
    </div>
  );
}

function LogWizard({
  record,
  step,
  setStep,
  patch,
  toggle,
  close,
  save,
  saving,
  error,
  references,
  returnFocusRef,
}: {
  record: WorkRecord;
  step: number;
  setStep: (s: number) => void;
  patch: (p: Partial<WorkRecord>) => void;
  toggle: (
    k: "projectIds" | "organizationIds" | "contactIds" | "categoryIds",
    id: string,
  ) => void;
  close: () => void;
  save: (another?: boolean) => void;
  saving: boolean;
  error: string;
  references: ReferenceData;
  returnFocusRef: React.RefObject<HTMLElement | null>;
}) {
  const titles = [
    "What did you do?",
    "Who was it for or with?",
    "How long?",
    "What resulted?",
    "Does it connect to STEM / ORBIT?",
  ];
  const dialogRef = useRef<HTMLElement>(null);
  const firstFieldRef = useRef<HTMLInputElement>(null);
  useEffect(() => {
    const returnFocus = returnFocusRef.current;
    firstFieldRef.current?.focus();
    return () => returnFocus?.focus();
  }, [returnFocusRef]);
  useEffect(() => {
    const handleDialogKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        close();
        return;
      }
      if (event.key !== "Tab" || !dialogRef.current) return;
      const focusable = [...dialogRef.current.querySelectorAll<HTMLElement>('button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), summary, [tabindex]:not([tabindex="-1"])')].filter((element) => {
        const details = element.closest("details");
        return !element.hidden && (!details || details.open || element.tagName === "SUMMARY");
      });
      if (!focusable.length) return;
      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
    };
    document.addEventListener("keydown", handleDialogKeyDown);
    return () => document.removeEventListener("keydown", handleDialogKeyDown);
  }, [close]);
  const canNext = step !== 1 || Boolean(record.title.trim() && record.activityType);
  return (
    <div className="modal-backdrop" role="presentation">
      <section
        ref={dialogRef}
        className="log-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="log-title"
        aria-describedby="log-step-status"
        tabIndex={-1}
      >
        <header className="log-header">
          <div>
            <p className="eyebrow">
              {record.isSample
                ? "Edit development sample"
                : "Universal work record"}
            </p>
            <h2 id="log-title">{titles[step - 1]}</h2>
          </div>
          <button onClick={close} aria-label="Close work entry">
            ×
          </button>
        </header>
        <p className="sr-only" id="log-step-status">Step {step} of 5: {titles[step - 1]}. Current step.</p>
        <div className="stepper" aria-label="Work entry steps">
          {titles.map((title, index) => (
            <button
              key={index}
              onClick={() => setStep(index + 1)}
              className={
                step === index + 1 ? "active" : step > index + 1 ? "done" : ""
              }
              aria-current={step === index + 1 ? "step" : undefined}
              aria-label={`Step ${index + 1} of 5: ${title}${step === index + 1 ? ", current step" : step > index + 1 ? ", completed" : ""}`}
            >
              <span>{step > index + 1 ? "✓" : index + 1}</span>
              <i />
            </button>
          ))}
        </div>
        <div className="log-content">
          {step === 1 && (
            <div className="form-stack">
              <label>
                <span>
                  Activity title <b>*</b>
                </span>
                <input
                  ref={firstFieldRef}
                  value={record.title}
                  onChange={(e) => patch({ title: e.target.value })}
                  placeholder="e.g. District STEELS planning meeting"
                />
              </label>
              <label>
                <span>
                  Activity type <b>*</b>
                </span>
                <select
                  value={record.activityType}
                  onChange={(e) => patch({ activityType: e.target.value })}
                >
                  <option value="">Choose an activity type…</option>
                  {references.settings.activityTypes.map((type) => (
                    <option key={type}>{type}</option>
                  ))}
                </select>
              </label>
              <label>
                <span>Short description</span>
                <textarea
                  value={record.description}
                  onChange={(e) => patch({ description: e.target.value })}
                  placeholder="A sentence is enough."
                  rows={3}
                />
              </label>
            </div>
          )}
          {step === 2 && (
            <div className="form-stack">
              <fieldset>
                <legend>District / LEA scope</legend>
                <div className="scope-grid">
                  {[
                    ["none", "No district"],
                    ["specific", "Specific district(s)"],
                    ["regional", "Regional work"],
                    ["allDistricts", "All districts"],
                  ].map(([value, label]) => (
                    <button
                      type="button"
                      key={value}
                      className={record.engagementScope === value ? "selected" : ""}
                      aria-pressed={record.engagementScope === value}
                      onClick={() => {
                        const districtIds = new Set(references.organizations.filter((item) => item.type === "district").map((item) => item.appId));
                        patch({
                          engagementScope: value as WorkRecord["engagementScope"],
                          organizationIds: value === "specific" ? record.organizationIds : record.organizationIds.filter((id) => !districtIds.has(id)),
                        });
                      }}
                    >
                      {label}
                    </button>
                  ))}
                </div>
              </fieldset>
              {record.engagementScope === "specific" && (
                <fieldset>
                  <legend>District(s) <small>select one or more</small></legend>
                  <div className="choice-grid">
                    {references.organizations.filter((item) => item.type === "district").map((item) => (
                      <button type="button" key={item.appId} className={record.organizationIds.includes(item.appId) ? "selected" : ""} onClick={() => toggle("organizationIds", item.appId)}>
                        <span>▦</span>{item.name}<i>{record.organizationIds.includes(item.appId) ? "✓" : ""}</i>
                      </button>
                    ))}
                  </div>
                </fieldset>
              )}
              <fieldset>
                <legend>Organization / partner <small>optional</small></legend>
                <div className="choice-grid">
                  {references.organizations.filter((item) => item.type !== "district").map((item) => (
                    <button type="button" key={item.appId} className={record.organizationIds.includes(item.appId) ? "selected" : ""} onClick={() => toggle("organizationIds", item.appId)}>
                      <span>{item.type === "partner" ? "◇" : "○"}</span>{item.name}<i>{record.organizationIds.includes(item.appId) ? "✓" : ""}</i>
                    </button>
                  ))}
                </div>
              </fieldset>
              <fieldset>
                <legend>
                  Project <small>optional</small>
                </legend>
                <div className="chip-grid">
                  {references.projects
                    // A completed project remains a valid, historical connection (see
                    // docs/AI_HANDOFF.md "Durable Projects (Patch 7)" — Completed-project
                    // behavior), but is not offered as a normal default choice for new work.
                    // An already-connected complete project on the record being edited stays
                    // visible so an existing relationship is never hidden.
                    .filter((item) => item.status !== "complete" || record.projectIds.includes(item.appId))
                    .map((item) => (
                      <button
                        type="button"
                        key={item.appId}
                        className={
                          record.projectIds.includes(item.appId) ? "selected" : ""
                        }
                        onClick={() => toggle("projectIds", item.appId)}
                      >
                        {item.name}
                      </button>
                    ))}
                </div>
              </fieldset>
              <details className="advanced-details">
                <summary>Add classification / reach</summary>
                <div className="advanced-content">
                  <fieldset>
                    <legend>Work areas <small>select all that apply</small></legend>
                    <div className="chip-grid">
                      {references.categories.map((item) => (
                        <button type="button" key={item.appId} className={record.categoryIds.includes(item.appId) ? "selected" : ""} onClick={() => toggle("categoryIds", item.appId)}>{item.name}</button>
                      ))}
                    </div>
                  </fieldset>
                  <fieldset>
                    <legend>Audience reach <small>optional</small></legend>
                    <div className="reach-grid">
                      {[["Educators / leaders", "educatorsLeaders"], ["Students / families", "studentsFamilies"], ["Workforce / community", "workforceCommunity"], ["Other", "other"]].map(([label, key]) => (
                        <label key={key}><span>{label}</span><input type="number" min="0" value={record.reach[key as keyof typeof record.reach]} onChange={(e) => patch({ reach: { ...record.reach, [key]: Number(e.target.value) } })} /></label>
                      ))}
                    </div>
                  </fieldset>
                  <fieldset>
                    <legend>Contacts <small>optional</small></legend>
                    <div className="chip-grid">
                      {references.contacts.map((item) => <button type="button" key={item.appId} className={record.contactIds.includes(item.appId) ? "selected" : ""} onClick={() => toggle("contactIds", item.appId)}>{item.displayName}</button>)}
                    </div>
                  </fieldset>
                  <label><span>Evidence summary</span><textarea rows={2} value={record.evidenceSummary} onChange={(e) => patch({ evidenceSummary: e.target.value })} placeholder="Optional note about supporting evidence" /></label>
                  <label><span>Stable evidence reference IDs <small>comma-separated</small></span><input value={record.evidenceReferenceIds.join(", ")} onChange={(e) => patch({ evidenceReferenceIds: e.target.value.split(",").map((value) => value.trim()).filter(Boolean) })} placeholder="e.g. dev-evidence-roadmap" /></label>
                </div>
              </details>
            </div>
          )}
          {step === 3 && (
            <div className="form-stack time-step">
              <label>
                <span>
                  Activity date <b>*</b>
                </span>
                <input
                  type="date"
                  value={record.activityDate}
                  onChange={(e) => patch({ activityDate: e.target.value })}
                />
              </label>
              <fieldset>
                <legend>
                  Duration <b>*</b>
                </legend>
                <div className="duration-options">
                  {[30, 60, 90, 120, 210, 420].map((mins) => (
                    <button
                      key={mins}
                      type="button"
                      className={
                        record.durationMinutes === mins ? "selected" : ""
                      }
                      onClick={() => patch({ durationMinutes: mins })}
                    >
                      {mins < 60
                        ? `${mins} min`
                        : mins === 210
                          ? "3.5 hr"
                          : `${mins / 60} hr`}
                    </button>
                  ))}
                </div>
                <label className="custom-duration">
                  <span>Custom minutes</span>
                  <input
                    type="number"
                    min="1"
                    step="5"
                    value={record.durationMinutes}
                    onChange={(e) =>
                      patch({
                        durationMinutes: Math.max(1, Number(e.target.value)),
                      })
                    }
                  />
                </label>
              </fieldset>
              <div className="time-summary">
                <span>Recorded precisely</span>
                <strong>{hours(record.durationMinutes)}</strong>
                <small>
                  Reporting days are calculated later; source time stays intact.
                </small>
              </div>
            </div>
          )}
          {step === 4 && (
            <div className="form-stack">
              <label>
                <span>
                  Output <small>What was produced?</small>
                </span>
                <textarea
                  rows={2}
                  value={record.output}
                  onChange={(e) => patch({ output: e.target.value })}
                  placeholder="Resource, plan, event, connection…"
                />
              </label>
              <label>
                <span>
                  Outcome <small>What changed?</small>
                </span>
                <textarea
                  rows={2}
                  value={record.outcome}
                  onChange={(e) => patch({ outcome: e.target.value })}
                  placeholder="Progress, learning, decision, next action…"
                />
              </label>
              <label>
                <span>Next step</span>
                <input
                  value={record.nextStep}
                  onChange={(e) => patch({ nextStep: e.target.value })}
                  placeholder="Optional follow-up action"
                />
              </label>
              <div className="toggle-line">
                <input
                  aria-label="Follow-up needed"
                  type="checkbox"
                  checked={record.followUpNeeded}
                  onChange={(e) =>
                    patch({
                      followUpNeeded: e.target.checked,
                      followUpDate: e.target.checked
                        ? (record.followUpDate ?? record.activityDate)
                        : null,
                    })
                  }
                />
                <span>
                  <strong>Follow-up needed</strong>
                  <small>Keep this action visible in Today.</small>
                </span>
              </div>
              {record.followUpNeeded && (
                <label>
                  <span>Follow-up date</span>
                  <input
                    type="date"
                    value={record.followUpDate ?? ""}
                    onChange={(e) =>
                      patch({ followUpDate: e.target.value || null })
                    }
                  />
                </label>
              )}
            </div>
          )}
          {step === 5 && (
            <div className="form-stack">
              <div className="orbit-question">
                <div>
                  <span className="orbit-symbol">◎</span>
                  <span>
                    <strong>Is this activity ORBIT reportable?</strong>
                    <small>Optional. Non-STEM work belongs here too.</small>
                  </span>
                </div>
                <label className="switch">
                  <input
                    aria-label="ORBIT reportable"
                    type="checkbox"
                    checked={record.orbit.reportable}
                    onChange={(e) =>
                      patch({
                        orbit: {
                          ...record.orbit,
                          reportable: e.target.checked,
                          primaryDeliverable: e.target.checked ? record.orbit.primaryDeliverable : null,
                          supportingDeliverables: e.target.checked ? record.orbit.supportingDeliverables : [],
                          stemPocMinutes:
                            e.target.checked &&
                            record.orbit.stemPocMinutes === 0
                              ? record.durationMinutes
                              : e.target.checked ? record.orbit.stemPocMinutes : 0,
                          tacMinutes: e.target.checked ? record.orbit.tacMinutes : 0,
                        },
                      })
                    }
                  />
                  <span aria-hidden="true" />
                </label>
              </div>
              {record.orbit.reportable ? (
                <>
                  <fieldset>
                    <legend>Primary deliverable</legend>
                    <div className="deliverable-choices">
                      {references.deliverables.map(({ code, label }) => (
                        <button
                          type="button"
                          key={code}
                          aria-label={`Primary deliverable ${code}: ${label}`}
                          className={
                            record.orbit.primaryDeliverable === code
                              ? "selected"
                              : ""
                          }
                          onClick={() =>
                            patch({
                              orbit: {
                                ...record.orbit,
                                primaryDeliverable: code,
                                supportingDeliverables: record.orbit.supportingDeliverables.filter((value) => value !== code),
                              },
                            })
                          }
                        >
                          <b>{code}</b>
                          <span>{label}</span>
                        </button>
                      ))}
                    </div>
                  </fieldset>
                  <details className="advanced-details">
                    <summary>Add supporting deliverables</summary>
                    <div className="advanced-content">
                      <div className="chip-grid">
                        {references.deliverables.filter(({ code }) => code !== record.orbit.primaryDeliverable).map(({ code, label }) => (
                          <button
                            type="button"
                            key={code}
                            aria-label={`Supporting deliverable ${code}: ${label}`}
                            className={record.orbit.supportingDeliverables.includes(code) ? "selected" : ""}
                            onClick={() => patch({ orbit: { ...record.orbit, supportingDeliverables: record.orbit.supportingDeliverables.includes(code) ? record.orbit.supportingDeliverables.filter((value) => value !== code) : [...record.orbit.supportingDeliverables, code] } })}
                          >
                            {code} · {label}
                          </button>
                        ))}
                      </div>
                    </div>
                  </details>
                  <div className="form-two">
                    <label>
                      <span>STEM PoC minutes</span>
                      <input
                        type="number"
                        min="0"
                        value={record.orbit.stemPocMinutes}
                        onChange={(e) =>
                          patch({
                            orbit: {
                              ...record.orbit,
                              stemPocMinutes: Number(e.target.value),
                            },
                          })
                        }
                      />
                    </label>
                    <label>
                      <span>TaC minutes</span>
                      <input
                        type="number"
                        min="0"
                        value={record.orbit.tacMinutes}
                        onChange={(e) =>
                          patch({
                            orbit: {
                              ...record.orbit,
                              tacMinutes: Number(e.target.value),
                            },
                          })
                        }
                      />
                    </label>
                  </div>
                  <label>
                    <span>Qualitative evidence</span>
                    <textarea
                      rows={2}
                      value={record.orbit.evidence}
                      onChange={(e) =>
                        patch({
                          orbit: { ...record.orbit, evidence: e.target.value },
                        })
                      }
                      placeholder="Notes, links, resources, success story…"
                    />
                  </label>
                </>
              ) : (
                <div className="non-orbit-note">
                  <span>✓</span>
                  <div>
                    <strong>This is still valuable IU work.</strong>
                    <p>
                      Save it normally. ORBIT does not define the rest of your
                      work.
                    </p>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
        {error && (
          <div className="form-error" role="alert">
            {error}
          </div>
        )}
        <footer className="log-footer">
          <button
            className="ghost-button"
            onClick={step === 1 ? close : () => setStep(step - 1)}
          >
            {step === 1 ? "Cancel" : "← Back"}
          </button>
          <div>
            {step < 5 ? (
              <button
                className="primary-action"
                disabled={!canNext}
                onClick={() => setStep(step + 1)}
              >
                Continue <span>→</span>
              </button>
            ) : (
              <>
                <button
                  className="secondary-button"
                  disabled={saving}
                  onClick={() => save(true)}
                >
                  Save & log another
                </button>
                <button
                  className="primary-action"
                  disabled={saving}
                  onClick={() => save(false)}
                >
                  {saving ? "Saving…" : "Save & done"}
                </button>
              </>
            )}
          </div>
        </footer>
      </section>
    </div>
  );
}

function Metric({ value, label }: { value: string; label: string }) {
  return (
    <div className="metric">
      <strong>{value}</strong>
      <span>{label}</span>
    </div>
  );
}
function MetricCard({
  value,
  label,
  detail,
  tone,
}: {
  value: string;
  label: string;
  detail?: string;
  tone: string;
}) {
  return (
    <article className={`metric-card ${tone}`}>
      <span className="metric-orb" />
      <strong>{value}</strong>
      <h3>{label}</h3>
      {detail && <p>{detail}</p>}
    </article>
  );
}
function RecordRow({
  record,
  onClick,
  organizations,
}: {
  record: WorkRecord;
  onClick: () => void;
  organizations: ReferenceData["organizations"];
}) {
  return (
    <button className="record-row" onClick={onClick}>
      <span
        className={`record-dot ${record.orbit.reportable ? "orbit" : "iu"}`}
      />
      <span>
        <strong>{record.title}</strong>
        <small>
          {record.isSample && <em>Sample</em>}
          {record.organizationIds[0]
            ? entityName(organizations, record.organizationIds[0])
            : record.engagementScope === "regional"
              ? "Regional scope"
              : record.engagementScope === "allDistricts"
                ? "All districts"
                : record.activityType}
        </small>
      </span>
      <b>{hours(record.durationMinutes)}</b>
      <i>→</i>
    </button>
  );
}
/**
 * The Home attention surface for already-durable Inbox Intelligence — a compact view, not a
 * second management screen. Every row/count here is derived from records the parent already
 * loaded through the existing provider; this component makes zero network/AI calls itself.
 */
function ActionCenter({
  summary,
  needsAttentionCount,
  needsAttentionItems,
  waitingItems,
  loadFailed,
  setView,
}: {
  summary: { openCount: number; waitingCount: number; resolvedCount: number };
  needsAttentionCount: number;
  needsAttentionItems: InboxIntelligenceRecord[];
  waitingItems: InboxIntelligenceRecord[];
  loadFailed: boolean;
  setView: (v: View) => void;
}) {
  const openInbox = () => setView("inbox");
  const isEmpty = summary.openCount + summary.waitingCount + summary.resolvedCount === 0;

  return (
    <section className="panel list-panel action-center">
      <div className="panel-heading">
        <div>
          <p className="eyebrow">What deserves your attention</p>
          <h2>Action Center</h2>
        </div>
        {!isEmpty && !loadFailed && (
          <button className="text-action" onClick={openInbox}>
            View Intelligence →
          </button>
        )}
      </div>
      {loadFailed ? (
        <Empty
          title="Inbox Intelligence is temporarily unavailable"
          copy="Your Work Record data above is unaffected. Try Inbox Intelligence directly."
          action="Open Inbox Intelligence"
          onAction={openInbox}
        />
      ) : isEmpty ? (
        <Empty
          title="Nothing needs your attention"
          copy="Analyze an email in Inbox Intelligence to begin building your work intelligence."
          action="Open Inbox Intelligence"
          onAction={openInbox}
        />
      ) : (
        <>
          {needsAttentionItems.length > 0 && (
            <div className="action-center-section">
              <div className="panel-heading">
                <h3>Needs attention</h3>
                <span className="count-badge">{needsAttentionCount}</span>
              </div>
              {needsAttentionItems.map((record) => (
                <ActionCenterRow key={record.appId} record={record} showDue onClick={openInbox} />
              ))}
            </div>
          )}
          {waitingItems.length > 0 && (
            <div className="action-center-section">
              <div className="panel-heading">
                <h3>Waiting on</h3>
                <span className="count-badge">{summary.waitingCount}</span>
              </div>
              {waitingItems.map((record) => (
                <ActionCenterRow key={record.appId} record={record} showDue={false} onClick={openInbox} />
              ))}
            </div>
          )}
          {needsAttentionItems.length === 0 && waitingItems.length === 0 && (
            <p className="muted-copy">Nothing needs attention or is waiting right now.</p>
          )}
          <div className="metric-strip">
            <Metric value={String(summary.openCount)} label="open" />
            <Metric value={String(summary.waitingCount)} label="waiting" />
            <Metric value={String(summary.resolvedCount)} label="resolved" />
          </div>
        </>
      )}
    </section>
  );
}
function ActionCenterRow({
  record,
  showDue,
  onClick,
}: {
  record: InboxIntelligenceRecord;
  showDue: boolean;
  onClick: () => void;
}) {
  const dueLabel = showDue ? formatDueLabel(earliestDueDate(record)) : null;
  return (
    <button className="record-row" onClick={onClick}>
      <span className={`record-dot ${showDue ? "iu" : "orbit"}`} />
      <span>
        <strong>{primaryLabel(record)}</strong>
        {dueLabel?.overdue && (
          <small>
            <em>Overdue</em>
          </small>
        )}
      </span>
      <b>{dueLabel && !dueLabel.overdue ? dueLabel.text : ""}</b>
      <i>→</i>
    </button>
  );
}
function Empty({
  title,
  copy,
  action,
  onAction,
}: {
  title: string;
  copy: string;
  action?: string;
  onAction?: () => void;
}) {
  return (
    <div className="empty">
      <span>○</span>
      <strong>{title}</strong>
      <p>{copy}</p>
      {action && <button onClick={onAction}>{action}</button>}
    </div>
  );
}
function Command({
  icon,
  title,
  copy,
  onClick,
  accent = false,
}: {
  icon: string;
  title: string;
  copy: string;
  onClick: () => void;
  accent?: boolean;
}) {
  return (
    <button
      className={`command-card ${accent ? "accent" : ""}`}
      onClick={onClick}
    >
      <span>{icon}</span>
      <span>
        <strong>{title}</strong>
        <small>{copy}</small>
      </span>
      <b>→</b>
    </button>
  );
}
function PageHeading({
  eyebrow,
  title,
  copy,
  action,
  onAction,
}: {
  eyebrow: string;
  title: string;
  copy: string;
  action?: string;
  onAction?: () => void;
}) {
  return (
    <div className="page-heading">
      <div>
        <p className="eyebrow">{eyebrow}</p>
        <h1>{title}</h1>
        <p>{copy}</p>
      </div>
      {action && (
        <button className="primary-action" onClick={onAction}>
          {action}
        </button>
      )}
    </div>
  );
}
