"use client";

import { useEffect, useRef, useState } from "react";
import {
  ApiDataProvider,
  PrototypeFallbackProvider,
  type DataProvider,
} from "../lib/data-provider";
import {
  CATEGORIES,
  DELIVERABLES,
  MINUTES_PER_REPORTING_DAY,
  ORGANIZATIONS,
  PROJECTS,
  type WorkRecord,
} from "../lib/models";
import { SAMPLE_RECORDS } from "../lib/sample-data";

type View = "home" | "today" | "history" | "projects" | "orbit";
const activityTypes = [
  "District meeting",
  "Professional learning",
  "Classroom support",
  "Project planning",
  "Partner meeting",
  "Student program",
  "Internal planning",
  "Resource development",
  "Follow-up communication",
  "Other",
];
const navItems: [View, string, string][] = [
  ["home", "⌂", "Home"],
  ["today", "◷", "Today"],
  ["history", "≡", "History"],
  ["projects", "▤", "Projects"],
  ["orbit", "◎", "STEM / ORBIT"],
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
  const now = new Date().toISOString();
  return {
    appId: crypto.randomUUID(),
    title: "",
    activityDate: todayIso(),
    activityType: "District meeting",
    description: "",
    detailedNotes: "",
    durationMinutes: 60,
    status: "complete",
    projectIds: [],
    organizationIds: [],
    categoryIds: [],
    reach: {
      educatorsLeaders: 0,
      studentsFamilies: 0,
      workforceCommunity: 0,
      other: 0,
    },
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
    isSample: false,
    createdAt: now,
    modifiedAt: now,
    syncState: "saved",
  };
}

export default function IUWorkTracker() {
  const [view, setView] = useState<View>("home");
  const [records, setRecords] = useState<WorkRecord[]>(SAMPLE_RECORDS);
  const [loading, setLoading] = useState(true);
  const [storageMode, setStorageMode] = useState<"connected" | "fallback">(
    "connected",
  );
  const [logging, setLogging] = useState(false);
  const [draft, setDraft] = useState<WorkRecord>(emptyRecord);
  const [step, setStep] = useState(1);
  const [saveError, setSaveError] = useState("");
  const [saving, setSaving] = useState(false);
  const [search, setSearch] = useState("");
  const provider = useRef<DataProvider>(new ApiDataProvider());
  useEffect(() => {
    let live = true;
    provider.current
      .listWorkRecords()
      .then((data) => {
        if (live) setRecords(data);
      })
      .catch(async () => {
        const fallback = new PrototypeFallbackProvider();
        provider.current = fallback;
        if (live) {
          setRecords(await fallback.listWorkRecords());
          setStorageMode("fallback");
        }
      })
      .finally(() => {
        if (live) setLoading(false);
      });
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
  const openLog = (record?: WorkRecord) => {
    setDraft(record ? structuredClone(record) : emptyRecord());
    setStep(1);
    setSaveError("");
    setLogging(true);
  };
  const patchDraft = (patch: Partial<WorkRecord>) =>
    setDraft((current) => ({ ...current, ...patch }));
  const toggle = (
    key: "projectIds" | "organizationIds" | "categoryIds",
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
    setSaving(true);
    setSaveError("");
    const pending = {
      ...draft,
      title: draft.title.trim(),
      modifiedAt: new Date().toISOString(),
      syncState: "saving" as const,
    };
    try {
      const saved = await provider.current.saveWorkRecord(pending);
      setRecords((current) =>
        [saved, ...current.filter((item) => item.appId !== saved.appId)].sort(
          (a, b) => b.activityDate.localeCompare(a.activityDate),
        ),
      );
      if (another) {
        setDraft(emptyRecord());
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
  return (
    <main className="os-shell">
      <Header view={view} setView={setView} onLog={() => openLog()} />
      <div className="os-body">
        <SideNav view={view} setView={setView} onLog={() => openLog()} />
        <section className="screen" aria-live="polite">
          {storageMode === "fallback" && (
            <div className="storage-banner">
              <span>Prototype mode</span> Changes work in this preview session;
              the connected data store will be used when available.
            </div>
          )}
          {loading ? (
            <div className="loading-card">Preparing your workspace…</div>
          ) : view === "home" ? (
            <Home
              records={records}
              todayRecords={todayRecords}
              totalToday={totalToday}
              followups={followups}
              openLog={openLog}
              setView={setView}
            />
          ) : view === "today" ? (
            <Today
              records={todayRecords}
              total={totalToday}
              followups={followups}
              openLog={openLog}
            />
          ) : view === "history" ? (
            <History
              records={records}
              search={search}
              setSearch={setSearch}
              openLog={openLog}
            />
          ) : view === "projects" ? (
            <Projects records={records} setView={setView} />
          ) : (
            <Orbit records={records} />
          )}
        </section>
      </div>
      <footer className="os-status">
        <span>
          <i />{" "}
          {storageMode === "connected"
            ? "Prototype data store connected"
            : "Preview session active"}
        </span>
        <span>Log it once. Use it everywhere.</span>
      </footer>
      {logging && (
        <LogWizard
          record={draft}
          step={step}
          setStep={setStep}
          patch={patchDraft}
          toggle={toggle}
          close={() => setLogging(false)}
          save={save}
          saving={saving}
          error={saveError}
        />
      )}
    </main>
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
        <button className="avatar" aria-label="Account options">
          GM
        </button>
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
}: {
  records: WorkRecord[];
  todayRecords: WorkRecord[];
  totalToday: number;
  followups: WorkRecord[];
  openLog: (r?: WorkRecord) => void;
  setView: (v: View) => void;
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
      </div>
    </div>
  );
}
function Today({
  records,
  total,
  followups,
  openLog,
}: {
  records: WorkRecord[];
  total: number;
  followups: WorkRecord[];
  openLog: (r?: WorkRecord) => void;
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
}: {
  records: WorkRecord[];
  search: string;
  setSearch: (v: string) => void;
  openLog: (r: WorkRecord) => void;
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
                  ? entityName(PROJECTS, record.projectIds[0])
                  : "No project"}
              </strong>
              <small>
                {record.organizationIds[0]
                  ? entityName(ORGANIZATIONS, record.organizationIds[0])
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
function Projects({
  records,
  setView,
}: {
  records: WorkRecord[];
  setView: (v: View) => void;
}) {
  return (
    <div className="screen-inner">
      <PageHeading
        eyebrow="Connected containers"
        title="Projects"
        copy="See the activity, time, and outcomes accumulating around each initiative."
      />
      <div className="project-grid">
        {PROJECTS.map((project) => {
          const related = records.filter((record) =>
            record.projectIds.includes(project.appId),
          );
          const mins = related.reduce(
            (sum, record) => sum + record.durationMinutes,
            0,
          );
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
              <button onClick={() => setView("history")}>
                View connected work →
              </button>
            </article>
          );
        })}
      </div>
      <div className="planned-note">
        <span>＋</span>
        <div>
          <strong>
            Project creation is planned for the next capability slice.
          </strong>
          <p>
            V1 proves the reusable project relationship and calculates totals
            from universal Work Records.
          </p>
        </div>
      </div>
    </div>
  );
}
function Orbit({ records }: { records: WorkRecord[] }) {
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
          value={(poc / MINUTES_PER_REPORTING_DAY).toFixed(2)}
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
          {DELIVERABLES.map(([code, label]) => {
            const related = reportable.filter(
              (record) => record.orbit.primaryDeliverable === code,
            );
            const mins = related.reduce(
              (sum, record) => sum + record.orbit.stemPocMinutes,
              0,
            );
            const max = Math.max(
              1,
              ...DELIVERABLES.map(([c]) =>
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
            <b>{MINUTES_PER_REPORTING_DAY}</b>
            <span>minutes/day</span>
            <i>=</i>
            <b>{(poc / MINUTES_PER_REPORTING_DAY).toFixed(2)}</b>
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
}: {
  record: WorkRecord;
  step: number;
  setStep: (s: number) => void;
  patch: (p: Partial<WorkRecord>) => void;
  toggle: (
    k: "projectIds" | "organizationIds" | "categoryIds",
    id: string,
  ) => void;
  close: () => void;
  save: (another?: boolean) => void;
  saving: boolean;
  error: string;
}) {
  const titles = [
    "What did you do?",
    "Who was it for or with?",
    "How long?",
    "What resulted?",
    "Does it connect to STEM / ORBIT?",
  ];
  const canNext = step !== 1 || Boolean(record.title.trim());
  return (
    <div className="modal-backdrop" role="presentation">
      <section
        className="log-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="log-title"
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
        <div className="stepper" aria-label={`Step ${step} of 5`}>
          {titles.map((_, index) => (
            <button
              key={index}
              onClick={() => setStep(index + 1)}
              className={
                step === index + 1 ? "active" : step > index + 1 ? "done" : ""
              }
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
                  {activityTypes.map((type) => (
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
                <legend>Organization or district</legend>
                <div className="choice-grid">
                  {ORGANIZATIONS.map((item) => (
                    <button
                      type="button"
                      key={item.appId}
                      className={
                        record.organizationIds.includes(item.appId)
                          ? "selected"
                          : ""
                      }
                      onClick={() => toggle("organizationIds", item.appId)}
                    >
                      <span>
                        {item.type === "district"
                          ? "▦"
                          : item.type === "partner"
                            ? "◇"
                            : "○"}
                      </span>
                      {item.name}
                      <i>
                        {record.organizationIds.includes(item.appId) ? "✓" : ""}
                      </i>
                    </button>
                  ))}
                </div>
              </fieldset>
              <fieldset>
                <legend>
                  Project <small>optional</small>
                </legend>
                <div className="chip-grid">
                  {PROJECTS.map((item) => (
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
              <fieldset>
                <legend>
                  Work areas <small>select all that apply</small>
                </legend>
                <div className="chip-grid">
                  {CATEGORIES.map((item) => (
                    <button
                      type="button"
                      key={item.appId}
                      className={
                        record.categoryIds.includes(item.appId)
                          ? "selected"
                          : ""
                      }
                      onClick={() => toggle("categoryIds", item.appId)}
                    >
                      {item.name}
                    </button>
                  ))}
                </div>
              </fieldset>
              <fieldset>
                <legend>
                  Audience reach <small>optional</small>
                </legend>
                <div className="reach-grid">
                  {[
                    ["Educators / leaders", "educatorsLeaders"],
                    ["Students / families", "studentsFamilies"],
                    ["Workforce / community", "workforceCommunity"],
                    ["Other", "other"],
                  ].map(([label, key]) => (
                    <label key={key}>
                      <span>{label}</span>
                      <input
                        type="number"
                        min="0"
                        value={record.reach[key as keyof typeof record.reach]}
                        onChange={(e) =>
                          patch({
                            reach: {
                              ...record.reach,
                              [key]: Math.max(0, Number(e.target.value)),
                            },
                          })
                        }
                      />
                    </label>
                  ))}
                </div>
              </fieldset>
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
                <div className="switch">
                  <input
                    aria-label="ORBIT reportable"
                    type="checkbox"
                    checked={record.orbit.reportable}
                    onChange={(e) =>
                      patch({
                        orbit: {
                          ...record.orbit,
                          reportable: e.target.checked,
                          stemPocMinutes:
                            e.target.checked &&
                            record.orbit.stemPocMinutes === 0
                              ? record.durationMinutes
                              : record.orbit.stemPocMinutes,
                        },
                      })
                    }
                  />
                  <span />
                </div>
              </div>
              {record.orbit.reportable ? (
                <>
                  <fieldset>
                    <legend>Primary deliverable</legend>
                    <div className="deliverable-choices">
                      {DELIVERABLES.map(([code, label]) => (
                        <button
                          type="button"
                          key={code}
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
}: {
  record: WorkRecord;
  onClick: () => void;
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
            ? entityName(ORGANIZATIONS, record.organizationIds[0])
            : record.activityType}
        </small>
      </span>
      <b>{hours(record.durationMinutes)}</b>
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
