"use client";

// Patch 7.3 — the "Process selected" review panel for Voice Intelligence. AI EXTRACTS. HUMAN
// SELECTS. THE SYSTEM ROUTES. This component is purely a view: it groups the SELECTED candidates
// by type, shows each type's destination, and hands every action back to VoiceIntelligence,
// which reuses the existing Work Record wizard, ContactFormModal/ContactMatchPanel, and the
// (extracted) OrganizationFormModal/ProjectFormModal — no second form, no second provider, no
// silent routing, no AI, no network. Opening this panel writes nothing.

import { useState } from "react";
import { matchContactCandidates } from "../lib/contact-matching";
import type { Organization, Project, ReferenceData } from "../lib/models";
import {
  formatCandidateForCopy,
  groupSelectedCandidates,
  matchOrganizationCandidates,
  matchProjectCandidates,
  VOICE_DESTINATIONS,
  type VoiceCandidateGroup,
} from "../lib/voice-routing";
import type { VoiceReviewCandidate } from "../lib/voice-intelligence-session";
import ContactMatchPanel from "./ContactMatchPanel";

type RoutingPatch = Partial<Pick<VoiceReviewCandidate, "contactDecision" | "contactCreated" | "routing">>;

export default function VoiceRoutingPanel({
  candidates,
  references,
  skippedIds,
  onSkip,
  onUnskip,
  onLogAsWork,
  onPatch,
  onAddPerson,
  onCreateEntity,
  onClose,
}: {
  candidates: VoiceReviewCandidate[];
  references: ReferenceData;
  /** Completed-work candidates the human skipped in THIS visit (transient — nothing is lost). */
  skippedIds: Set<string>;
  onSkip: (id: string) => void;
  onUnskip: () => void;
  onLogAsWork: (candidate: VoiceReviewCandidate) => void;
  onPatch: (id: string, patch: RoutingPatch) => void;
  onAddPerson: (id: string) => void;
  onCreateEntity: (id: string) => void;
  onClose: () => void;
}) {
  const { ready, unsupported } = groupSelectedCandidates(candidates);
  const selectedCount = ready.concat(unsupported).reduce((sum, group) => sum + group.candidates.length, 0);
  const allGroups = ready.concat(unsupported);

  return (
    <section className="panel voice-routing" aria-label="Process selected">
      <div className="panel-heading">
        <h2>Process selected</h2>
        <button type="button" className="ghost-button" onClick={onClose}>
          Done
        </button>
      </div>
      <p className="muted-copy">
        {selectedCount} selected. Nothing is written until you review and explicitly save each item — opening this panel changes nothing.
      </p>
      <div className="candidate-summary" aria-label="Selected by type">
        {allGroups.map((group) => (
          <span key={group.type}>
            {group.label} — {group.candidates.length}
          </span>
        ))}
      </div>

      {ready.length > 0 && (
        <>
          <p className="eyebrow">Ready destinations</p>
          {ready.map((group) => (
            <section key={group.type} className="voice-routing-group" aria-label={`${group.label} destination`}>
              <h3>
                {group.label} — {group.candidates.length}
                <small> → {VOICE_DESTINATIONS[group.type].label}</small>
              </h3>
              {group.type === "COMPLETED_WORK" && (
                <CompletedWorkQueue group={group} skippedIds={skippedIds} onSkip={onSkip} onUnskip={onUnskip} onLogAsWork={onLogAsWork} />
              )}
              {group.type === "PERSON" &&
                group.candidates.map((candidate) => (
                  <ContactMatchPanel
                    key={candidate.id}
                    personName={candidate.title}
                    candidates={matchContactCandidates(candidate.title, references.contacts)}
                    decision={candidate.contactDecision}
                    contacts={references.contacts}
                    organizations={references.organizations}
                    onMatch={(contactAppId) => onPatch(candidate.id, { contactDecision: { type: "matched", contactAppId }, contactCreated: false })}
                    onIgnore={() => onPatch(candidate.id, { contactDecision: { type: "ignored" }, contactCreated: false })}
                    onReset={() => onPatch(candidate.id, { contactDecision: undefined, contactCreated: false })}
                    onAddPerson={() => onAddPerson(candidate.id)}
                  />
                ))}
              {(group.type === "ORGANIZATION" || group.type === "DISTRICT") &&
                group.candidates.map((candidate) => (
                  <EntityRoutingRow
                    key={candidate.id}
                    candidate={candidate}
                    noun="Organization"
                    matches={matchOrganizationCandidates(candidate.title, references.organizations, group.type as "ORGANIZATION" | "DISTRICT")}
                    entities={references.organizations}
                    onPatch={(patch) => onPatch(candidate.id, patch)}
                    onCreate={() => onCreateEntity(candidate.id)}
                  />
                ))}
              {group.type === "PROJECT" &&
                group.candidates.map((candidate) => (
                  <EntityRoutingRow
                    key={candidate.id}
                    candidate={candidate}
                    noun="Project"
                    matches={matchProjectCandidates(candidate.title, references.projects)}
                    entities={references.projects}
                    onPatch={(patch) => onPatch(candidate.id, patch)}
                    onCreate={() => onCreateEntity(candidate.id)}
                  />
                ))}
            </section>
          ))}
        </>
      )}

      {unsupported.length > 0 && (
        <>
          <p className="eyebrow">No destination yet</p>
          <p className="muted-copy">These stay selected and intact in your working session — nothing is written or discarded.</p>
          {unsupported.map((group) => (
            <UnsupportedGroup key={group.type} group={group} />
          ))}
        </>
      )}
    </section>
  );
}

function CompletedWorkQueue({
  group,
  skippedIds,
  onSkip,
  onUnskip,
  onLogAsWork,
}: {
  group: VoiceCandidateGroup;
  skippedIds: Set<string>;
  onSkip: (id: string) => void;
  onUnskip: () => void;
  onLogAsWork: (candidate: VoiceReviewCandidate) => void;
}) {
  const items = group.candidates;
  const pending = items.filter((candidate) => !candidate.loggedWorkRecordAppId && !skippedIds.has(candidate.id));
  const current = pending[0];
  const skippedCount = items.filter((candidate) => !candidate.loggedWorkRecordAppId && skippedIds.has(candidate.id)).length;

  return (
    <div>
      {current ? (
        <div className="contact-match-panel">
          <span>
            <small>
              Completed work {items.indexOf(current) + 1} of {items.length}
            </small>
            <strong>{current.title}</strong>
            {current.detail && <small>{current.detail}</small>}
          </span>
          <span className="contact-match-actions">
            <button type="button" className="candidate-log-button" onClick={() => onLogAsWork(current)}>
              Log this work
            </button>
            <button type="button" className="ghost-button" onClick={() => onSkip(current.id)}>
              Skip
            </button>
          </span>
        </div>
      ) : skippedCount === 0 ? (
        <p className="muted-copy">All selected completed work is logged ✓</p>
      ) : (
        <p className="muted-copy">
          {skippedCount} skipped for now.{" "}
          <button type="button" className="ghost-button" onClick={onUnskip}>
            Show skipped again
          </button>
        </p>
      )}
      {items.length > 1 && (
        <ul className="contact-match-candidates">
          {items.map((candidate) => (
            <li key={candidate.id}>
              <span>{candidate.title}</span>
              <small>
                {candidate.loggedWorkRecordAppId
                  ? "Logged ✓"
                  : skippedIds.has(candidate.id)
                    ? "Skipped"
                    : candidate.id === current?.id
                      ? "Up next"
                      : "Waiting"}
              </small>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

function EntityRoutingRow({
  candidate,
  noun,
  matches,
  entities,
  onPatch,
  onCreate,
}: {
  candidate: VoiceReviewCandidate;
  noun: "Organization" | "Project";
  matches: (Organization | Project)[];
  entities: (Organization | Project)[];
  onPatch: (patch: RoutingPatch) => void;
  onCreate: () => void;
}) {
  const decision = candidate.routing;
  // Current name if the entity is in the live list, else the name recorded at decision time.
  const entityName = (appId: string, recorded?: string) => entities.find((entity) => entity.appId === appId)?.name ?? recorded ?? noun;

  if (decision?.type === "created") {
    return (
      <div className="contact-match-panel matched">
        <span>
          <strong>{candidate.title}</strong>
          <small>
            Created {noun} ✓: {entityName(decision.entityAppId, decision.entityName)}
          </small>
        </span>
      </div>
    );
  }
  if (decision?.type === "matched") {
    return (
      <div className="contact-match-panel matched">
        <span>
          <strong>{candidate.title}</strong>
          <small>
            Matched {noun} ✓: {entityName(decision.entityAppId, decision.entityName)}
          </small>
        </span>
        <button type="button" className="ghost-button" onClick={() => onPatch({ routing: undefined })}>
          Change
        </button>
      </div>
    );
  }
  if (decision?.type === "ignored") {
    return (
      <div className="contact-match-panel ignored">
        <span>
          <strong>{candidate.title}</strong>
          <small>Ignored for this item</small>
        </span>
        <button type="button" className="ghost-button" onClick={() => onPatch({ routing: undefined })}>
          Reconsider
        </button>
      </div>
    );
  }
  return (
    <div className="contact-match-panel">
      <span>
        <strong>{candidate.title}</strong>
        {matches.length === 0 ? (
          <small>No exact match found.</small>
        ) : (
          <ul className="contact-match-candidates">
            {matches.map((match) => (
              <li key={match.appId}>
                <span>
                  {match.name} <small>Exact name match</small>
                </span>
                <button type="button" className="ghost-button" onClick={() => onPatch({ routing: { type: "matched", entityAppId: match.appId, entityName: match.name } })}>
                  Match Existing
                </button>
              </li>
            ))}
          </ul>
        )}
        <span className="contact-match-actions">
          <button type="button" className="ghost-button" onClick={onCreate}>
            Create {noun}
          </button>
          <button type="button" className="ghost-button" onClick={() => onPatch({ routing: { type: "ignored" } })}>
            Ignore
          </button>
        </span>
      </span>
    </div>
  );
}

function UnsupportedGroup({ group }: { group: VoiceCandidateGroup }) {
  const [copiedId, setCopiedId] = useState<string | null>(null);

  // Copy is a convenience only — it writes to the clipboard, never to any store.
  const copy = async (candidate: VoiceReviewCandidate) => {
    try {
      await navigator.clipboard.writeText(formatCandidateForCopy(candidate));
      setCopiedId(candidate.id);
    } catch {
      setCopiedId(null);
    }
  };

  return (
    <section className="voice-routing-group" aria-label={`${group.label} — no destination yet`}>
      <h3>
        {group.label} — {group.candidates.length}
      </h3>
      <p className="muted-copy">{group.label} — destination not yet implemented</p>
      <ul className="contact-match-candidates">
        {group.candidates.map((candidate) => (
          <li key={candidate.id}>
            <span>{candidate.title}</span>
            <button type="button" className="ghost-button" onClick={() => void copy(candidate)}>
              {copiedId === candidate.id ? "Copied" : "Copy"}
            </button>
          </li>
        ))}
      </ul>
    </section>
  );
}
