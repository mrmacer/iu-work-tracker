# V1 Plan

## Goal

Prove that one universal Work Record can make daily documentation fast while also supporting projects, districts, and optional ORBIT reporting.

## Included

- Finite kiosk-style shell and command center.
- Five-step Quick Log with progressive disclosure, Save & Done, and Save & Log Another.
- Today view with time, records, follow-ups, and quick add.
- Work History with search/filter and basic edit flow.
- Projects workspace backed by canonical project entities and calculated record totals.
- STEM/ORBIT workspace with reportable time, configurable seven-hour day conversion, deliverable distribution, reach, and explicit development-data labeling.
- Canonical Work Record TypeScript model plus projects, organizations, contacts, categories, evidence, follow-up, and configuration types.
- DataProvider boundary with a prototype provider; future SharePoint adapter contract documented.
- Small, clearly labeled development sample set spanning STEM and non-STEM work, districts, partners, students, and projects.
- Responsive and accessible keyboard/touch interaction.

## Deferred

Microsoft Entra sign-in, Graph synchronization, SharePoint list creation, file upload, Power Automate, reminders, contacts CRM, official ORBIT export, advanced dashboards, multi-user roles, and production performance claims.

## Implementation slices

1. Domain schema, configuration, provider contract, and labeled sample seed.
2. Recognizable command-center shell and navigation.
3. Quick Log workflow and persistence.
4. Today, History/editing, Projects, and ORBIT summaries.
5. Responsive/accessibility refinement, validation, and deployment.

## Acceptance checks

- A first-time user can find Log Work in under five seconds.
- A routine record requires only activity, audience/context, time, result, and optional ORBIT steps.
- Non-STEM records never require ORBIT fields.
- A saved or edited record appears consistently in Today, History, Projects, and summary calculations.
- Provider-specific calls do not appear in screen components.
- Sample metrics are always labeled development data.
