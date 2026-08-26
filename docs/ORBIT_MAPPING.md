# ORBIT Reporting Mapping — V1.1

ORBIT is an optional projection over a canonical Work Record. A non-reportable record is valid without ORBIT classification or reporting time. A reportable record requires exactly one valid primary deliverable and may include distinct supporting deliverables behind progressive disclosure.

## Deliverables

| Code | Runtime label |
| --- | --- |
| A | Statewide STEM & CS systems |
| B | PA STEELS implementation |
| C | CS, AI & computational thinking |
| D | Educational leadership |
| E | Workforce & ecosystem development |
| F | Student competitions & experiences |
| G | Math instruction & data literacy |

The primary code drives official grouping. Supporting codes preserve cross-cutting context but do not add another copy of the record or its time.

## Time rules

The model distinguishes:

- `durationMinutes`: total elapsed/work duration for the Work Record.
- `orbit.stemPocMinutes`: the portion allocated to STEM Point-of-Contact reporting.
- `orbit.tacMinutes`: the portion allocated to TaC reporting.

V1.1 uses the conservative allocation rule:

```text
stemPocMinutes >= 0
tacMinutes >= 0
stemPocMinutes + tacMinutes <= durationMinutes
```

PoC and TaC may coexist only as non-overlapping allocations. The same minute cannot be counted in both. This rule was chosen because the existing V1 material did not authorize overlapping role time; conservative allocation avoids silently inflating institutional totals.

Non-reportable records store no ORBIT deliverables or reporting minutes. Turning ORBIT off in Quick Log clears those fields.

## Pure derivations

`lib/reporting.ts` supplies tested, side-effect-free calculations:

- School year: July 1 starts the next labeled school year. `2026-07-01` is `2026-2027`; `2026-06-30` is `2025-2026`.
- Quarter: Q1 July–September, Q2 October–December, Q3 January–March, Q4 April–June.
- Reporting days: `stemPocMinutes / minutesPerReportingDay`.
- PoC/TaC projection: returns stored allocation only when `reportable` is true.

`minutesPerReportingDay` is supplied by `ReportingConfig` and is currently `420`; it is not scattered through UI code. School year and quarter are derived rather than stored.

## Validation

Shared validation rejects missing or invalid primary deliverables, duplicate supporting codes, the primary code repeated as supporting, invalid codes, malformed arrays/ORBIT objects, negative reach/time, reporting allocation above duration, and hidden ORBIT data on a non-reportable record.

The ORBIT screen reuses the Work Record’s reach and reporting allocations. It never creates a reporting-only business record.
