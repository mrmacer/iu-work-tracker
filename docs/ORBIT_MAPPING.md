# ORBIT Reporting Mapping

ORBIT is an optional classification and reporting projection over canonical Work Records. Non-ORBIT work follows the same daily workflow and remains fully valid.

## Deliverables

| Code | 2026–2027 deliverable | Work Record mapping |
| --- | --- | --- |
| A | Strengthening Statewide STEM & CS Systems | primary or supporting deliverable code `A` |
| B | Supporting Implementation of PA STEELS | code `B` |
| C | Strengthening CS, AI, and Computational Thinking | code `C` |
| D | Equipping Educational Leaders for Student-Centered STEM/CS | code `D` |
| E | Workforce Collaboration and Ecosystem Development | code `E` |
| F | Student Competitions and Experiences | code `F` |
| G | Math Instruction and Data Literacy through Transdisciplinary STEM | code `G` |

A reportable record has one primary deliverable for official grouping and zero or more supporting deliverables for honest cross-cutting analysis.

## Metric mapping

| ORBIT need | Source/derivation |
| --- | --- |
| STEM PoC time | `orbit.stemPocMinutes`; sum and divide by configured minutes per day |
| TaC time | `orbit.tacMinutes`; sum and divide by the same or role-specific configured rule |
| Educators/leaders | `reach.educatorsLeaders` |
| Students/families | `reach.studentsFamilies` |
| Workforce/community | `reach.workforceCommunity` |
| LEAs served | distinct linked Organization AppIds with an LEA type |
| Reporting quarter | derived from activity date and versioned school-year quarter settings |
| Qualitative evidence | output, outcome, success/evidence summary, notes, linked Evidence records |

Source duration remains in minutes. Reporting days are always calculated (`minutes / configuredMinutesPerDay`, initially 420) so a future policy change does not corrupt activity history.

## Reporting projection

1. Select reportable Work Records in the requested date range.
2. Derive the applicable school year and quarter from each activity date.
3. Group official totals by primary deliverable; keep supporting-deliverable attribution available as a separate analytical view to prevent double-counting.
4. Sum role minutes and reach; count distinct LEAs.
5. include linked qualitative evidence and human-review flags.

V1 summaries are indicative planning views, not an official ORBIT export. A later export must document aggregation, double-counting, and evidence-selection rules and should provide a review queue before submission.
