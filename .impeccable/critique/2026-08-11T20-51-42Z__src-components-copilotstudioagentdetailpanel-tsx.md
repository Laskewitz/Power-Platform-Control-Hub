---
target: this Copilot Studio agent detail page
total_score: 21
max_score: 40
na_heuristics: 
p0_count: 0
p1_count: 3
timestamp: 2026-08-11T20-51-42Z
slug: src-components-copilotstudioagentdetailpanel-tsx
---
Method: dual-agent (A: agent-detail-reviewer · B: agent-detail-detector)

## Design Health Score

| # | Heuristic | Score | Key issue |
|---|---|---:|---|
| 1 | Visibility of system status | 3 | State badges are strong, but quarantine availability is not explained. |
| 2 | Match system / real world | 2 | Raw capability keys and Dataverse terminology leak into the UI. |
| 3 | User control and freedom | 2 | Navigation and unquarantine exist, but destructive outcomes have no undo. |
| 4 | Consistency and standards | 3 | The visual system is coherent; analysis rows use nonstandard interaction. |
| 5 | Error prevention | 2 | Delete is confirmed; quarantine lacks a consequence checkpoint. |
| 6 | Recognition rather than recall | 2 | Findings do not visibly advertise that they expand. |
| 7 | Flexibility and efficiency | 2 | Dense but mouse-oriented; no shortcuts or direct remediation. |
| 8 | Aesthetic and minimalist design | 2 | Duplicate facts and equal visual weight bury the important signals. |
| 9 | Error recovery | 2 | Privilege recovery is strong; generic action failures are less useful. |
| 10 | Help and documentation | 1 | Technical terms and high-stakes actions lack contextual explanation. |
| **Total** | | **21/40** | **Acceptable; significant improvement needed** |

## Design Specificity Verdict

The graphite, cyan, and amber control-room language is authored and coherent, but the composition is still a category-interchangeable accordion inspector. Copilot governance does not shape the page strongly enough. The deterministic scan returned zero findings, confirming that the weakness is information architecture rather than a detectable design-system violation. Browser overlay injection was unavailable because the shared Power Apps page had redirected to Microsoft sign-in; no reliable overlay was produced.

## Overall Impression

The surface feels operational and trustworthy, but metadata dominates the administrator's real task. The first viewport should answer whether this agent is safe, exposed, and actionable.

## What's Working

- Agent identity, environment, active state, and quarantine state are anchored clearly.
- The control-room visual language is consistent with the rest of Control Hub.
- Admin-access failure recovery explains the limitation and offers direct remediation.

## Priority Issues

1. **[P1] Governance risk is subordinate to metadata.** Best Practice Analysis appears below routine inventory. Move posture and findings directly below the header; collapse routine details by default. Suggested command: `/impeccable layout`.
2. **[P1] Quarantine lacks a consequence checkpoint.** Explain impact, scope, and reversibility before applying it. Suggested command: `/impeccable harden`.
3. **[P1] Analysis findings are click-only rows.** Replace clickable divs with keyboard-operable buttons or accordion controls with chevrons and `aria-expanded`. Suggested command: `/impeccable audit`.
4. **[P2] Raw system language burdens interpretation.** Translate capability keys and move GUIDs/raw configuration into Technical details. Suggested command: `/impeccable clarify`.
5. **[P2] All information has similar visual weight.** Elevate exposure and critical findings; mute duplicate statuses and inventory facts. Suggested command: `/impeccable distill`.

## Persona Red Flags

- **Power admin:** must scan multiple expanded sections before reaching actionable risk; findings do not lead directly to remediation.
- **Keyboard or screen-reader user:** analysis rows cannot be expanded reliably with keyboard and do not expose expanded state.
- **New administrator:** Definition, capability keys, Entra IDs, and quarantine consequences assume unexplained domain knowledge.

## Minor Observations

- Type and active-state badges both use green, weakening semantic distinction.
- Ten-pixel uppercase labels may be too small for sustained scanning.
- “Re-analyze” appears to reload more than analysis alone.
- The body scrollbar competes with the host page and makes location harder to track.

## Questions to Consider

- What if the first viewport answered only: Is this agent safe, exposed, and actionable?
- Why are inventory facts privileged over governance findings?
- Could each finding end in a concrete admin action?
