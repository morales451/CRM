---
target: account detail
total_score: 23
max_score: 40
na_heuristics: 
p0_count: 2
p1_count: 2
target_identity: "file:/home/user/CRM/templates/account_detail.html"
target_fingerprint: "sha256:8779250f87c931690a5c11bb5aee786488c9666bfaa631b35ad8eda8d0c1cec5"
target_path: /home/user/CRM/templates/account_detail.html
timestamp: 2026-10-08T22-08-16Z
slug: templates-account-detail-html
closed: true
---
Method: dual-agent (A: design review · B: detector + browser)

## Design Health Score
| # | Heuristic | Score | Key Issue |
|---|---|---|---|
| 1 | Visibility of System Status | 2 | Due step buried in 4th card |
| 2 | Match System / Real World | 3 | Good vocabulary; "Total Props", General Note default |
| 3 | User Control and Freedom | 3 | Undo/confirm good; Save easy to miss |
| 4 | Consistency and Standards | 2 | 3 outcome-logging controls; colors carry no meaning |
| 5 | Error Prevention | 2 | Log defaults to General Note; status change silently stops cadence |
| 6 | Recognition Rather Than Recall | 2 | Outcome effects in prose; hover-only icon buttons |
| 7 | Flexibility and Efficiency | 2 | No one-tap call outcomes on this page |
| 8 | Aesthetic and Minimalist Design | 1 | ~25 fields, ~35 buttons all open |
| 9 | Error Recovery | 3 | Plain confirms, paste preview respects edits |
| 10 | Help and Documentation | 3 | Great but always-open help |
| Total | | 23/40 | Below average |

## Design Specificity Verdict
Half product-specific (door knock, cadence, ZoomInfo, matching), half generic Bootstrap record view. Detector: CLI 0 (blind); browser 10 — 5 low-contrast (Bootstrap #0d6efd/#6c757d/#198754 on #f4f6f9/#f8f9fa, 4.2-4.3), 3 undersized badges (10.5px), 1 nested-cards, 1 skipped-heading. Mobile layout 687px on 390 viewport (text-nowrap header action row). 106/116 mobile targets <44px. Log Interaction at y≈3171 mobile.

## Priority Issues
1. [P0] Mobile overflow + 31px buttons; ✕ next to ★ → wrap header actions, top-3 large + More, thumb bar, ✕ into menu. /impeccable adapt
2. [P0] Next action not visible → "Next up" strip under name with Call + Log outcome. /impeccable layout, clarify
3. [P1] Logging slow, defaults to General Note → step-aware log with one-tap outcomes showing effects, near top. /impeccable distill, harden
4. [P1] Always-open edit form leads; unsaved edits lost when logging; status change silently stops cadence → read-only + Edit, contact edited in place, confirm cadence stop. /impeccable layout
5. [P2] Rare tools equal weight → collapse into Research tools. /impeccable quieter

## Persona Red Flags
- Desk: General Note trap; separate saves lose edits; no shortcuts.
- iPhone: horizontal pan; Call/Log 2,800px apart; 31px targets; hover-only labels; Door knock menu obstructed; no thumb bar.
- New rep: no "where things stand" summary; unexplained badges; ▶ Next reads as navigation; no author on log entries.

## Minor Observations
Cadence shows due dates for done steps (Oct 05 before Oct 07 in demo); green Visited badge regardless of outcome; uniform grey timeline badges; Created timestamp low value; mobile label wrap/placeholder clip; extra contacts lack Call buttons.

## Questions to Consider
Open on "what to do with Beth now"? Why can a door knock be logged in two taps but not a call? Should editing ever silently stop a cadence?
