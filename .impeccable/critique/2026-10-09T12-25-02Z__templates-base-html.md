---
target: whole app (critique)
total_score: 26
max_score: 40
na_heuristics: 
p0_count: 0
p1_count: 4
target_identity: "file:/home/user/CRM/templates/base.html"
target_fingerprint: "sha256:619566955840a8d6e592865feca46b7d0a7d2f23726b70e9ac4340b3a2b7adfc"
target_path: /home/user/CRM/templates/base.html
timestamp: 2026-10-09T12-25-02Z
slug: templates-base-html
---
# Critique: app shell (base.html + app.css), whole app
Method: dual-agent (A: design review · B: detector/browser evidence)

| # | Heuristic | Score | Key issue |
|---|---|---|---|
| 1 | Visibility of system status | 3 | Pipeline 2544px wide, 4.5 of 9 lanes visible |
| 2 | Match system / real world | 2 | "None / In Cadence" enum leaks; Research account shows Status Prospecting |
| 3 | User control and freedom | 3 | Pipeline select auto-submits |
| 4 | Consistency and standards | 2 | 4 pages lack .page-head; title/nav mismatch; two blue Add account on /accounts |
| 5 | Error prevention | 2 | Bad number tile next to everyday outcomes; pipeline moves on change |
| 6 | Recognition rather than recall | 3 | Matching/bldgs unexplained |
| 7 | Flexibility and efficiency | 3 | no shortcuts outside Queue; blank top bar on /accounts |
| 8 | Aesthetic and minimalist design | 2 | ~105 buttons on Today desktop; 9-action account header; two logging surfaces |
| 9 | Error recovery | 3 | red dashed goal line reads as error |
| 10 | Help and documentation | 3 | long help paragraphs |
| Total | | 26/40 | Acceptable |

Specificity: faithful CRM canon; gaps in domain views (pipeline, enum leaks, too many filled buttons).
Detector: CLI 2 low-contrast (invoice watermark, FP). Browser: overused-font x16 (FP, intentional), cramped-padding x6 (account list-group-item px-0 on grey bg, real), line-length x9 (import/templates help, minor), flat-type-hierarchy x1 (phone queue h2 15px < body 16px). No overflow. Tap targets: topbar search/menu 36x25, checkboxes 14x14, inline links ~17px.

## Priority issues
- [P1] Too many filled buttons on account page and Queue; Spoke tile looks pre-selected. Fix: neutral outcome tiles with colored text; hide header Call when next-step shows Call; neutral header Email on Queue. (/impeccable quieter)
- [P1] Internal enum/status contradictions ("None / In Cadence", Research vs Prospecting). Fix: display label or dash; single status source; drop uniform Milestone column. (/impeccable clarify)
- [P1] Phone nav buttons 36x25 top-right; only nav on action-bar pages; duplicate menu triggers. Fix: 44px targets, bottom Back/Today on action-bar pages, bigger checkboxes/links. (/impeccable adapt)
- [P1] Pipeline is dropdown list, not pipeline: no $/age, cadence lane of 114, onchange submit, nested scroll on phone. Fix: drop cadence lane, $ and days-in-stage, lane totals, Move menu, phone stage tabs. (/impeccable shape, layout)
- [P2] Faint focus ring (~1.4:1) and light borders (#d0d5dd). Fix: solid 2px accent outline, borders ~#98a2b3. (/impeccable audit)

## Persona red flags
Alex: no search shortcut; pagination top only; no bulk pipeline move; repeated row buttons on Today.
Sam: faint focus; 14px checkboxes; select fires on arrows; unlabeled Your Info fields; color-only outcome meaning.
Casey: 25px top buttons; Call shown on Email steps in Today phone list; scroll-in-scroll email body; undo bar at top.

## Minor
Insights red goal line, calls breakdown 5 of 6; queue progress sliver; filter search taller than selects; account list rows flush-left; underlined "Email bounced?"; "0 matching of 39" vs "39 owned"; capitalization inconsistencies; long help lines; phone Queue h2 < body.

## Questions
Should Today stop repeating the queue as a table? Should Pipeline start at Accepted Meeting and be about bid dollars? Should the phone account page drop the header actions and Log something form?
