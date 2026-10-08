---
target: dashboard
total_score: 21
max_score: 40
na_heuristics: 
p0_count: 2
p1_count: 2
target_identity: "file:/home/user/CRM/templates/dashboard.html"
target_fingerprint: "sha256:586f45909812e4109fc0f64f12191545fddf0fbd1ee8f94094f7d229a2e1aff8"
target_path: /home/user/CRM/templates/dashboard.html
timestamp: 2026-10-08T21-32-51Z
slug: templates-dashboard-html
---
Method: dual-agent (A: design review · B: detector + browser)

## Design Health Score
| # | Heuristic | Score | Key Issue |
|---|---|---|---|
| 1 | Visibility of System Status | 3 | Progress bar good; three different "today" totals (83 / 76 / 70) |
| 2 | Match System / Real World | 3 | "Buildings in today's tasks: 512" not actionable; unlabelled 🎯 column |
| 3 | User Control and Freedom | 2 | Skip and snooze have no undo |
| 4 | Consistency and Standards | 2 | Log/Log call/Done/Mark Paid; yellow means three things |
| 5 | Error Prevention | 2 | Skip adjacent to Log at 31px, no confirm |
| 6 | Recognition Rather Than Recall | 2 | "Day 1: Email 1", icon-only 🎯/📞 rely on hover |
| 7 | Flexibility and Efficiency | 2 | No shortcuts on dashboard; Queue visually secondary |
| 8 | Aesthetic and Minimalist Design | 1 | 10 stacked sections, ~7,500px mobile page |
| 9 | Error Recovery | 2 | Skip/snooze irreversible |
| 10 | Help and Documentation | 2 | Mostly hover titles |
| Total | | 21/40 | Below average |

## Design Specificity Verdict
Logic product-specific, visual language stock Bootstrap admin. Detector: CLI 0 (blind: Bootstrap unresolved); browser 6 desktop / 5 mobile — 4 low-contrast (#6c757d on #f4f6f9/#f8f9fa 4.3-4.4; #0d6efd on #cff4fc/#fff3cd 3.9-4.1), line-length ~144ch, 18 em dashes. 221/244 mobile targets <44px.

## Priority Issues
1. [P0] Mobile reminders table hides Log/Script/Skip off-canvas → stacked task cards with big Call + Log. /impeccable adapt
2. [P0] Tap targets <44px; Skip/snooze lack undo → 44px actions, phone as button, undo via Undo bar. /impeccable harden
3. [P1] No single next action; 10 equal sections compete with Queue → "Next up" hero + progress, trim KPIs, disclose the rest. /impeccable distill, layout
4. [P1] Color carries no meaning; glare contrast failures → token palette, one meaning per color, red only for truly late. /impeccable colorize, typeset
5. [P2] Conflicting totals and labels, fake midnight times → one reconciled today number, labelled columns. /impeccable clarify

## Persona Red Flags
- Desk calling block: no shortcuts; call→return→dropdown→outcome→reload per call; 25-row cap detours to Queue.
- iPhone in truck: Log off-screen; 16px stacked phone/email links; hover hints invisible; unlabelled 📞; grey outline buttons wash out.
- New rep: unexplained cadence jargon; dashboard vs Queue unclear; Skip semantics unclear; no record of who touched an account.

## Minor Observations
Back button inside h1; KPI 2-2-1 orphan on mobile; shared all=1 expands every section; 0.6rem progress bar; repeated "Accepted Meeting" badges.

## Questions to Consider
Should the dashboard be the first Queue card plus the day's score? Which KPI changes the next ten minutes? What does a truck-first version look like?
