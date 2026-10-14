# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users
Alexis Morales, owner of Silicone Roof Pros (Houston, TX), runs the whole sales motion alone today: prospecting commercial building owners, cold outreach, roof inspections, bids and project billing. One to three more people (a sales rep or office help) may share the same data later, so screens should never assume only one person has ever touched an account. Building owners never use the app, but they receive what it prints: roof reports, bids and invoices.

## Product Purpose
A local CRM for selling silicone roof coatings to commercial building owners. It turns a CoStar owner list into worked accounts: research decision-makers (ZoomInfo), run a 5-step cold cadence (email, call and text, call, email, breakup) one person at a time per company, book inspections, produce roof reports and bids, then track projects and invoices after a win. Success means more calls made and logged per day, nothing slipping through, and the next action always obvious.

## Positioning
Built around one roofer's real loop rather than a generic pipeline: owner portfolios ranked by buildings that fit the sales criteria (pre-1980s), a business-day cadence whose clock resets from each action, one-tap call outcomes, ZoomInfo exports that import as-is, and door-knock tracking for the field. Runs on the owner's own computer, so the data stays local with no monthly fee.

## Operating Context
- Used on a desktop browser and on an iPhone in equal measure: desk blocks for imports, email and calling, and the phone in the truck for calls, texts, door knocks and quick logging between stops. The phone reaches the app over Wi-Fi (or Tailscale), served from the desktop.
- Calls and texts go through the phone's dialer or Google Voice, and emails through Gmail (computer) or Apple Mail (phone). The CRM never sends anything itself.
- Inputs: CoStar exports (Excel/CSV), ZoomInfo contact exports and page pastes.
- Outputs: printed or PDF roof reports, bids and invoices that go to building owners, plus Excel exports.
- Rituals: the daily dashboard and Queue (one task at a time), a calling block, door-knock days planned from a zip-sorted list.

## Capabilities and Constraints
- Flask + SQLite + Jinja + vendored Bootstrap 5, served locally by waitress. No build step, and no external CDN at runtime (works offline on the LAN).
- No login or PIN screen, and no SMTP or automatic sending, by explicit owner decision.
- Data lives outside the app folder (Documents/RoofCRM), and the app updates itself on start.
- Must stay fast on a large owner list (thousands of accounts and tasks).
- Terminology: account (owner company), primary contact, cadence step, Research (no contact yet), follow-up, roof report/bid, project, door knock.
- Undecided: multi-user support (separate logins, ownership of accounts) isn't built yet and may be needed when help is hired.

## Brand Commitments
- Company name "Silicone Roof Pros, Inc." and its logo (`static/brand-logo.png`) stay as they are, especially on customer-facing documents. Colors, type and the rest of the visual language are open.
- Voice to owners: plain, direct, expert contractor, never salesy hype. Alexis signs as "Alex M".
- The app itself looks like a real sales CRM, played straight: the category standard at the craft level of Pipedrive, HubSpot and Close. Left sidebar navigation (a bottom tab bar on the phone), quiet neutral surfaces, one blue accent taken from the logo (#0088DF), Inter, drawn line icons (no emoji), dense readable tables and record pages. No novelty styling in the working app.

## Evidence on Hand
- Real product data lives only on the owner's machine. Test fixtures use fictional people and companies (`tests/fixtures/`).
- No customer testimonials, case studies or project photos are in the repo. Don't invent any for customer-facing documents.

## Product Principles
1. The next action is always obvious: one task at a time, with the step, script and contact on one screen.
2. Fewer taps between "done" and "logged". Logging is the habit the whole system depends on.
3. Equal on phone and desk: field use is first-class, not a squeezed desktop.
4. Never silently lose or invent: undo for anything destructive, missing info shown in [brackets], no guessed data.
5. Customer-facing output reads like a seasoned contractor wrote it.

## Accessibility & Inclusion
Used outdoors on a phone (in glare, one-handed, between stops), so it needs large tap targets, strong contrast and short readable labels. No other product-specific standard has been set.
