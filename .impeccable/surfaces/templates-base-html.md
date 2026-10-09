---
version: 1
slug: "templates-base-html"
primary_target: "templates/base.html"
related_targets: []
---

THESIS
Roof CRM should look and behave like a real sales CRM, played straight: the category canon at the craft level of Pipedrive, HubSpot and Close. The owner asked for exactly this ("give it a real CRM look") after calling the old look AI slop. Convention is the commitment; no novelty styling.

OWN-WORLD
Canon path, not an own-world page. Quiet neutral surfaces (canvas #f4f5f7, white panels with 1px #e4e7ec borders), one blue accent from the logo (#0088DF, darkened to #0070c0 for text and fills), green reserved for Call and logging, red/amber only for overdue and status. Inter (vendored, offline). Lucide line icons from a local SVG sprite; no emoji anywhere in the working app. Customer-facing documents (bid report, invoice) keep their own print styling.

STORY
Every page sits in one app shell: dark left sidebar (Today, Queue, Accounts, Pipeline, Projects; Reports; Setup) with Add account on top, a white top bar with global account search. On phones the sidebar becomes an off-canvas menu and a five-tab bottom bar (Today, Queue, Accounts, Pipeline, More) appears, except on the Queue and account pages, which keep their own thumb action bar.

FIRST VIEWPORT
Today: page header with date, order toggle and Work the queue; the Next up panel with its one primary action; the cadence steps table. Account: record header (initials tile, name, status pills, action row), then the next-step panel.

FORM
Page header pattern (h1 + muted sub + right-aligned actions) on every page. Saved-view tabs above the accounts list. Dense tables with small grey headers, right-aligned numbers, tabular figures. KPI tiles with the label above the number. Pipeline as a kanban board of white deal cards under column headers. Secondary buttons are one neutral white style whatever Bootstrap colour they were named; one filled primary per region.

FINISH
Themed focus rings, selection colour, thin scrollbars, tabular numerals, 16px inputs on phones, safe-area padding on the bottom bars, reduced-motion respected, print hides the shell. Contrast AA on every text role including placeholders.
