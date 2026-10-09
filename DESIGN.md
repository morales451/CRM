# Design

Roof CRM looks like a conventional sales CRM, on purpose. The bar is Pipedrive, HubSpot and Close: quiet surfaces, one accent, dense tables, and a record page per account. Everything lives in `static/app.css`, which loads after Bootstrap 5 and re-themes it, so templates keep using plain Bootstrap classes.

Customer-facing documents (`bid_report.html`, `invoice_print.html`) don't use the app shell. They keep their own print styling.

## Shell

- **Desktop (992px and up):** a 232px dark sidebar (`.side`) and a sticky white top bar (`.topbar`) with global account search. The search is hidden on the Accounts page, which has its own.
  - Sidebar groups: Today, Queue, Accounts, Pipeline, Projects; then **Reports** (Insights); then **Setup** (Templates & settings, Import & export, Guide).
  - **Add account** is the sidebar's one primary button.
- **Phone and tablet:** the same sidebar becomes a Bootstrap off-canvas menu (`offcanvas-lg`).
  - A five-tab bottom bar (`.tabbar`: Today, Queue, Accounts, Pipeline, More) sits at the thumb.
  - Pages with their own bottom action bar (Queue, account) set `{% block body_class %}has-actionbar{% endblock %}` to hide the tabs.
- **Page header:** `.page-head` on every page holds the h1, a muted `.sub` line on the left and `.actions` on the right.

## Color

| Token | Value | Use |
|---|---|---|
| `--canvas` | #f4f5f7 | page background |
| `--surface` | #ffffff | panels, inputs, menus |
| `--line` / `--line-2` | #e4e7ec / #d0d5dd | borders, control borders |
| `--ink` / `--ink-2` / `--ink-3` | #101828 / #344054 / #475467 | text |
| `--muted` | #5d6879 | secondary text (AA on white and on the canvas) |
| `--accent` | #0070c0 | the logo blue (#0088DF) darkened for AA; links, primary buttons, focus |
| `--green` | #0f7a55 | Call, and anything that logs work |
| `--red` / `--amber` | #c0331d / #8a4b00 | overdue, destructive, warnings |
| `--side-bg` | #0f1724 | sidebar |

Color is for the primary action and for status. Every Bootstrap secondary variant (`btn-outline-*`, `btn-light`, `btn-info`, `btn-warning`, `btn-secondary`) renders as one neutral white button. `.btn-outline-success` and `.btn-outline-primary` tint only their icon. Badges are soft tinted pills, not solid fills.

## Type

- **Font:** Inter, self-hosted as a variable font (`static/vendor/fonts`), so the app works offline.
- **Body:** 14px on desktop. Inputs are 16px on phones so iOS doesn't zoom into them.
- **Scale:** h1 22px, h2 20px, h3 19px, h4/h5 15px, all at weight 650 with slight negative tracking.
- **Numbers:** tables, badges and stats use tabular figures.

## Components

- **Icons:** Lucide line icons in `static/icons.svg`. Use `{{ icon('name') }}` in templates, or `{{ icon('name', 'Label') }}` for an icon-only control. Asking for an unknown name raises an error. No emoji in the app. Inside buttons, menu items and pills, a flex gap spaces the icon.
- **Card:** a white panel with a 1px border, 8px radius and a hairline shadow. Card headers are white with a bottom rule. Don't nest cards.
- **Table:** 12px grey header row on `--surface-2`, 1px row rules and hover highlight. Numbers are right-aligned.
- **View tabs** (`.viewtabs`): saved views above a list, with an underline marking the active one and pill counts.
- **KPI tile** (`.kpi`): the label sits above the number.
- **Pipeline board** (`.board` / `.lane` / `.deal`): column headers with a 2px rule and white deal cards on the canvas.
- **Record header** (account page): an initials tile, the company name, status pills, then the action row.
- **Status dots** (`.dot-done/-due/-wait/-skip`): cadence progress.
- **Undo bar** (`.undo-bar`): a dark toast-style bar at the top of the page.
- **Empty state** (`.empty`): a centered icon and a sentence.

## Finish

- **Browser surfaces:** focus rings, text selection, scrollbars and `kbd`/`code` all use the palette.
- **Motion:** respects `prefers-reduced-motion`.
- **Print:** the shell is hidden when printing.
- **Phones:** the bottom bars respect safe-area insets.
