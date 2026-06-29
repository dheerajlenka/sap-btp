# Manifest AI Agent Hub — Interactive Inventory Dashboard

An interactive, shareable dashboard of the Manifest AI agent portfolio,
driven by the Excel inventory workbook as the single source of truth.

The deliverable is **one self-contained file — `dist/index.html`** — with all data
and charts baked in and **zero external dependencies**. You can:

- Publish it as a **Claude Artifact** and share the link with partners, customers
  and the internal sales team, or
- Host it on **GitHub Pages** / any static host, or
- Just email the HTML file — it opens straight from disk.

## What's in the dashboard

- **KPI band** — total agents, live clients, production users, documents/month, categories.
- **Agents** — searchable, filterable card grid (or table) with each agent's
  business function, problem, **Before → After** states, headline metric, and an
  expandable description.
- **Filters** — Product line (DeepDelve / DeepProbe / Cross-Platform), deployment
  status (Live / In Progress / Planned), and category. Filter state is encoded in the
  URL hash, so a filtered view is itself a shareable link.
- **Categories** — per-category panels with industries, agents, and live-deployment clients.
- **Insights** — charts: agents by category, by product line, by status, and live
  deployments by category.
- **Roadmap** — the backlog grouped by category with go-live sprints (**internal view only**).
- **ROI & Differentiators** — the ROI metrics table and key differentiators.
- **Export** — Print / Save as PDF, and Download CSV of the current filtered set.
- **Internal view toggle** — off by default; reveals internal comments and the Roadmap tab.

## How it's built

```
inventory/   source .xlsx (drop the latest workbook here)
build/
  convert.py   xlsx -> data/data.json  (normalizes the workbook)
  build.py     data/data.json + template -> dist/index.html
data/data.json normalized data (committed for reference)
src/dashboard.template.html  HTML/CSS/JS template with a {{DATA}} placeholder
dist/index.html  ⭐ the self-contained dashboard to share
```

## Updating the dashboard when the spreadsheet changes

1. Drop the new workbook into `inventory/` (any `.xlsx` with the same sheet layout).
2. Regenerate the data and rebuild:
   ```bash
   pip install openpyxl            # first time only
   python build/convert.py inventory/<your-file>.xlsx
   python build/build.py
   ```
   `dist/index.html` is now up to date. Commit, push, and re-publish the artifact /
   let GitHub Pages pick it up.

### Ad-hoc preview without rebuilding

A salesperson can click **"⬆ Load latest .xlsx"** on the live dashboard and drop in a
newer workbook. It is parsed **entirely in the browser** (nothing is uploaded) and the
view refreshes instantly. The published link still shows the baked-in data until you
rebuild — this is just a local preview.

> The in-browser reader uses the browser's native `DecompressionStream`, so no external
> library is bundled. It works in current Chrome, Edge, Safari and Firefox.

## Source workbook layout

The converter reads these sheets (others are treated as working copies and ignored):

| Sheet | Used for |
|-------|----------|
| `Master Data` | Categories → product line, industries, and live deployments |
| `Agent Inventory Live` | The 58 agents (before/after, metric, description, internal notes) |
| `Agents Backlog` | Roadmap (internal) |
| `Summary Dashboard` | KPIs, ROI metrics, key differentiators |

Notes / caveats:

- Categories `CAT-10` (Construction & Field Operations) and `CAT-11` (Procurement &
  Supply Chain) are not in `Master Data`; their product line defaults to *Cross-Platform*
  (override in `PRODUCT_LINE_FALLBACK` in `build/convert.py`).
- The **Internal view toggle is cosmetic, not a security boundary** — anyone can edit the
  URL. Keep genuinely sensitive notes out of the workbook columns that feed the dashboard.

## Verifying a build

```bash
python build/convert.py        # expect: 58 agents, 11 categories, 20 deployments
python build/build.py          # writes dist/index.html and fails if any external URL slips in
```
Then open `dist/index.html` in a browser (or `file://` it) and click through the tabs.
