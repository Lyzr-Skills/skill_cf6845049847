---
name: m3-release-impact
description: Cross-reference an Infor M3 release report against the APIs and components a customer actually uses, and produce an interactive HTML impact-analysis report. Use whenever the user has an M3 Release Report (a "Report" sheet whose first column is the release version, e.g. 2026.07.00) plus an API Release Management workbook (sheets such as "Distinct APIs", "Connection Points", "Workflows", "H5 Scripts", "XtendM3 APIs") and wants to know how upcoming M3 releases affect their existing applications. Trigger on "M3 release impact", "impact analysis", "analyze this release report against our APIs", "which of our APIs/integrations are affected by the new release", "what changed in M3 that hits us", "release notes impact", or whenever both workbooks are present and the goal is release impact, finding new-API opportunities, or reviewing the release landscape by module — even if the user never says "skill".
---

# M3 Release Impact Analysis

Generate a single self-contained interactive HTML report that tells an M3 customer
**which upcoming release changes touch the APIs and components they actually use**,
plus a landscape view of everything in the release window and a list of adoptable
new-API opportunities.

## Inputs (two Excel workbooks)

Nothing about sheet names, component types, or column positions is hard-coded in the
script beyond locating the master "used APIs" list — everything else is **discovered
from each workbook at runtime**, because the dataset differs on every run (different
customer footprint, different component types, possibly reordered columns).

1. **M3 Release Report** — Infor's release notes export; the data lives in the `Report`
   sheet. The script scans for the header row (the row containing `Release`) and maps
   columns **by header name** — `Release`, `Component`, `Issue Type`, `Summary`,
   `Overview of Change`, `Detailed Description`, `Reference ID` — falling back to their
   usual positions only if a header is missing. So a leading blank row or a reordered
   export still parses. The sheet holds several years of history; the script filters by
   release version (see cutoff below).

2. **API Release Management workbook** — the customer's footprint:
   - A **master "used APIs" sheet** (typically named `Distinct APIs`) — the authoritative
     list of `PROG/Txn` the customer calls. Located by name; if absent, the used set is
     derived from the component sheets instead.
   - **Component sheets** — *every other sheet* is treated as a component type, and the
     **sheet name becomes the component-type label**. In the example workbooks these are
     `Connection Points`, `Workflows`, `H5 Scripts`, `XtendM3 APIs`, `H5 SDK APIs`,
     `IEC Mapping APIs`, `Event Based API`, but the code does not assume that list — add or
     rename a sheet and it flows through to the filters and the By Component view
     automatically. Within each sheet the join column is the one whose header contains "API"
     (the M3 transaction called); the component **name** is the other column. Header
     detection matters for sheets like XtendM3 where an MI-looking value appears in *both*
     the extension's own column and the called-API column.
   - **Versioned names** — if a component sheet's name-column header contains "version"
     (e.g. the IEC sheet's `IEC Mapping (Version)`), the trailing ` (version)` is stripped so
     the same artifact across many versions collapses to one. Each program's component list
     is then de-duplicated by (type, name, API). This matters for big sheets — the IEC sheet
     can carry tens of thousands of version rows for a few hundred real mappings.
   - **Display cap** — to keep files small and cards readable, the affected-component grid
     is capped per impact (120 in the analyst report, 60 in the customer dashboard) with a
     "+N more not shown (N total)" note; the headline "Affects N component(s)" always shows
     the true total.

   The master sheet defines *what* is used; the component sheets define *where*.

If either workbook is missing, ask for it rather than guessing.

**Workbook identity (filename).** The API workbook is named
`{tenant}~{CustomerName}~API_Release_Mgt~{M3|All}~{date}.xlsx`. The skill parses the tenant
and customer name from this filename and displays them in the header of **both** output files
(the analyst report and the customer dashboard). It also tolerates the underscore-sanitized
form (when `~` is replaced by `_` on upload). If the name doesn't match, the identity tags are
simply omitted.

## How to run

**Runtime.** This skill runs on **Node.js** (v18+) — no Python required. SheetJS is
vendored at `scripts/vendor/xlsx.mini.min.js`, so there is **no `npm install` step**: the
script runs with just `node`. Everything else uses Node built-ins. This is the version to
use from an agent that can shell out to `node` but not `python`. (If you ever remove the
vendored file, the script falls back to a normally-installed `xlsx` package.)

The matching and HTML build are deterministic — run the bundled script, do not
re-implement the matching by hand. Because the **Impact Summary** is a judgement
step (it reads Infor's over-explained multi-column prose and distills it), it is
produced by *you*, not the script, via a two-pass workflow:

**Pass 1 — discover the impacted items and read their full text:**
```bash
node scripts/m3_impact_report.mjs \
  --release-report "<M3 Release Report>.xlsx" \
  --api-release    "<API Release Mgt>.xlsx" \
  [--cutoff 2026.06] \
  --emit-impacts /tmp/impacts.json --out /tmp/_pre.html
```
`/tmp/impacts.json` is a list of the impacted items with their full `summary`,
`overview`, `detailed`, `progs`, `txns`, and matched `affected` components.

**Author summaries.** Read `impacts.json` and, for each unique `ref`, write a short
Impact Summary grounded in the overview + detailed text. Keep it to two labelled
lines and never pad:
```
Changed: <one sentence — what was fixed / introduced / enhanced>.
Impact: <one sentence — the effect on THIS customer's integration; whether action is needed>.
```
For **defects**, state that the previously-wrong behaviour is now corrected (and that
any local workaround can usually be retired / re-tested). For **enhancements**, state
that it is optional/additive on a transaction they already call. Reference the matched
transaction or component where it sharpens the point. Save as JSON keyed by ref:
```json
{ "RN-9348346": "Changed: …\nImpact: …", "RN-9312598": "Changed: …\nImpact: …" }
```
(Keying by the exact `summary` text also works if a ref is missing. Duplicate refs
across releases reuse the same summary.)

**Pass 2 — build the final report with summaries injected:**
```bash
node scripts/m3_impact_report.mjs \
  --release-report "<M3 Release Report>.xlsx" \
  --api-release    "<API Release Mgt>.xlsx" \
  [--cutoff 2026.06] \
  --summaries /tmp/summaries.json \
  --out M3_Release_Impact_Analysis.html
```

Then present the resulting HTML with `present_files`. Keep the spoken summary short:
counts per severity, the High items, the heaviest-hit program, and a pointer to the
Opportunities tab.

(If the user explicitly wants a quick report without summaries, skip the two passes and
run once without `--emit-impacts`/`--summaries`.)

### Release filtering (cutoff + toggle-expiry pre-filter)

`--cutoff` is `YYYY.MM` and **defaults to the current month**. Filtering uses only the
year and month of a release version, never the day, so `2026.06` keeps `2026.06.00`,
`2026.06.04`, `2026.07.00`, … and drops older. The `.00` day is intentionally included —
Infor sometimes back-dates mid-month fixes to `.00`.

Before that release filter runs, the dataset is assembled as **Set A ∪ Set B** using the
`Expiration of Toggle Ability` column. That column is itself an M3 release version
(e.g. `2026.10.00`), so it is compared by year.month exactly like the release column:

- **Set A** — items whose toggle is blank or already expired (`expiry < cutoff`): keep
  only if `release >= cutoff` (the normal rule).
- **Set B** — items whose toggle window is still open (`expiry >= cutoff`): keep
  regardless of release version. This is the important part: a *past-released* enhancement
  whose activation window is still open is still actionable (you can still turn it on, and
  enabling it changes behaviour), so it must not be dropped just because it shipped before
  the cutoff.

The two sets are disjoint (the expiry condition partitions them) and their union is the
analysis input for both the Impact and Landscape views. Every item carries its expiry
version; items that have one show a `⏻ toggle expires <version>` badge, and the Impact and
Landscape tabs each have a **Toggle expiry** dropdown (mirroring the release dropdown) to
filter by it.

## Matching logic (the core — keep this stable)

For each release item, find program codes matching `XXXnnnMI`. Consider the item only
if it mentions a **program the customer uses**. Then decide scope:

- **Specific** — the fix names an exact transaction, either attached (`OIS100MI/ChgWarehouse`,
  `OIS100MI.ChgWarehouse`) **or in prose** ("the API transaction `AddBatchLine` in OIS100MI",
  "…when changing warehouse using `ChgWarehouse`"). Only components calling that exact
  transaction are flagged — and only transactions the customer actually uses. If the named
  transactions are ones the customer does **not** use, the item is **not** an impact (it
  still appears in the Landscape tab).
- **General** — no transaction is named (the fix references panels/programs like OIS275,
  OIS345 in prose, not an MI transaction). Flag **every** component on that program, since
  a program-level change could touch any call.

Prose transaction detection relies on a curated verb-prefix list (`Add|Chg|Del|Get|Lst|
List|Upd|Sel|...`) near the top of the script as `VERBS`. **This is the single most
important thing to maintain**: an unrecognised verb means a named transaction is missed
and the item wrongly falls back to General scope (flagging far too many components). When
adding support for a new release, skim a few items and, if a real transaction name slips
through, add its verb prefix to `VERBS`.

### Severity

- **High** — Defect on an exact transaction the customer calls.
- **Medium** — Enhancement to an exact transaction they call, OR a program-level (General) Defect.
- **Low** — Enhancement at program level.

## Module taxonomy (Landscape + Opportunities grouping)

The `Component` column is rolled up into top-level modules by `module_of()` in the script.
The modules that appear are entirely data-driven — only buckets present in the current
run are shown — and anything unrecognised falls through to `Other / Add-ons`, so an
unfamiliar component string never breaks the report. `module_of()` is a normalisation
heuristic, not a fixed output: extend its rules if a new component family appears and you
want it broken out of `Other / Add-ons`.

## What is and isn't fixed

Fixed in the script are only **domain heuristics**, not dataset values: the MI verb-prefix
list (`VERBS`) used for prose transaction detection, the opportunity keyword set
(`OPP_RE`), and the `module_of()` normalisation map (with a catch-all). Everything tied to
a particular customer or release — component types, programs/transactions, modules,
release versions, counts, the report's columns — is read from the supplied workbooks each
run. Treat the heuristics as editable config; never bake a specific dataset's values into
the code.

## Opportunities

An item is flagged as an opportunity when it is an **Enhancement** whose text matches
"new API / new MI / Add MI Program / new transaction / new output field / new field /
new input field / now possible / new option". These are capability additions the
customer could adopt. In the report they are sorted so items on programs the customer
already uses come first (lowest adoption effort).

## Output (two files)

Every run produces **two** self-contained HTML files:

### 1. Analyst report (`--out`)
Dark-themed, interactive, for the consultant. Four tabs (Impact, By Component, Release
Landscape, Opportunities). The page opens with a **gradient hero** — eyebrow, the working
month as a large title (e.g. "June 2026 · current release period & onward"), the
plain-language scope sentence, and a severity donut of the impacts. Across the tabs:
multi-select toggle filters, release dropdown, a **Toggle expiry** dropdown, a **Major /
Minor** release-kind filter (major = April & October), a one-click **Current month only**
toggle (filters to the cutoff month), a one-click **Customer dashboard records** toggle
(current month **and** specific-scope — reproduces exactly the set the customer dashboard
shows, since the dashboard hides general/program-level matches), text search over full
Overview + Detailed text, and per-tab "Expand … by default" toggles. The hero also shows
the parsed **customer** and **tenant** as identity tags. Each tab shows a **live count bar**
("Showing N of M …") that updates as filters change. Lists are sorted by release version
ascending (earliest first); the release version is a prominent badge (purple for major),
followed by a cyan **RN** badge and green **KB** badges.

1. **Impact** — only items touching the customer's APIs. Filters: Severity, Type, Release
   kind, Current-month, Component type; two display toggles (Expand details / Expand affected
   components). Each card shows the release badge, scope badge, programs/transactions (exact
   in red), an always-visible **Impact summary** (authored "Changed / Impact" digest), full
   Overview + Detailed behind the Details toggle, and an **Affected components** grid (1/2/3
   columns by count).
2. **By Component** — impacts pivoted per component, each listing every release change that
   hits it (deduped by release), ordered by release ascending.
3. **Release Landscape** — every in-scope item grouped into collapsible modules; Current-month
   toggle available here too.
4. **Opportunities** — adoptable enhancements.

### 2. Customer dashboard (`--dashboard`, else `<out>_customer_dashboard.html`)
A clean, **light-themed, static** snapshot to share with the customer, focused on the
**current month's** releases only (release year.month == cutoff month). To keep it credible
and uncluttered it lists **only specific-scope impacts** — items that name an exact
API/transaction the customer calls; vague program-level ("general") matches are excluded.
It states whether the month is a major (Apr/Oct) or minor release, then shows KPI tiles
(release items this month, how many affect the customer, components affected, defects,
enhancements), an "X of Y items touch your applications" line, donut charts (exposure by
severity, change type, and module focus), and a clean card per affecting item: release
badge, **KB number(s)**, severity, the Impact summary, and a **collapsed-by-default**
`<details>` block listing the actual impacted components and the transactions they call
(exact matches highlighted), plus toggle expiry. Customer-facing text reads "Infor
CloudSuite" rather than "M3" (component names that literally contain "M3" are left intact).
Suppress with `--no-dashboard`.

The dashboard reuses the same authored summaries, so run the two-pass summary flow first;
the customer cards lead with the "Changed / Impact" digest.

## Notes

- Always emit the full Overview of Change **and** Detailed Description text — never
  truncate; users rely on the complete "Previously… / After this correction… / ADDITIONAL
  REFERENCE" narrative. The report renders them with `white-space: pre-wrap` so Infor's
  line breaks and `KB ID: / SOLUTION:` structure survive.
- The report is read-only and offline; it embeds its data as JSON, so it can be archived
  and reopened without the source workbooks.
- The stated "Potential Business/Technical Process Impact" fields in the release report
  are frequently "No impact expected" — severity in this report is derived from issue type
  and exact-call matching, not from Infor's own rating. Say so when summarising.
