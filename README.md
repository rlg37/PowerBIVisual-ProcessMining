# Process Mining Flow — Power BI Custom Visual

A custom Power BI visual that renders sequential process flow diagrams with stage nodes,
duration edges, and an optional drill-down breakdown view. Designed for **process mining**
use cases where you want to visualize how cases (tickets, orders, patients, etc.) move
through a defined sequence of stages and understand the average time spent between each step.

**Visual GUID:** `ProcessMiningVisual1A2B3C4D5E6F7890`  
**API version:** 5.3.0  
**pbiviz tools:** 7.0.3  
**License:** MIT

---

## What This Visual Does

Process mining is the practice of reconstructing and analyzing real-world process execution
from event log data. This visual consumes that data and draws it as a **linear flow diagram**:

- Each **stage node** represents a step in the process (e.g., `New → In Progress → Resolved`)
- Each **edge arrow** between stages shows the **average duration** of time cases spent
  transitioning between them
- An optional **Expand By** field breaks each stage node into sub-nodes grouped by a
  categorical attribute (e.g., Owner, Department, Region) to reveal which segments drive
  volume or delay
- **Cross-filtering** lets you click any node to filter all other visuals on the page to
  the cases that passed through that stage

Typical data sources include CRM case records, ITSM ticket systems, ERP workflows,
manufacturing event logs, or any tabular dataset with a stage/status column and a
date-based duration measure.

---

## Field Wells

| Well | Role | Required | Example field |
|------|------|----------|---------------|
| **Stage** | Grouping (categorical) | Yes | `ava_stageindex`, `Status` |
| **Count** | Measure | No | `Count of Cases` |
| **Duration Measure** | Numeric measure | No | `Avg Days Waiting` |
| **Expand By** | Grouping (categorical) | No | `Owner`, `Team`, `Region` |

**Stage** is the only required field. Until it is mapped the visual displays a landing page:
> *"Add Stage and Duration Measure fields to begin"*

Stages are sorted **numerically** when all labels parse as integers, otherwise **alphabetically**.
Edges display the **average** of the Duration Measure for all rows between consecutive stages.

---

## How It Works

### Data Processing

```
Power BI DataView (categorical)
  ├─ categories[0]   → stage labels            → StageNode[]
  ├─ categories[1]   → expand-by labels        → ExpandNode[] per StageNode  (optional)
  ├─ values[count]   → count measure per row
  └─ values[measure] → duration measure per row
          ↓
  buildViewModel()
  ├─ Groups rows by stage key → StageNode[] with aggregated count and duration
  ├─ Resolves per-stage conditional-format colors from category.objects[rowIdx]
  └─ Builds StageEdge[] between consecutive nodes with avg duration labels
          ↓
  renderFromState()
  ├─ SVG rect + text per StageNode (or expand children in expand mode)
  ├─ SVG line + arrowhead + label per StageEdge
  └─ ISelectionManager → cross-filters other visuals on node click
```

### Layout

The diagram renders in either **vertical** (top-to-bottom) or **horizontal** (left-to-right)
orientation. Zoom is adjustable from 40% to 200% in 20% steps via the toolbar. Node spacing
(the gap between stage blocks) is configurable in the Format pane.

When **Expand By** is active, each stage container expands vertically to show one sub-node
per unique value of the expand field. Stage containers select all of their children's
selection IDs at once; expand children can be clicked individually.

---

## Toolbar Controls

| Button | Action |
|--------|--------|
| **−** / **+** | Zoom out / in (40%–200%, step 20%) |
| **⇅ Vertical** / **⇆ Horizontal** | Toggle layout orientation |
| **⊞ Expand** | Toggle expand-by view (enabled only when Expand By is mapped) |

The Expand toggle state is persisted via `host.persistProperties()` so format-pane
changes do not reset it between interactions.

---

## Format Pane

### Layout
| Setting | Default | Description |
|---------|---------|-------------|
| Node Spacing | 80 px | Gap between consecutive stage blocks |
| Expand Owners | Off | Mirror of the toolbar Expand toggle |

### Stage Nodes
| Setting | Default | Conditional Formatting (fx) |
|---------|---------|---|
| Fill Color | `#0078D4` | Yes |
| Border Color | `#004578` | Yes |
| Font Color | `#FFFFFF` | Yes |
| Font Size | 12 pt | No |
| Bold | On | No |
| Corner Radius | 6 px | No |

### Expand Nodes
| Setting | Default | Description |
|---------|---------|-------------|
| Fill Color | `#005A9E` | Background for expand sub-nodes |
| Border Color | `#003966` | Outline for expand sub-nodes |
| Font Color | `#FFFFFF` | Text color |
| Font Size | 11 pt | |
| Use Data Colors | Off | When On, assigns Power BI theme palette colors per expand value |

### Edges & Arrows
| Setting | Default | Conditional Formatting (fx) |
|---------|---------|---|
| Arrow Color | `#605E5C` | Yes |
| Label Font Color | `#201F1E` | Yes |
| Label Background | `#F3F2F1` | Yes |
| Label Font Size | 11 pt | No |
| Duration Unit | `days` | No — suffix shown in hover tooltip |

### Selection
| Setting | Default | Description |
|---------|---------|-------------|
| Highlight Color | `#FFB900` | Stroke color applied to selected nodes |
| Dim Opacity % | 30 | Opacity of unselected nodes while a selection is active |

---

## Interactions

| Action | Result |
|--------|--------|
| **Click** a stage node or expand rect | Cross-filters other visuals to matching cases |
| **Ctrl / Cmd + click** | Additive multi-select across nodes |
| **Click empty canvas** | Clears the current selection |
| **Right-click** | Power BI context menu (drill-through, spotlight, copy value, etc.) |
| **Hover** | Native tooltip showing stage label, count, and average duration |

In expand mode, clicking a stage container selects all of its child expand IDs simultaneously.

---

## Build Instructions

### Prerequisites
```bash
node --version   # >= 16
npm install -g powerbi-visuals-tools@7
```

### Install and build
```bash
npm install
pbiviz package
```

Output: `dist/ProcessMiningVisual.pbiviz`

### Dev server (live reload in Power BI Desktop)
```bash
pbiviz start
```
In Power BI Desktop → **Insert → More Visuals → Developer Visual**.

### Lint (AppSource compliance check)
```bash
npm run eslint
```

### Certification audit
```bash
pbiviz package --certification-audit
```
Expected: `No external requests found in the visual.`

---

## AppSource Certification

### Source Code Review

| # | Check | Status |
|---|-------|--------|
| 1 | No `innerHTML` or `d3.html()` usage | PASS |
| 2 | No `eval()` or `new Function()` | PASS |
| 3 | No `XMLHttpRequest`, `fetch()`, or `WebSocket` | PASS |
| 4 | No `setTimeout` / `setInterval` with user input | PASS |
| 5 | `renderingStarted` called as first line of `update()` | PASS |
| 6 | `renderingFinished` called on success path | PASS |
| 7 | `renderingFailed` called in catch block | PASS |
| 8 | All user data written to DOM via `.text()` — no `innerHTML` | PASS |
| 9 | No minified `.js` in `src/` | PASS |
| 10 | All dependencies are public OSS | PASS |

### File Inventory

| # | Check | Status |
|---|-------|--------|
| 1 | `pbiviz.json` present and valid | PASS |
| 2 | GUID `ProcessMiningVisual1A2B3C4D5E6F7890` — never changed | PASS |
| 3 | `capabilities.json` — `"privileges": []` (no webAccess) | PASS |
| 4 | `package.json` includes `"eslint"` script | PASS |
| 5 | `package-lock.json` in sync with `package.json` | PASS |
| 6 | `tsconfig.json` present | PASS |
| 7 | `.gitignore` excludes `node_modules/`, `.tmp/`, `dist/`, `*.pbiviz` | PASS |
| 8 | `assets/icon.png` present and 20×20 px | PASS |
| 9 | `dist/` and `node_modules/` not committed | PASS |

### Certification Audit

| # | Check | Status |
|---|-------|--------|
| 1 | `pbiviz package --certification-audit` — no external requests | PASS |
| 2 | Full `d3` replaced with `d3-selection` only (excludes `d3-fetch`) | PASS |

---

## Architecture

```
PowerBIVisual-ProcessMining/
├── src/
│   ├── visual.ts      — IVisual implementation: data model, SVG rendering,
│   │                    toolbar, selection, cross-filter, highlight handling
│   └── settings.ts    — Typed format-pane settings (DataViewObjectsParser subclass)
│                        with custom parse() for wildcard-selector color properties
├── style/
│   └── visual.less    — Minimal base CSS (layout handled entirely via D3/SVG)
├── capabilities.json  — Field wells, two categorical data-view mappings,
│                        format objects with wildcard selectors, privileges: []
├── pbiviz.json        — Visual metadata, GUID, API version, author
└── assets/
    └── icon.png       — 20×20 px submission icon
```

### Format-pane property storage

`nodeFormat`, `expandFormat`, and `edgeFormat` use a **wildcard selector** with
`ConstantOrRule` color properties, which enables the **fx** conditional-formatting button
per stage. A side effect is that Power BI stores *all* property values in
`category.objects` rather than `metadata.objects`. The custom `VisualSettings.parse()`
override reads them back from `category.objects[0]` via `catFill`, `catNum`, `catBool`,
and `catStr` helpers so that global (non-conditional) settings work correctly.

`selectionFormat` and `layout` use `selector: null` and are read by the standard
`DataViewObjectsParser` base without override.

---

## Known Limitations

- **Format Pane API** — Currently uses `enumerateObjectInstances()`. Microsoft plans to
  require migration to `getFormattingModel()` in a future release; existing API still
  supported as of Power BI API 5.x.
- **Zoom** — Zoom level is session-only and resets to 100% on report reload.
- **Row caps** — 500 rows (no-expand mapping) and 2,000 rows (expand mapping), set via
  `dataReductionAlgorithm.top` in `capabilities.json`.
