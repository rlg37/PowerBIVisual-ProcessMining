# Process Mining Flow — Power BI Custom Visual

A certification-ready Power BI custom visual that renders a sequential process flow diagram
with stage nodes, duration edges, and an optional expand-by breakdown view.

**Visual GUID:** `ProcessMiningVisual1A2B3C4D5E6F7890`
**API version:** 5.3.0
**pbiviz tools:** 7.0.3

---

## Field Wells

| Well | Role | Required | Example field |
|------|------|----------|---------------|
| **Stage** | Grouping (categorical) | Yes | `ava_stageindex` |
| **Count** | Measure | No | `Count of Name` |
| **Duration Measure** | Numeric measure | No | `ava_dayswaiting` |
| **Expand By** | Grouping (categorical) | No | `Owner` |

The visual shows a landing page ("Add Stage and Duration Measure fields to begin") until
Stage is mapped. If Expand By is mapped without Stage, no error is thrown — the landing
page is shown consistently.

Stages are sorted numerically when all labels parse as integers, otherwise alphabetically.
Edges between consecutive stages show the average duration measure value.

---

## Toolbar Controls

| Button | Action |
|--------|--------|
| **−** / **+** | Zoom out / in (40 % – 200 %, step 20 %) |
| **⇅ Vertical** / **⇆ Horizontal** | Toggle layout orientation |
| **⊞ Expand** | Toggle expand-by view (enabled only when Expand By field is mapped) |

The Expand toggle state is persisted via `host.persistProperties()` so format-pane
changes do not reset it.

---

## Format Pane

### Layout
| Setting | Default | Description |
|---------|---------|-------------|
| Node Spacing | 80 | Gap (px) between stage blocks |
| Expand Owners | Off | Mirror of the toolbar Expand toggle |

### Stage Nodes
| Setting | Default |
|---------|---------|
| Fill Color | `#0078D4` |
| Border Color | `#004578` |
| Font Color | `#FFFFFF` |
| Font Size (pt) | 12 |
| Bold | On |
| Corner Radius | 6 |

Fill Color, Border Color, and Font Color support the **fx** conditional-formatting button
for per-stage rule-based coloring.

### Expand Nodes
| Setting | Default |
|---------|---------|
| Fill Color | `#005A9E` |
| Border Color | `#003966` |
| Font Color | `#FFFFFF` |
| Font Size (pt) | 11 |
| Use Data Colors | Off |

When Use Data Colors is Off (default), all expand rectangles share the same Fill Color.
When On, each expand value is assigned a color from the Power BI palette.

### Edges & Arrows
| Setting | Default |
|---------|---------|
| Arrow Color | `#605E5C` |
| Label Font Color | `#201F1E` |
| Label Background | `#F3F2F1` |
| Label Font Size (pt) | 11 |
| Duration Unit | `days` |

Arrow Color, Label Font Color, and Label Background support the **fx** button.

### Selection
| Setting | Default |
|---------|---------|
| Highlight Color | `#FFB900` |
| Dim Opacity % | 30 |

---

## Interactions

- **Single click** on a stage node or expand rect → cross-filter other visuals
- **Ctrl/Cmd + click** → additive multi-select
- **Click empty canvas** → clear selection
- **Right-click** → Power BI context menu (drill-through, spotlight, etc.)
- **Hover** → native tooltip showing label, count, and avg duration

In expand mode, clicking a stage container selects all of its child expand IDs at once.

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

The packaged `.pbiviz` file will be output to `dist/`.

### Dev server (live reload in Power BI Desktop)
```bash
pbiviz start
```
In Power BI Desktop → Insert → More Visuals → Developer Visual.

### Certification audit
```bash
pbiviz package --certification-audit
```

Expected result: `No external requests found in the visual.`

---

## AppSource Certification Checklist

### Source Code Review

| # | Check | Status |
|---|-------|--------|
| 1 | No `innerHTML` or `d3.html()` usage | PASS |
| 2 | No `eval()` or `new Function()` | PASS |
| 3 | No `XMLHttpRequest`, `fetch()`, or `WebSocket` in source | PASS |
| 4 | No `setTimeout` / `setInterval` / `requestAnimationFrame` with user input | PASS |
| 5 | `renderingStarted` called as first line of `update()` | PASS |
| 6 | `renderingFinished` called on success path | PASS |
| 7 | `renderingFailed` called in catch block | PASS |
| 8 | All user data written to DOM via `.text()` only — no `innerHTML` | PASS |
| 9 | No minified `.js` files in `src/` | PASS |
| 10 | All dependencies are public OSS — no commercial/private packages | PASS |

### File Inventory

| # | Check | Status |
|---|-------|--------|
| 1 | `pbiviz.json` present and valid | PASS |
| 2 | GUID `ProcessMiningVisual1A2B3C4D5E6F7890` — must never change | PASS |
| 3 | `capabilities.json` — `"privileges": []` (no webAccess) | PASS |
| 4 | `package.json` has `"eslint"` script | PASS |
| 5 | `package-lock.json` present and in sync with `package.json` | PASS |
| 6 | `tsconfig.json` present | PASS |
| 7 | `.gitignore` excludes `node_modules/`, `.tmp/`, `dist/`, `*.pbiviz` | PASS |
| 8 | `assets/icon.png` present and 20×20 px | PASS |
| 9 | `dist/` and `node_modules/` not committed to repo | PASS |

### Certification Audit

| # | Check | Status |
|---|-------|--------|
| 1 | `pbiviz package --certification-audit` — no external requests | PASS |
| 2 | Full `d3` replaced with `d3-selection` only (removes `d3-fetch`) | PASS |

---

## Architecture

```
src/
  visual.ts      — IVisual implementation: toolbar, rendering, selection, highlight
  settings.ts    — Typed format-pane settings (DataViewObjectsParser subclass)
style/
  visual.less    — Minimal base CSS (all layout via D3/SVG)
capabilities.json  — Field wells, format objects, privileges
pbiviz.json        — Visual metadata and API version
assets/
  icon.png       — 20×20 px submission icon
```

### Data flow

```
DataView (categorical)
  ├─ categories[0]  → stage labels  (StageNode[])
  ├─ categories[1]  → expand labels (ExpandNode[] per StageNode) — optional
  ├─ values[count]  → count measure per row
  └─ values[measure]→ duration measure per row
          ↓
  buildViewModel()
  ├─ groups rows by stage key → StageNode[]
  ├─ resolves conditional-format colors from category.objects[rowIdx]
  └─ builds StageEdge[] between consecutive nodes
          ↓
  renderFromState()
  ├─ SVG rect + text per StageNode (or expand children)
  ├─ SVG line + label per StageEdge
  └─ ISelectionManager → cross-filter on node click
```

### Format-pane property storage

`nodeFormat`, `expandFormat`, and `edgeFormat` use a **wildcard selector** with
`ConstantOrRule` color properties to enable the fx conditional-formatting button.
A side effect is that Power BI stores *all* property values (colors and scalars) in
`category.objects` rather than `metadata.objects`. The custom `VisualSettings.parse()`
override reads them back from `category.objects[0]` using `catFill`, `catNum`,
`catBool`, and `catStr` helpers so that all nodes and edges reflect user changes.

`selectionFormat` and `layout` use `selector: null` and are read correctly by the
base `DataViewObjectsParser` without any override.

---

## Known Limitations

- **Format Pane API** — Microsoft will require migration to the new Format Pane API
  (`getFormattingModel`) in a future Power BI release. The current `enumerateObjectInstances`
  API is still supported but flagged as a future-required update.
- **Zoom level** — Zoom is session-only; it resets to 100 % when the report is reloaded.
- **Max rows** — Capped at 500 rows (no-expand mapping) and 2 000 rows (expand mapping)
  via `dataReductionAlgorithm.top` in `capabilities.json`.
