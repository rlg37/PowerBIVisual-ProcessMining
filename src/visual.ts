"use strict";


import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions       = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                   = powerbi.extensibility.visual.IVisual;
import IVisualHost               = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager         = powerbi.extensibility.ISelectionManager;
import ISelectionId              = powerbi.visuals.ISelectionId;
import IVisualEventService       = powerbi.extensibility.IVisualEventService;
import DataView                  = powerbi.DataView;
import DataViewCategorical       = powerbi.DataViewCategorical;
import VisualUpdateType          = powerbi.VisualUpdateType;

import { VisualSettings } from "./settings";
import * as d3 from "d3-selection";
import { dataViewWildcard } from "powerbi-visuals-utils-dataviewutils";

// ─────────────────────────────────────────────────────────────────────────────
// Interfaces
// ─────────────────────────────────────────────────────────────────────────────

interface ExpandNode {
    label:       string;
    count:       number;
    measureAgg:  number;
    selectionId: ISelectionId;
    color:       string;
}

interface StageNode {
    label:       string;
    count:       number;
    measureAgg:  number;
    index:       number;
    firstRowIdx: number;   // first dataView row index for this stage (used for CF lookup)
    selectionId: ISelectionId;
    expandNodes: ExpandNode[];
    highlighted: boolean | null;
    // Conditional-format resolved colors (null = use global setting)
    cfNodeFill:   string | null;
    cfNodeBorder: string | null;
    cfNodeFont:   string | null;
}

interface StageEdge {
    fromIndex:     number;
    toIndex:       number;
    durationValue: number;
    // Conditional-format resolved colors (null = use global setting)
    cfArrowColor:  string | null;
    cfLabelFont:   string | null;
    cfLabelBg:     string | null;
}

interface ViewModel {
    nodes:       StageNode[];
    edges:       StageEdge[];
    hasData:     boolean;
    hasExpand:   boolean;
    measureName: string;
    isTableMode: boolean;
}

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

const NODE_W     = 160;
const NODE_H     = 56;
const EXPAND_W   = 140;
const EXPAND_H   = 48;
const EXPAND_GAP = 8;
const ARROW_SZ   = 8;
const PAD        = 32;
const TOOLBAR_H  = 34;
const ZOOM_STEP  = 20;
const ZOOM_MIN   = 40;
const ZOOM_MAX   = 200;

// ─────────────────────────────────────────────────────────────────────────────
// Visual
// ─────────────────────────────────────────────────────────────────────────────

export class Visual implements IVisual {
    private host:             IVisualHost;
    private events:           IVisualEventService;
    private selectionManager: ISelectionManager;
    private rootDiv:          d3.Selection<HTMLDivElement, unknown, null, undefined>;
    private toolbar:          d3.Selection<HTMLDivElement, unknown, null, undefined>;
    private scrollDiv:        d3.Selection<HTMLDivElement, unknown, null, undefined>;
    private svgWrap:          d3.Selection<SVGSVGElement,  unknown, null, undefined>;
    private diagramG:         d3.Selection<SVGGElement,    unknown, null, undefined>;
    private settings:         VisualSettings;
    private viewModel:        ViewModel;
    private lastViewport:          powerbi.IViewport = { width: 400, height: 300 };
    private lastViewPort_dataView: DataView | undefined;

    // UI state — only synced from format pane on Data updates, not on every update()
    private orientation:      string  = "vertical";
    private expandMode:       boolean = false;
    private zoomLevel:        number  = 100;
    // FIX: track whether interactions are allowed independently of isInFocus
    private allowInteractions: boolean = true;

    constructor(options: VisualConstructorOptions) {
        this.host             = options.host;
        this.events           = options.host.eventService;
        this.selectionManager = this.host.createSelectionManager();

        this.rootDiv = d3.select(options.element)
            .append("div").classed("pm-root", true)
            .style("display", "flex").style("flex-direction", "column")
            .style("width", "100%").style("height", "100%")
            .style("box-sizing", "border-box");

        // ── Toolbar ───────────────────────────────────────────────────────────
        this.toolbar = this.rootDiv.append("div").classed("pm-toolbar", true)
            .style("display", "flex").style("align-items", "center")
            .style("justify-content", "flex-end").style("gap", "4px")
            .style("padding", "3px 8px").style("flex-shrink", "0")
            .style("height", `${TOOLBAR_H}px`)
            .style("border-bottom", "1px solid #EDEBE9");

        this.toolbar.append("button")
            .classed("pm-btn pm-btn-zoom-out", true).attr("title", "Zoom out").text("−")
            .on("click", () => {
                this.zoomLevel = Math.max(ZOOM_MIN, this.zoomLevel - ZOOM_STEP);
                this.updateToolbarState();
                this.renderFromState(this.lastViewport);
            });

        this.toolbar.append("span").classed("pm-zoom-label", true)
            .style("font-size", "11px").style("font-family", "Segoe UI, sans-serif")
            .style("color", "#605E5C").style("min-width", "34px")
            .style("text-align", "center").text("100%");

        this.toolbar.append("button")
            .classed("pm-btn pm-btn-zoom-in", true).attr("title", "Zoom in").text("+")
            .on("click", () => {
                this.zoomLevel = Math.min(ZOOM_MAX, this.zoomLevel + ZOOM_STEP);
                this.updateToolbarState();
                this.renderFromState(this.lastViewport);
            });

        this.toolbar.append("div")
            .style("width", "1px").style("height", "18px")
            .style("background", "#EDEBE9").style("margin", "0 2px");

        this.toolbar.append("button")
            .classed("pm-btn pm-btn-orient", true)
            .attr("title", "Toggle vertical / horizontal layout").text("⇅ Vertical")
            .on("click", () => {
                this.orientation = this.orientation === "vertical" ? "horizontal" : "vertical";
                this.updateToolbarState();
                this.renderFromState(this.lastViewport);
            });

        this.toolbar.append("button")
            .classed("pm-btn pm-btn-expand", true)
            .attr("title", "Toggle expand view").text("⊞ Expand")
            .on("click", () => {
                if (!this.viewModel?.hasExpand) return;
                this.expandMode = !this.expandMode;
                // Persist so format-pane updates don't reset the toggle state
                this.host.persistProperties({
                    merge: [{
                        objectName: "layout",
                        selector: null,
                        properties: { expandMode: this.expandMode }
                    }]
                });
                this.updateToolbarState();
                this.renderFromState(this.lastViewport);
            });

        this.scrollDiv = this.rootDiv.append("div").classed("pm-scroll", true)
            .style("flex", "1 1 auto").style("overflow", "auto")
            .style("box-sizing", "border-box");

        this.svgWrap  = this.scrollDiv.append("svg").classed("pm-svg", true).style("display", "block");
        this.diagramG = this.svgWrap.append("g").classed("pm-diagram", true);

        this.svgWrap.on("click", () => {
            // FIX: always allow clearing selection — do not gate on allowInteractions
            // Power BI's allowInteractions only gates cross-visual filtering, not internal UX
            this.selectionManager.clear().then(() => this.paintHighlight([], false));
        });
    }

    // ── update ────────────────────────────────────────────────────────────────
    public update(options: VisualUpdateOptions): void {
        this.events.renderingStarted(options);
        try {
            const dataView = options?.dataViews?.[0];
            // Guard: parse() can throw if dataView is undefined (e.g. only Expand By
            // is mapped but Stage is absent — neither dataViewMapping condition is met)
            this.settings = dataView
                ? VisualSettings.parse<VisualSettings>(dataView)
                : new VisualSettings();
            this.lastViewport          = options.viewport;
            this.lastViewPort_dataView = dataView;

            // FIX: correct property for cross-visual interactions permission
            // isInFocus is focus/expand mode, not the interactions flag
            this.allowInteractions =
                (options as any).interactivityOptions?.allowInteractions !== false;

            if (options.type === VisualUpdateType.All ||
                options.type === VisualUpdateType.Data) {
                this.orientation = this.settings.layout.orientation;
                this.expandMode  = this.settings.layout.expandMode;
                this.zoomLevel   = Math.max(ZOOM_MIN, Math.min(ZOOM_MAX,
                    this.settings.layout.zoomLevel ?? 100));
                this.viewModel   = this.buildViewModel(dataView);
            }

            this.updateToolbarState();
            this.processInboundHighlights(dataView);
            this.renderFromState(options.viewport);
        } catch (e) {
            this.events.renderingFailed(options, String(e));
            return;
        }
        this.events.renderingFinished(options);
    }

    // ── enumerateObjectInstances ──────────────────────────────────────────────
    // Drives the Visual tab in the format pane.
    // This API is called by Desktop AFTER update() has run with the current
    // dataView, so this.settings always reflects the latest user values.
    //
    // Conditional formatting (fx button) support:
    // The correct pattern (per Microsoft docs) is to emit one instance per
    // data point using:
    //   selector:              dataViewWildcard wildcard selector (InstancesAndTotals)
    //   altConstantValueSelector: the data point's own selectionId selector
    //   propertyInstanceKind:  { propName: VisualEnumerationInstanceKinds.ConstantOrRule }
    //
    // IMPORTANT — color swatch accuracy:
    // When a constant color is set via the plain picker (not fx), Power BI stores
    // the value in category.objects[0][objectName][propName], NOT in
    // metadata.objects.  DataViewObjectsParser only reads metadata.objects, so
    // this.settings will always show the class default (e.g. #0078D4) for color
    // properties that use a wildcard selector.  We must read the saved value back
    // directly from category.objects here so the swatch reflects the real color.
    public enumerateObjectInstances(
        options: powerbi.EnumerateVisualObjectInstancesOptions
    ): powerbi.VisualObjectInstanceEnumeration {
        const s   = this.settings ?? new VisualSettings();
        const obj = options.objectName;

        // Shorthand helpers
        const fill   = (color: string) => ({ solid: { color } });
        const COrR   = powerbi.VisualEnumerationInstanceKinds.ConstantOrRule;
        const wcSel  = () => dataViewWildcard.createDataViewWildcardSelector(
            dataViewWildcard.DataViewWildcardMatchingOption.InstancesAndTotals);

        // Retrieve the stage / expand category columns
        const dvCat     = this.lastViewPort_dataView?.categorical as DataViewCategorical | undefined;
        const stageCol  = dvCat?.categories?.find(c => c.source?.roles?.["stage"]);
        const expandCol = dvCat?.categories?.find(c => c.source?.roles?.["expand"]);

        const stageAltSel  = stageCol?.identity?.[0]
            ? { data: [stageCol.identity[0]] } : null;
        const expandAltSel = expandCol?.identity?.[0]
            ? { data: [expandCol.identity[0]] } : stageAltSel;

        // Read a color from category.objects[rowIdx], falling back to `fallback`.
        // This is necessary because constant colors saved via the plain color picker
        // are stored in category.objects when a wildcard selector is used — not in
        // metadata.objects — so DataViewObjectsParser never sees them.
        const catColor = (
            col: powerbi.DataViewCategoryColumn | undefined,
            rowIdx: number,
            objectName: string,
            propName: string,
            fallback: string
        ): string => {
            if (!col?.objects) return fallback;
            const o = col.objects[rowIdx];
            if (!o) return fallback;
            const p = (o as any)[objectName]?.[propName];
            if (!p) return fallback;
            if (typeof p === "object" && p.solid?.color) return p.solid.color;
            if (typeof p === "string" && /^#[0-9A-Fa-f]{6}$/i.test(p.trim())) return p.trim();
            return fallback;
        };

        switch (obj) {
            case "layout":
                return [{ objectName: obj, properties: {
                    nodeSpacing: s.layout.nodeSpacing,
                    expandMode:  s.layout.expandMode
                }, selector: null }];

            case "nodeFormat":
                return [{
                    objectName: obj,
                    properties: {
                        // Read color back from category.objects so swatch shows the
                        // actual saved value, not the class default from settings.
                        fillColor:    fill(catColor(stageCol, 0, obj, "fillColor",   s.nodeFormat.fillColor)),
                        borderColor:  fill(catColor(stageCol, 0, obj, "borderColor", s.nodeFormat.borderColor)),
                        fontColor:    fill(catColor(stageCol, 0, obj, "fontColor",   s.nodeFormat.fontColor)),
                        fontSize:     s.nodeFormat.fontSize,
                        fontBold:     s.nodeFormat.fontBold,
                        borderRadius: s.nodeFormat.borderRadius
                    },
                    selector: stageAltSel ? wcSel() : null,
                    altConstantValueSelector: stageAltSel ?? undefined,
                    propertyInstanceKind: {
                        fillColor:   COrR,
                        borderColor: COrR,
                        fontColor:   COrR
                    }
                }];

            case "expandFormat": {
                const col = expandCol ?? stageCol;
                return [{
                    objectName: obj,
                    properties: {
                        fillColor:     fill(catColor(col, 0, obj, "fillColor",   s.expandFormat.fillColor)),
                        borderColor:   fill(catColor(col, 0, obj, "borderColor", s.expandFormat.borderColor)),
                        fontColor:     fill(catColor(col, 0, obj, "fontColor",   s.expandFormat.fontColor)),
                        fontSize:      s.expandFormat.fontSize,
                        useDataColors: s.expandFormat.useDataColors
                    },
                    selector: expandAltSel ? wcSel() : null,
                    altConstantValueSelector: expandAltSel ?? undefined,
                    propertyInstanceKind: {
                        fillColor:   COrR,
                        borderColor: COrR,
                        fontColor:   COrR
                    }
                }];
            }

            case "edgeFormat":
                return [{
                    objectName: obj,
                    properties: {
                        arrowColor:     fill(catColor(stageCol, 0, obj, "arrowColor",     s.edgeFormat.arrowColor)),
                        labelFontColor: fill(catColor(stageCol, 0, obj, "labelFontColor", s.edgeFormat.labelFontColor)),
                        labelBgColor:   fill(catColor(stageCol, 0, obj, "labelBgColor",   s.edgeFormat.labelBgColor)),
                        labelFontSize:  s.edgeFormat.labelFontSize,
                        measureUnit:    s.edgeFormat.measureUnit
                    },
                    selector: stageAltSel ? wcSel() : null,
                    altConstantValueSelector: stageAltSel ?? undefined,
                    propertyInstanceKind: {
                        arrowColor:     COrR,
                        labelFontColor: COrR,
                        labelBgColor:   COrR
                    }
                }];

            case "selectionFormat":
                return [{ objectName: obj, properties: {
                    highlightColor: fill(s.selectionFormat.highlightColor),
                    dimOpacity:     s.selectionFormat.dimOpacity
                }, selector: null }];

            default:
                return [];
        }
    }

    // ── buildViewModel ────────────────────────────────────────────────────────
    private buildViewModel(dataView: DataView): ViewModel {
        const empty: ViewModel = {
            nodes: [], edges: [], hasData: false,
            hasExpand: false, measureName: "", isTableMode: false
        };
        if (!dataView) return empty;

        // Power BI selects the dataViewMapping that matches the current fields.
        // When expand is present the table mapping should win (conditions: expand min:1).
        // When expand is absent the categorical mapping wins (conditions: expand max:0).
        // Both the no-expand and expand mappings now produce categorical dataViews.
        // When expand is present: categories[0] = stage, categories[1] = expand.
        // When expand is absent: categories[0] = stage only.
        if (dataView.categorical) return this.buildFromCategorical(dataView);
        return empty;
    }

    // ── Categorical (handles both no-expand and expand-by modes) ─────────────
    // With no Expand By field:  categories[0] = stage
    // With Expand By field:     categories[0] = stage, categories[1] = expand
    //
    // Both cases arrive as categorical because the table mapping's "select" syntax
    // is not valid in the Power BI custom visuals API (only "for"/"bind" are valid
    // inside table.rows).  Using categorical with two category columns is the
    // correct approach for multi-grouping.
    private buildFromCategorical(dataView: DataView): ViewModel {
        const cat = dataView.categorical as DataViewCategorical;

        // Find stage and optional expand columns by role, not by position,
        // so the mapping order in capabilities.json doesn't matter.
        const stageCol  = cat?.categories?.find(c => c.source?.roles?.["stage"]);
        const expandCol = cat?.categories?.find(c => c.source?.roles?.["expand"]);

        if (!stageCol?.values?.length) return {
            nodes: [], edges: [], hasData: false, hasExpand: false, measureName: "", isTableMode: false
        };

        const countCol   = cat.values?.find(v => v.source?.roles?.["count"]);
        const measureCol = cat.values?.find(v => v.source?.roles?.["measure"]);
        const hasExpand  = !!expandCol;
        const numRows    = stageCol.values.length;
        const palette    = this.host.colorPalette;

        // ── Categorical object color reader ───────────────────────────────────
        // When the user sets a conditional format rule via the fx button, Power BI
        // evaluates it and writes the resolved color into
        //   category.objects[rowIdx][objectName][propName]
        // as a Fill value { solid: { color: "#..." } }.
        // This helper reads that value, returning null when absent so callers can
        // fall back to the global setting from this.settings.
        const getCatObjColor = (
            col: powerbi.DataViewCategoryColumn,
            rowIdx: number,
            objectName: string,
            propName: string
        ): string | null => {
            const obj = col.objects?.[rowIdx];
            if (!obj) return null;
            const prop = (obj as any)[objectName]?.[propName];
            if (!prop) return null;
            if (typeof prop === "object" && prop.solid?.color) return prop.solid.color;
            if (typeof prop === "string" && /^#[0-9A-Fa-f]{6}$/i.test(prop.trim())) return prop.trim();
            return null;
        };

        interface EA { cs: number; ms: number; rc: number; fr: number; }
        interface SA { cs: number; ms: number; rc: number; fr: number; em: Map<string, EA>; }
        const sm = new Map<string, SA>();

        for (let i = 0; i < numRows; i++) {
            const sk = stageCol.values[i] == null ? "(blank)" : String(stageCol.values[i]);
            if (!sm.has(sk)) sm.set(sk, { cs: 0, ms: 0, rc: 0, fr: i, em: new Map() });
            const sa = sm.get(sk)!;
            sa.ms += measureCol ? (Number(measureCol.values[i]) || 0) : 0;
            sa.cs += countCol   ? (Number(countCol.values[i])   || 0) : 0;
            sa.rc++;

            if (hasExpand) {
                const ek = expandCol!.values[i] == null ? "(blank)" : String(expandCol!.values[i]);
                if (!sa.em.has(ek)) sa.em.set(ek, { cs: 0, ms: 0, rc: 0, fr: i });
                const ea = sa.em.get(ek)!;
                ea.ms += measureCol ? (Number(measureCol.values[i]) || 0) : 0;
                ea.cs += countCol   ? (Number(countCol.values[i])   || 0) : 0;
                ea.rc++;
            }
        }

        const entries = Array.from(sm.entries());
        const allNum  = entries.every(([k]) => !isNaN(Number(k)));
        entries.sort((a, b) => allNum ? Number(a[0]) - Number(b[0]) : a[0].localeCompare(b[0]));

        const nodes: StageNode[] = entries.map(([label, sa], idx) => {
            const expandNodes: ExpandNode[] = hasExpand
                ? Array.from(sa.em.entries())
                    .sort((a, b) => a[0].localeCompare(b[0]))
                    .map(([ek, ea]) => ({
                        label:      ek,
                        count:      countCol ? ea.cs : ea.rc,
                        measureAgg: ea.rc > 0 ? ea.ms / ea.rc : 0,
                        // Single-category selection on the expand column only.
                        // Chaining withCategory(stageCol) + withCategory(expandCol)
                        // produces a compound predicate that Dataverse TDS cannot
                        // translate, causing rsDataShapeQueryTranslationError.
                        selectionId: this.host.createSelectionIdBuilder()
                            .withCategory(expandCol!, ea.fr)
                            .createSelectionId(),
                        color: palette.getColor(ek).value
                    }))
                : [];

            // When expand is present, stage nodes do NOT emit their own selection ID.
            // Clicking a stage node multi-selects all child expand IDs instead
            // (handled in renderFromState). This avoids a bare stage-only filter
            // predicate that Dataverse rejects on a cross-joined categorical query.
            const stageSelId = !hasExpand
                ? this.host.createSelectionIdBuilder()
                    .withCategory(stageCol, sa.fr).createSelectionId()
                : null;

            return {
                label,
                count:       countCol ? sa.cs : sa.rc,
                measureAgg:  sa.rc > 0 ? sa.ms / sa.rc : 0,
                index:       idx,
                firstRowIdx: sa.fr,
                selectionId: stageSelId!,
                expandNodes,
                highlighted: null,
                // Read conditional-format colors from category.objects at the first
                // row for this stage value.  null = no rule set, use global setting.
                cfNodeFill:   getCatObjColor(stageCol, sa.fr, "nodeFormat",  "fillColor"),
                cfNodeBorder: getCatObjColor(stageCol, sa.fr, "nodeFormat",  "borderColor"),
                cfNodeFont:   getCatObjColor(stageCol, sa.fr, "nodeFormat",  "fontColor")
            };
        });

        return {
            nodes, edges: this.buildEdges(nodes, stageCol, getCatObjColor), hasData: true,
            hasExpand,
            measureName: measureCol?.source?.displayName ?? "Duration",
            isTableMode: false
        };
    }

    private buildEdges(
        nodes: StageNode[],
        stageCol?: powerbi.DataViewCategoryColumn,
        getCatObjColor?: (col: powerbi.DataViewCategoryColumn, rowIdx: number, obj: string, prop: string) => string | null
    ): StageEdge[] {
        const edges: StageEdge[] = [];
        for (let i = 0; i < nodes.length - 1; i++) {
            const rowIdx  = nodes[i].firstRowIdx;
            const cfArrow = (stageCol && getCatObjColor)
                ? getCatObjColor(stageCol, rowIdx, "edgeFormat", "arrowColor") : null;
            const cfLFont = (stageCol && getCatObjColor)
                ? getCatObjColor(stageCol, rowIdx, "edgeFormat", "labelFontColor") : null;
            const cfLBg   = (stageCol && getCatObjColor)
                ? getCatObjColor(stageCol, rowIdx, "edgeFormat", "labelBgColor") : null;
            edges.push({
                fromIndex:     i,
                toIndex:       i + 1,
                durationValue: (nodes[i].measureAgg + nodes[i + 1].measureAgg) / 2,
                cfArrowColor:  cfArrow,
                cfLabelFont:   cfLFont,
                cfLabelBg:     cfLBg
            });
        }
        return edges;
    }

    // ── Inbound highlights ────────────────────────────────────────────────────
    private processInboundHighlights(dataView: DataView): void {
        if (!this.viewModel?.hasData) return;
        const cat = dataView?.categorical as DataViewCategorical | undefined;
        // Find the stage category by role, not by position (categories[1] may now be expand)
        const stageCol = cat?.categories?.find(c => c.source?.roles?.["stage"]);
        if (!stageCol) { this.viewModel.nodes.forEach(n => n.highlighted = null); return; }
        const anyH = cat!.values?.some(v => v.highlights?.some(h => h !== null));
        if (!anyH) { this.viewModel.nodes.forEach(n => n.highlighted = null); return; }
        const hs = new Set<string>();
        stageCol.values.forEach((v, i) => {
            if (cat!.values?.some(col => col.highlights && col.highlights[i] !== null))
                hs.add(v == null ? "(blank)" : String(v));
        });
        this.viewModel.nodes.forEach(n => { n.highlighted = hs.has(n.label); });
    }

    // ── renderFromState ───────────────────────────────────────────────────────
    private renderFromState(viewport: powerbi.IViewport): void {
        this.diagramG.selectAll("*").remove();
        if (!this.viewModel?.hasData) { this.renderLanding(viewport); return; }

        const vm      = this.viewModel;
        const s       = this.settings;
        const isV     = this.orientation !== "horizontal";
        const doExp   = this.expandMode && vm.hasExpand;
        const spacing = Math.max(20, s.layout.nodeSpacing ?? 80);
        const scale   = this.zoomLevel / 100;
        const hc      = this.host.colorPalette.isHighContrast;

        // ── Conditional-format color resolver ────────────────────────────────
        // With the wildcard selector + propertyInstanceKind pattern, Power BI
        // evaluates the conditional formatting rule and writes the resolved color
        // into dataView.metadata.objects[objectName][propName] — the same location
        // that DataViewObjectsParser already reads in VisualSettings.parse().
        // So this.settings already contains the rule-resolved color; no per-row
        // lookup is needed.  The helpers below simply read from this.settings,
        // keeping the render path clean and identical to the non-CF case.
        const dv = this.lastViewPort_dataView; // retained for future use / debugging

        // Global (non-conditional) fallback colors
        const nodeFill   = hc ? this.host.colorPalette.background.value : s.nodeFormat.fillColor;
        const nodeBorder = hc ? this.host.colorPalette.foreground.value : s.nodeFormat.borderColor;
        const nodeFontC  = hc ? this.host.colorPalette.foreground.value : s.nodeFormat.fontColor;
        const arrowC     = hc ? this.host.colorPalette.foreground.value : s.edgeFormat.arrowColor;
        const lblC       = hc ? this.host.colorPalette.foreground.value : s.edgeFormat.labelFontColor;
        const lblBg      = hc ? this.host.colorPalette.background.value : s.edgeFormat.labelBgColor;
        const dimOp      = Math.max(0, Math.min(100, s.selectionFormat.dimOpacity ?? 30)) / 100;

        interface Box { w: number; h: number; }
        const boxes: Box[] = vm.nodes.map(n => {
            if (!doExp || !n.expandNodes.length) return { w: NODE_W, h: NODE_H };
            const cnt = n.expandNodes.length;
            return isV
                ? { w: Math.max(NODE_W, cnt * EXPAND_W + (cnt - 1) * EXPAND_GAP), h: NODE_H }
                : { w: NODE_W, h: Math.max(NODE_H, cnt * EXPAND_H + (cnt - 1) * EXPAND_GAP) };
        });

        const LABEL_OH = doExp ? 22 : 0;
        const crossMax = Math.max(...boxes.map(b => isV ? b.w : b.h));

        // Unscaled content dimensions
        let mainLen = PAD;
        for (let i = 0; i < vm.nodes.length; i++) {
            mainLen += (isV ? boxes[i].h : boxes[i].w) + LABEL_OH + spacing;
        }
        mainLen = mainLen - spacing + PAD;

        const contentW = isV ? (2 * PAD + crossMax) : mainLen;
        const contentH = isV ? mainLen               : (2 * PAD + crossMax + LABEL_OH);

        // Scaled SVG dimensions — always at least as wide/tall as the viewport
        const svgW = Math.max(contentW * scale, viewport.width);
        const svgH = Math.max(contentH * scale, viewport.height - TOOLBAR_H);
        this.svgWrap.attr("width", svgW).attr("height", svgH);

        // FIX: center the diagram within the actual SVG width, not just contentW.
        // When svgW > contentW * scale the diagram must be offset to stay centered.
        const offsetX = isV ? (svgW - contentW * scale) / 2 : 0;
        const offsetY = isV ? 0 : (svgH - contentH * scale) / 2;
        this.diagramG.attr("transform", `translate(${offsetX},${offsetY}) scale(${scale})`);

        // Cross-center in unscaled space (content coords)
        const crossCtr = (isV ? contentW : contentH) / 2;

        const centers: { cx: number; cy: number }[] = [];
        let cursor = PAD + LABEL_OH;
        for (let i = 0; i < vm.nodes.length; i++) {
            const dim = isV ? boxes[i].h : boxes[i].w;
            centers.push(isV
                ? { cx: crossCtr, cy: cursor + dim / 2 }
                : { cx: cursor + dim / 2, cy: crossCtr });
            cursor += dim + spacing + LABEL_OH;
        }

        // Arrowhead in defs (outside diagramG so it's not double-scaled)
        this.svgWrap.select("defs").remove();
        const defs = this.svgWrap.append("defs");
        defs.append("marker").attr("id", "pmArrow")
            .attr("viewBox", "0 -5 10 10").attr("refX", 9).attr("refY", 0)
            .attr("markerWidth", ARROW_SZ).attr("markerHeight", ARROW_SZ).attr("orient", "auto")
            .append("path").attr("d", "M0,-5L10,0L0,5").attr("fill", arrowC);

        // ── Edges ─────────────────────────────────────────────────────────────
        const edgeG = this.diagramG.append("g").classed("pm-edges", true);
        vm.edges.forEach(edge => {
            const f = centers[edge.fromIndex]; const t = centers[edge.toIndex];
            const bF = boxes[edge.fromIndex];  const bT = boxes[edge.toIndex];
            let x1: number, y1: number, x2: number, y2: number;
            if (isV) {
                x1 = f.cx; y1 = f.cy + bF.h / 2;
                x2 = t.cx; y2 = t.cy - bT.h / 2 - ARROW_SZ;
            } else {
                x1 = f.cx + bF.w / 2;            y1 = f.cy;
                x2 = t.cx - bT.w / 2 - ARROW_SZ; y2 = t.cy;
            }

            // Use CF-resolved color if available, else fall back to global setting
            const eArrow = edge.cfArrowColor ?? arrowC;
            const eLblC  = edge.cfLabelFont  ?? lblC;
            const eLblBg = edge.cfLabelBg    ?? lblBg;

            edgeG.append("line")
                .attr("x1", x1).attr("y1", y1).attr("x2", x2).attr("y2", y2)
                .attr("stroke", eArrow).attr("stroke-width", 1.5)
                .attr("marker-end", "url(#pmArrow)");

            const mx = (x1 + x2) / 2; const my = (y1 + y2) / 2;
            const lbl = edge.durationValue.toFixed(1);
            const fz  = s.edgeFormat.labelFontSize;
            const lw  = lbl.length * fz * 0.65 + 14; const lh = fz + 10;
            edgeG.append("rect")
                .attr("x", mx - lw / 2).attr("y", my - lh / 2)
                .attr("width", lw).attr("height", lh).attr("rx", 3)
                .attr("fill", eLblBg).attr("stroke", eArrow).attr("stroke-width", 0.5);
            edgeG.append("text")
                .attr("x", mx).attr("y", my)
                .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                .attr("font-size", `${fz}px`).attr("font-family", "Segoe UI, sans-serif")
                .attr("fill", eLblC).text(lbl)
                .append("title")
                    .text(`Avg ${vm.measureName}: ${edge.durationValue.toFixed(2)} ${s.edgeFormat.measureUnit}`);
        });

        // ── Nodes ─────────────────────────────────────────────────────────────
        const anyHL = vm.nodes.some(n => n.highlighted !== null);
        const nodeG = this.diagramG.append("g").classed("pm-nodes", true);
        const rx    = Math.max(0, Math.min(28, s.nodeFormat.borderRadius));

        vm.nodes.forEach((node, ni) => {
            const pos = centers[ni]; const box = boxes[ni];
            const g   = nodeG.append("g").classed("pm-node-group", true).attr("data-idx", ni);
            const op  = anyHL ? (node.highlighted ? 1 : dimOp) : 1;

            // Use CF-resolved color from view model if set, else global setting
            const nFill   = node.cfNodeFill   ?? nodeFill;
            const nBorder = node.cfNodeBorder ?? nodeBorder;
            const nFont   = node.cfNodeFont   ?? nodeFontC;

            if (!doExp || !node.expandNodes.length) {
                // When expand children exist (hasExpand) but expand display is off,
                // still wire up multiIds so clicking cross-filters other visuals.
                const nonExpChildIds = node.expandNodes.length > 0
                    ? node.expandNodes.map(en => en.selectionId).filter(Boolean)
                    : undefined;
                const nodeTip = `${node.label}\nCount: ${node.count}` +
                    (vm.measureName ? `\nAvg ${vm.measureName}: ${node.measureAgg.toFixed(2)} ${s.edgeFormat.measureUnit}` : "");
                this.drawRect(g, pos.cx - box.w / 2, pos.cy - box.h / 2, box.w, box.h,
                    nFill, nBorder, rx,
                    node.selectionId,   // non-null when no expand field at all
                    false, op,
                    nonExpChildIds,
                    nodeTip);

                g.append("text").attr("x", pos.cx).attr("y", pos.cy - 9)
                    .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                    .attr("font-size", `${s.nodeFormat.fontSize}px`)
                    .attr("font-weight", s.nodeFormat.fontBold ? "600" : "400")
                    .attr("font-family", "Segoe UI, sans-serif")
                    .attr("fill", nFont).attr("opacity", op)
                    .style("pointer-events", "none").text(this.trunc(node.label, 22));

                g.append("text").attr("x", pos.cx).attr("y", pos.cy + 11)
                    .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                    .attr("font-size", `${Math.max(9, s.nodeFormat.fontSize - 2)}px`)
                    .attr("font-family", "Segoe UI, sans-serif")
                    .attr("fill", nFont).attr("opacity", op * 0.85)
                    .style("pointer-events", "none").text(`Count: ${node.count}`);
            } else {
                // Stage container in expand mode: clicking multi-selects all child
                // expand node IDs. The container rect itself has no own selectionId
                // (null) to avoid emitting a bare stage predicate that Dataverse
                // TDS cannot translate (rsDataShapeQueryTranslationError).
                const childIds = node.expandNodes.map(en => en.selectionId).filter(Boolean);
                const containerTip = `${node.label}\nCount: ${node.count}` +
                    (vm.measureName ? `\nAvg ${vm.measureName}: ${node.measureAgg.toFixed(2)} ${s.edgeFormat.measureUnit}` : "");
                this.drawRect(
                    g,
                    pos.cx - box.w / 2 - 4, pos.cy - box.h / 2 - 4,
                    box.w + 8, box.h + 8,
                    "none", nBorder, 8,
                    null, false, op,
                    childIds,
                    containerTip
                );
                // Override the stroke-dasharray that drawRect can't set — add it directly
                g.select("rect.pm-node-rect")
                    .attr("stroke-dasharray", "4 3")
                    .attr("stroke-width", 1);

                g.append("text").attr("x", pos.cx).attr("y", pos.cy - box.h / 2 - 7)
                    .attr("text-anchor", "middle")
                    .attr("font-size", `${s.nodeFormat.fontSize}px`)
                    .attr("font-weight", s.nodeFormat.fontBold ? "600" : "400")
                    .attr("font-family", "Segoe UI, sans-serif")
                    .attr("fill", nFont === "#FFFFFF" ? "#201F1E" : nFont)
                    .attr("opacity", op).style("pointer-events", "none")
                    .text(this.trunc(node.label, 26));

                node.expandNodes.forEach((en, ei) => {
                    let ex: number, ey: number;
                    if (isV) {
                        const tw = node.expandNodes.length * EXPAND_W + (node.expandNodes.length - 1) * EXPAND_GAP;
                        ex = pos.cx - tw / 2 + ei * (EXPAND_W + EXPAND_GAP);
                        ey = pos.cy - EXPAND_H / 2;
                    } else {
                        const th = node.expandNodes.length * EXPAND_H + (node.expandNodes.length - 1) * EXPAND_GAP;
                        ex = pos.cx - EXPAND_W / 2;
                        ey = pos.cy - th / 2 + ei * (EXPAND_H + EXPAND_GAP);
                    }
                    const expFill   = (hc || !s.expandFormat.useDataColors) ? s.expandFormat.fillColor : en.color;
                    const expBorder = hc ? nodeBorder : s.expandFormat.borderColor;
                    const expFont   = hc ? nodeFontC  : s.expandFormat.fontColor;

                    const expandTip = `${en.label}\nCount: ${en.count}` +
                        (vm.measureName ? `\nAvg ${vm.measureName}: ${en.measureAgg.toFixed(2)} ${s.edgeFormat.measureUnit}` : "");
                    this.drawRect(g, ex, ey, EXPAND_W, EXPAND_H, expFill, expBorder, 5,
                        en.selectionId, true, op, undefined, expandTip);

                    g.append("text").attr("x", ex + EXPAND_W / 2).attr("y", ey + EXPAND_H * 0.36)
                        .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                        .attr("font-size", `${s.expandFormat.fontSize}px`)
                        .attr("font-family", "Segoe UI, sans-serif")
                        .attr("fill", expFont).attr("opacity", op)
                        .style("pointer-events", "none").text(this.trunc(en.label, 17));

                    g.append("text").attr("x", ex + EXPAND_W / 2).attr("y", ey + EXPAND_H * 0.70)
                        .attr("text-anchor", "middle").attr("dominant-baseline", "middle")
                        .attr("font-size", `${Math.max(8, s.expandFormat.fontSize - 2)}px`)
                        .attr("font-family", "Segoe UI, sans-serif")
                        .attr("fill", expFont).attr("opacity", op * 0.85)
                        .style("pointer-events", "none").text(`Count: ${en.count}`);
                });
            }
        });
    }

    // ── Draw clickable rect ───────────────────────────────────────────────────
    // selId:     single selection ID (used for both click and context menu).
    //            Pass null for stage container rects in expand mode — these use
    //            multiIds instead so they don't emit a bare stage predicate that
    //            Dataverse TDS cannot translate (rsDataShapeQueryTranslationError).
    // multiIds:  when provided, clicking selects ALL ids in the array (multi-select).
    //            Used for stage nodes when expand is active: selects all child owners.
    private drawRect(
        parent: d3.Selection<SVGGElement, unknown, null, undefined>,
        x: number, y: number, w: number, h: number,
        fill: string, stroke: string, rxv: number,
        selId: ISelectionId | null, isExpand: boolean, opacity: number,
        multiIds?: ISelectionId[],
        tooltip?: string
    ): void {
        const rect = parent.append("rect")
            .classed(isExpand ? "pm-expand-rect" : "pm-node-rect", true)
            .attr("x", x).attr("y", y).attr("width", w).attr("height", h)
            .attr("rx", rxv).attr("ry", rxv)
            .attr("fill", fill).attr("stroke", stroke).attr("stroke-width", 1.5)
            .attr("opacity", opacity).style("cursor", "pointer");

        if (tooltip) {
            rect.append("title").text(tooltip);
        }

        // Stamp the selection key directly on expand rects so paintHighlight can
        // match individual expand children rather than the whole parent group.
        if (isExpand && selId) {
            rect.attr("data-selkey", (selId as any).key ?? "");
        }

        if (multiIds && multiIds.length > 0) {
            // Stage node in expand mode: select all child expand IDs at once.
            // This translates to an IN-list on the expand column only, which
            // Dataverse can handle cleanly.
            rect.on("click", (event: MouseEvent) => {
                event.stopPropagation();
                const multi = event.ctrlKey || event.metaKey;

                // FIX: toggle deselect — if every child ID of this node is already
                // selected, clear the selection entirely instead of re-selecting.
                // This restores all nodes to their un-dimmed state on a second click,
                // mirroring the behaviour of single-selectionId rects.
                const currentKeys = new Set(
                    (this.selectionManager.getSelectionIds() as ISelectionId[])
                        .map(id => (id as any).key));
                const allAlreadySelected = multiIds.every(
                    id => currentKeys.has((id as any).key));

                if (allAlreadySelected && !multi) {
                    this.selectionManager.clear().then(() =>
                        this.paintHighlight([], false));
                    return;
                }

                // selectionManager.select() returns IPromise2, not a native Promise,
                // so we cannot type a recursive chain as Promise<void>.
                // Instead we reduce over the ids array using the IPromise2 .then()
                // directly, cast to `any` to bridge the IPromise2/Promise mismatch,
                // and repaint once all selections have been queued.
                const ids = multiIds;
                let chain: any = this.selectionManager.select(ids[0], multi);
                for (let i = 1; i < ids.length; i++) {
                    const id = ids[i];
                    chain = chain.then(() => this.selectionManager.select(id, true));
                }
                (chain as any).then(() =>
                    this.paintHighlight(
                        this.selectionManager.getSelectionIds() as ISelectionId[], true));
            });
            // Context menu: use first child's ID as anchor
            rect.on("contextmenu", (event: MouseEvent) => {
                event.preventDefault();
                this.selectionManager.showContextMenu(multiIds[0], {
                    x: event.clientX, y: event.clientY
                });
            });
        } else if (selId) {
            rect.on("click", (event: MouseEvent) => {
                event.stopPropagation();
                const multi = event.ctrlKey || event.metaKey;

                // FIX: toggle deselect for single-ID rects (expand children).
                // If this exact ID is the only selection, clear instead of re-selecting
                // so all nodes return to full opacity without requiring a background click.
                const currentIds = this.selectionManager.getSelectionIds() as ISelectionId[];
                const currentKeys = new Set(currentIds.map(id => (id as any).key));
                const alreadySelected = currentKeys.has((selId as any).key);

                if (alreadySelected && !multi && currentIds.length === 1) {
                    this.selectionManager.clear().then(() =>
                        this.paintHighlight([], false));
                    return;
                }

                this.selectionManager
                    .select(selId, multi)
                    .then(() => this.paintHighlight(
                        this.selectionManager.getSelectionIds() as ISelectionId[], true));
            })
            .on("contextmenu", (event: MouseEvent) => {
                event.preventDefault();
                this.selectionManager.showContextMenu(selId, {
                    x: event.clientX, y: event.clientY
                });
            });
        }
    }

    // ── Outbound highlight after node click ───────────────────────────────────
    private paintHighlight(selectedIds: ISelectionId[], hasSelection: boolean): void {
        if (!this.viewModel?.hasData) return;
        const s     = this.settings;
        const hc    = this.host.colorPalette.isHighContrast;
        const dimOp = Math.max(0, Math.min(100, s.selectionFormat.dimOpacity ?? 30)) / 100;
        const selKeys = new Set(selectedIds.map(id => (id as any).key));

        // Build a set of node-group indices that are "selected".
        // Matches on own stage selectionId OR any expand child selectionId.
        const selectedNodeIdxs = new Set<number>();
        if (hasSelection) {
            this.viewModel.nodes.forEach((node, ni) => {
                const ownMatch   = node.selectionId && selKeys.has((node.selectionId as any).key);
                const childMatch = node.expandNodes.some(
                    en => en.selectionId && selKeys.has((en.selectionId as any).key));
                if (ownMatch || childMatch) selectedNodeIdxs.add(ni);
            });
        }

        const nodeBorder   = hc ? this.host.colorPalette.foreground.value : s.nodeFormat.borderColor;
        const expandBorder = hc ? this.host.colorPalette.foreground.value : s.expandFormat.borderColor;

        this.diagramG.selectAll<SVGRectElement, unknown>(".pm-node-rect, .pm-expand-rect")
            .attr("opacity", (_, i, els) => {
                if (!hasSelection) return 1;
                const el  = els[i] as Element;
                const grp = el.parentElement;
                const idx = grp ? parseInt(grp.getAttribute("data-idx") ?? "-1", 10) : -1;
                if (idx < 0 || idx >= this.viewModel.nodes.length) return 1;

                // FIX: for individual expand rects, match by their own stamped key
                // rather than the parent group index. This ensures clicking "Approver"
                // only highlights "Approver", not every expand child in the same stage.
                if (el.classList.contains("pm-expand-rect")) {
                    const ek = el.getAttribute("data-selkey");
                    if (ek) return selKeys.has(ek) ? 1 : dimOp;
                    // fallback: no key stamped — dim if nothing in this group selected
                    return selectedNodeIdxs.has(idx) ? 1 : dimOp;
                }

                return selectedNodeIdxs.has(idx) ? 1 : dimOp;
            })
            .attr("stroke", (_, i, els) => {
                const el    = els[i] as Element;
                const isExp = el.classList.contains("pm-expand-rect");
                const defaultStroke = isExp ? expandBorder : nodeBorder;
                if (!hasSelection) return defaultStroke;
                const grp = el.parentElement;
                const idx = grp ? parseInt(grp.getAttribute("data-idx") ?? "-1", 10) : -1;
                if (idx < 0 || idx >= this.viewModel.nodes.length) return defaultStroke;

                // FIX: same individual-key matching for expand rect borders
                if (isExp) {
                    const ek = el.getAttribute("data-selkey");
                    if (ek) return selKeys.has(ek) ? s.selectionFormat.highlightColor : defaultStroke;
                    return selectedNodeIdxs.has(idx) ? s.selectionFormat.highlightColor : defaultStroke;
                }

                return selectedNodeIdxs.has(idx) ? s.selectionFormat.highlightColor : defaultStroke;
            })
            .attr("stroke-width", (_, i, els) => {
                if (!hasSelection) return 1.5;
                const el  = els[i] as Element;
                const grp = el.parentElement;
                const idx = grp ? parseInt(grp.getAttribute("data-idx") ?? "-1", 10) : -1;
                if (idx < 0 || idx >= this.viewModel.nodes.length) return 1.5;

                // FIX: same individual-key matching for expand rect stroke width
                if (el.classList.contains("pm-expand-rect")) {
                    const ek = el.getAttribute("data-selkey");
                    if (ek) return selKeys.has(ek) ? 2.5 : 1.5;
                    return selectedNodeIdxs.has(idx) ? 2.5 : 1.5;
                }

                return selectedNodeIdxs.has(idx) ? 2.5 : 1.5;
            });
    }

    // ── Toolbar ───────────────────────────────────────────────────────────────
    private updateToolbarState(): void {
        this.toolbar.select(".pm-btn-orient")
            .text(this.orientation === "vertical" ? "⇅ Vertical" : "⇆ Horizontal");
        this.toolbar.select(".pm-zoom-label").text(`${this.zoomLevel}%`);
        this.toolbar.select(".pm-btn-zoom-out")
            .style("opacity", this.zoomLevel <= ZOOM_MIN ? "0.4" : "1");
        this.toolbar.select(".pm-btn-zoom-in")
            .style("opacity", this.zoomLevel >= ZOOM_MAX ? "0.4" : "1");

        const expAvail = !!(this.viewModel?.hasExpand);
        this.toolbar.select(".pm-btn-expand")
            .style("background",   this.expandMode && expAvail ? "#0078D4" : "#F3F2F1")
            .style("color",        this.expandMode && expAvail ? "#FFFFFF"  : "#201F1E")
            .style("border-color", this.expandMode && expAvail ? "#004578"  : "#8A8886")
            .style("opacity",      expAvail ? "1" : "0.35")
            .style("cursor",       expAvail ? "pointer" : "not-allowed");
    }

    // ── Landing page ──────────────────────────────────────────────────────────
    private renderLanding(viewport: powerbi.IViewport): void {
        this.svgWrap.attr("width", viewport.width).attr("height", viewport.height - TOOLBAR_H);
        const cx = viewport.width / 2; const cy = (viewport.height - TOOLBAR_H) / 2;
        this.diagramG.append("text").attr("x", cx).attr("y", cy - 12)
            .attr("text-anchor", "middle").attr("font-size", "14px")
            .attr("font-family", "Segoe UI Semibold, Segoe UI, sans-serif").attr("fill", "#605E5C")
            .text("Process Mining Flow");
        this.diagramG.append("text").attr("x", cx).attr("y", cy + 10)
            .attr("text-anchor", "middle").attr("font-size", "12px")
            .attr("font-family", "Segoe UI, sans-serif").attr("fill", "#A19F9D")
            .text("Add Stage and Duration Measure fields to begin.");
    }

    private trunc(v: string, max: number): string {
        return v.length > max ? v.slice(0, max - 1) + "…" : v;
    }
}