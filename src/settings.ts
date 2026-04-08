"use strict";

import { DataViewObjectsParser } from "powerbi-visuals-utils-dataviewutils/lib/dataViewObjectsParser";
import DataView = powerbi.DataView;

// ---------------------------------------------------------------------------
// Helper — unwrap a Power BI Fill object to a plain hex string.
// DataViewObjectsParser copies the raw dataView object value into the typed
// class property.  For color fields, Power BI stores { solid: { color: "#…" } }
// (ISolidFill), not a plain string.  The parser doesn't know the property is a
// color, so it assigns the object as-is into a string-typed field.
// This helper normalises both cases so callers always receive a hex string.
// ---------------------------------------------------------------------------
function unwrapColor(value: any, fallback: string): string {
    if (!value) return fallback;
    // Power BI Fill object: { solid: { color: "#RRGGBB" } }
    if (typeof value === "object" && value.solid?.color) {
        return value.solid.color;
    }
    // Already a plain string (e.g. the class default was never overwritten)
    if (typeof value === "string" && /^#[0-9A-Fa-f]{6}$/i.test(value.trim())) {
        return value.trim();
    }
    return fallback;
}

// ---------------------------------------------------------------------------
// Settings classes — property defaults are used both as initial values and as
// the fallback inside unwrapColor when no user value has been saved yet.
// ---------------------------------------------------------------------------

export class LayoutSettings {
    public orientation: string  = "vertical";
    public expandMode:  boolean = false;
    public nodeSpacing: number  = 80;
    public zoomLevel:   number  = 100;
}

export class NodeFormatSettings {
    public fillColor:    string  = "#0078D4";
    public borderColor:  string  = "#004578";
    public fontColor:    string  = "#FFFFFF";
    public fontSize:     number  = 12;
    public fontBold:     boolean = true;
    public borderRadius: number  = 6;
}

export class ExpandFormatSettings {
    public fillColor:     string  = "#005A9E";
    public borderColor:   string  = "#003966";
    public fontColor:     string  = "#FFFFFF";
    public fontSize:      number  = 11;
    public useDataColors: boolean = false;
}

export class EdgeFormatSettings {
    public arrowColor:     string = "#605E5C";
    public labelFontSize:  number = 11;
    public labelFontColor: string = "#201F1E";
    public labelBgColor:   string = "#F3F2F1";
    public measureUnit:    string = "days";
}

export class SelectionFormatSettings {
    public highlightColor: string = "#FFB900";
    public dimOpacity:     number = 30;
}

export class VisualSettings extends DataViewObjectsParser {
    public layout:          LayoutSettings          = new LayoutSettings();
    public nodeFormat:      NodeFormatSettings      = new NodeFormatSettings();
    public expandFormat:    ExpandFormatSettings    = new ExpandFormatSettings();
    public edgeFormat:      EdgeFormatSettings      = new EdgeFormatSettings();
    public selectionFormat: SelectionFormatSettings = new SelectionFormatSettings();

    // -------------------------------------------------------------------------
    // Override parse() to unwrap Fill objects into plain hex strings immediately
    // after DataViewObjectsParser has populated the class properties.
    // This means every consumer of VisualSettings (visual.ts, getFormattingModel,
    // renderFromState) always receives clean hex strings — no further unwrapping
    // needed anywhere else in the codebase.
    // -------------------------------------------------------------------------
    public static parse<T extends DataViewObjectsParser>(dataView: DataView): T {
        // Let the base parser do its normal work first
        const settings = super.parse<VisualSettings>(dataView) as VisualSettings;

        // selectionFormat uses selector:null so metadata.objects is correct — unwrap only
        const sf = settings.selectionFormat;
        sf.highlightColor = unwrapColor(sf.highlightColor, "#FFB900");

        // Shorthand refs used below
        const nd = settings.nodeFormat;
        const ex = settings.expandFormat;
        const ed = settings.edgeFormat;

        // Fix: with a wildcard selector in enumerateObjectInstances, Power BI stores ALL
        // property values — both colors and non-color props — in category.objects rather
        // than metadata.objects. DataViewObjectsParser only reads metadata.objects, so
        // every saved value (fill, font size, bold, etc.) is invisible to the base parser.
        // Read everything back from category.objects[0] here, which is where Power BI
        // writes constant values (set via the plain picker, not the fx rule button).
        const dvCat   = (dataView as any)?.categorical;
        const stageCl = dvCat?.categories?.find((c: any) => c.source?.roles?.["stage"]);
        const expCl   = dvCat?.categories?.find((c: any) => c.source?.roles?.["expand"]);

        // Read a Fill color from category.objects[0]; fall back to `fallback`.
        const catFill = (col: any, obj: string, prop: string, fallback: string): string => {
            const v = col?.objects?.[0]?.[obj]?.[prop];
            if (!v) return fallback;
            if (typeof v === "object" && v.solid?.color) return v.solid.color;
            if (typeof v === "string" && /^#[0-9A-Fa-f]{6}$/i.test(v.trim())) return v.trim();
            return fallback;
        };
        const catNum  = (col: any, obj: string, prop: string, fallback: number): number => {
            const v = col?.objects?.[0]?.[obj]?.[prop];
            return (typeof v === "number" && isFinite(v) && v > 0) ? v : fallback;
        };
        const catBool = (col: any, obj: string, prop: string, fallback: boolean): boolean => {
            const v = col?.objects?.[0]?.[obj]?.[prop];
            return typeof v === "boolean" ? v : fallback;
        };
        const catStr  = (col: any, obj: string, prop: string, fallback: string): string => {
            const v = col?.objects?.[0]?.[obj]?.[prop];
            return typeof v === "string" && v.length > 0 ? v : fallback;
        };

        // Colors — try category.objects[0] first, fall back to what unwrapColor found
        nd.fillColor    = catFill(stageCl, "nodeFormat",   "fillColor",    unwrapColor(nd.fillColor,    "#0078D4"));
        nd.borderColor  = catFill(stageCl, "nodeFormat",   "borderColor",  unwrapColor(nd.borderColor,  "#004578"));
        nd.fontColor    = catFill(stageCl, "nodeFormat",   "fontColor",    unwrapColor(nd.fontColor,    "#FFFFFF"));

        const efCol = expCl ?? stageCl;
        ex.fillColor    = catFill(efCol,   "expandFormat", "fillColor",    unwrapColor(ex.fillColor,    "#005A9E"));
        ex.borderColor  = catFill(efCol,   "expandFormat", "borderColor",  unwrapColor(ex.borderColor,  "#003966"));
        ex.fontColor    = catFill(efCol,   "expandFormat", "fontColor",    unwrapColor(ex.fontColor,    "#FFFFFF"));

        ed.arrowColor     = catFill(stageCl, "edgeFormat", "arrowColor",     unwrapColor(ed.arrowColor,     "#605E5C"));
        ed.labelFontColor = catFill(stageCl, "edgeFormat", "labelFontColor", unwrapColor(ed.labelFontColor, "#201F1E"));
        ed.labelBgColor   = catFill(stageCl, "edgeFormat", "labelBgColor",   unwrapColor(ed.labelBgColor,   "#F3F2F1"));

        // Non-color props
        nd.fontSize     = catNum(stageCl,  "nodeFormat",   "fontSize",      nd.fontSize);
        nd.fontBold     = catBool(stageCl, "nodeFormat",   "fontBold",      nd.fontBold);
        nd.borderRadius = catNum(stageCl,  "nodeFormat",   "borderRadius",  nd.borderRadius);

        ex.fontSize      = catNum(efCol,   "expandFormat", "fontSize",      ex.fontSize);
        ex.useDataColors = catBool(efCol,  "expandFormat", "useDataColors", ex.useDataColors);

        ed.labelFontSize = catNum(stageCl, "edgeFormat",   "labelFontSize", ed.labelFontSize);
        ed.measureUnit   = catStr(stageCl, "edgeFormat",   "measureUnit",   ed.measureUnit);

        return settings as unknown as T;
    }
}