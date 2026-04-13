import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { InfoCardComponent, InfoCardData, InfoCardTheme, SlotField, LayoutMode, defaultTheme } from "./InfoCard";
import * as React from "react";

interface RelatedFieldMapping {
    sourceSlot: string;
    fetchField: string;
    targetSlot: string;
}

const CONTROL_VERSION = "2.4.7";

// Slot group definitions — order matters for rendering
const SUBTITLE_KEYS = ["subtitleField1", "subtitleField2", "subtitleField3"] as const;
const PHONE_KEYS = ["phoneField1", "phoneField2"] as const;
const DETAIL_KEYS = ["detailField1", "detailField2", "detailField3"] as const;
const GRID_KEYS = [
    "gridField1", "gridField2", "gridField3", "gridField4", "gridField5", "gridField6",
] as const;
const TAG_KEYS = ["tagField1", "tagField2", "tagField3"] as const;

export class InfoCard implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private context: ComponentFramework.Context<IInputs>;

    constructor() { }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
    ): void {
        this.context = context;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;

        const data = this.collectData(context);
        const layout = this.resolveLayout(context);
        const hideEmpty = context.parameters.hideEmptyFields?.raw !== false; // default true
        const showBorder = context.parameters.showCardBorder?.raw !== false; // default true
        const showVersion = context.parameters.showVersionInfo?.raw === true; // default false
        const relatedMappings = this.detectRelatedFields(context);
        const theme = this.resolveTheme(context);

        const onOpenRecord = (entityType: string, id: string) => {
            context.navigation.openForm({
                entityName: entityType,
                entityId: id,
            });
        };

        const fetchRelatedData = (entityType: string, id: string, columns: string[]) => {
            return this.fetchRelatedFields(context, entityType, id, columns);
        };

        return React.createElement(InfoCardComponent, {
            data,
            layout,
            hideEmpty,
            showBorder,
            showVersion,
            theme,
            version: CONTROL_VERSION,
            relatedMappings,
            fetchRelatedData,
            onOpenRecord,
        });
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void { }

    // ────────────────────────────────────────
    // Data collection
    // ────────────────────────────────────────

    private collectData(context: ComponentFramework.Context<IInputs>): InfoCardData {
        // Read lat/lng as numbers
        const latSlot = this.readSlot(context, "latitudeField");
        const lngSlot = this.readSlot(context, "longitudeField");
        const lat = latSlot && !latSlot.isEmpty ? Number(latSlot.rawValue) : null;
        const lng = lngSlot && !lngSlot.isEmpty ? Number(lngSlot.rawValue) : null;

        return {
            title: this.readSlot(context, "titleField"),
            subtitles: this.readSlotGroup(context, SUBTITLE_KEYS),
            phones: this.readSlotGroup(context, PHONE_KEYS),
            email: this.readSlot(context, "emailField"),
            web: this.readSlot(context, "webField"),
            address: this.readSlot(context, "addressField"),
            latitude: (lat != null && !isNaN(lat)) ? lat : null,
            longitude: (lng != null && !isNaN(lng)) ? lng : null,
            details: this.readSlotGroup(context, DETAIL_KEYS),
            gridFields: this.readSlotGroup(context, GRID_KEYS),
            tags: this.readSlotGroup(context, TAG_KEYS),
            imageUrl: null,
        };
    }

    private readSlotGroup(
        context: ComponentFramework.Context<IInputs>,
        keys: readonly string[],
    ): SlotField[] {
        const fields: SlotField[] = [];
        for (const key of keys) {
            const field = this.readSlot(context, key);
            if (field) {
                fields.push(field);
            }
        }
        return fields;
    }

    private readSlot(
        context: ComponentFramework.Context<IInputs>,
        key: string,
    ): SlotField | null {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const param = (context.parameters as Record<string, any>)[key];
        if (!param || param.type === "Unknown") return null;

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const attrs = (param as Record<string, any>).attributes;
        if (!attrs?.LogicalName) return null;

        const label = attrs.DisplayName ?? attrs.LogicalName;
        const raw = param.raw;
        const formatted = param.formatted;

        let displayValue = "";
        let isEmpty = false;

        // @sourceSlot.fieldName values are related field placeholders — treat as empty
        // The actual value will be populated by the related field fetch
        if (typeof raw === "string" && raw.startsWith("@")) {
            isEmpty = true;
            displayValue = "---";
        } else if (raw === null || raw === undefined) {
            isEmpty = true;
            displayValue = "---";
        } else if (formatted) {
            displayValue = formatted;
        } else if (typeof raw === "object" && Array.isArray(raw) && raw.length > 0) {
            // Lookup
            displayValue = raw[0].name ?? String(raw[0].id ?? "---");
        } else if (typeof raw === "object" && !Array.isArray(raw)) {
            // Edge case: single lookup object
            displayValue = (raw as { name?: string }).name ?? "---";
        } else {
            displayValue = String(raw);
        }

        if (displayValue === "" || displayValue === "---") {
            isEmpty = true;
        }

        // Format duration fields (stored as minutes in Dataverse)
        if (!isEmpty && this.isDurationField(attrs, raw, formatted)) {
            displayValue = this.formatDuration(Number(raw));
        }

        // Extract lookup entity/id for navigation
        let lookupEntityType: string | undefined;
        let lookupId: string | undefined;
        if (typeof raw === "object" && Array.isArray(raw) && raw.length > 0 && raw[0].id) {
            lookupEntityType = raw[0].entityType ?? raw[0].etn;
            lookupId = raw[0].id;
        }

        return {
            label,
            value: isEmpty ? "---" : displayValue,
            rawValue: raw,
            isEmpty,
            lookupEntityType,
            lookupId,
        };
    }

    // ────────────────────────────────────────
    // Related fields
    // ────────────────────────────────────────

    /**
     * Detect related field references using the @fieldName convention.
     *
     * Any property with a static value starting with @ is treated as a related field
     * fetched from the entity that titleField is a lookup to:
     *   @msdyn_serviceaccount → fetch msdyn_serviceaccount from the title lookup entity
     *   @telephone1 → fetch telephone1 from the title lookup entity
     *
     * The source is always titleField (the anchor/bound field).
     */
    private detectRelatedFields(context: ComponentFramework.Context<IInputs>): RelatedFieldMapping[] {
        const mappings: RelatedFieldMapping[] = [];
        const allSlotKeys = [
            ...SUBTITLE_KEYS, ...PHONE_KEYS, "emailField", "webField",
            "addressField", "latitudeField", "longitudeField",
            ...DETAIL_KEYS, ...GRID_KEYS, ...TAG_KEYS,
        ];

        for (const slotKey of allSlotKeys) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const param = (context.parameters as Record<string, any>)[slotKey];
            if (!param) continue;
            const raw = param.raw;
            if (typeof raw !== "string" || !raw.startsWith("@")) continue;

            const fetchField = raw.substring(1).trim();
            if (fetchField) {
                mappings.push({
                    sourceSlot: "titleField",
                    fetchField,
                    targetSlot: slotKey,
                });
            }
        }

        return mappings;
    }

    private async fetchRelatedFields(
        context: ComponentFramework.Context<IInputs>,
        entityType: string,
        id: string,
        columns: string[],
    ): Promise<Record<string, { value: string; label: string; lookupId?: string; lookupEntityType?: string }>> {
        try {
            const results: Record<string, { value: string; label: string; lookupId?: string; lookupEntityType?: string }> = {};

            // Split columns into hop1 (single field) and hop2 (dotted path)
            const hop1Columns: string[] = [];
            const hop2Paths: Array<{ original: string; hop1Field: string; hop2Field: string }> = [];

            for (const col of columns) {
                const dotIdx = col.indexOf(".");
                if (dotIdx > 0) {
                    const hop1Field = col.substring(0, dotIdx);
                    const hop2Field = col.substring(dotIdx + 1);
                    hop2Paths.push({ original: col, hop1Field, hop2Field });
                    // Ensure hop1 field is fetched too (we need the lookup ID for hop2)
                    if (!hop1Columns.includes(hop1Field)) hop1Columns.push(hop1Field);
                } else {
                    hop1Columns.push(col);
                }
            }

            // Hop 1: fetch from the title's entity with only the columns we need
            const selectCols = hop1Columns.flatMap(col => {
                // For each column, request both the direct field and the lookup pattern
                const cols = [col, `_${col}_value`];
                return cols;
            });
            const selectParam = selectCols.join(",");
            const hop1Record = await context.webAPI.retrieveRecord(entityType, id, `?$select=${selectParam}`);

            // Helper: read a field value from a record
            const readField = (record: Record<string, unknown>, col: string): {
                value: string; label: string; lookupId?: string; lookupEntityType?: string;
            } | null => {
                // Image columns
                if (col === "entityimage" || col.endsWith("image") || col.endsWith("_image")) {
                    const allImageKeys = Object.keys(record).filter(k =>
                        k.toLowerCase().includes("image") && !k.includes("@") && !k.includes("timestamp"));
                    for (const key of allImageKeys) {
                        if (key.toLowerCase().endsWith("_url") && record[key]) {
                            return { value: String(record[key]), label: "image" };
                        }
                    }
                    for (const key of allImageKeys) {
                        const val = record[key];
                        if (val && typeof val === "string" && (val as string).length > 100 && !key.endsWith("_url")) {
                            return { value: `data:image/png;base64,${val}`, label: "image" };
                        }
                    }
                    return null;
                }

                // Lookup pattern: _col_value
                const lookupKey = `_${col}_value`;
                const lookupFormatted = record[`${lookupKey}@OData.Community.Display.V1.FormattedValue`]
                    ?? record[`${col}@OData.Community.Display.V1.FormattedValue`];
                const lookupVal = record[lookupKey];
                if (lookupFormatted) {
                    const etnKey = `${lookupKey}@Microsoft.Dynamics.CRM.lookuplogicalname`;
                    return {
                        value: String(lookupFormatted),
                        label: col,
                        lookupId: lookupVal ? String(lookupVal) : undefined,
                        lookupEntityType: record[etnKey] ? String(record[etnKey]) : undefined,
                    };
                }
                if (lookupVal != null) {
                    return { value: String(lookupVal), label: col, lookupId: String(lookupVal) };
                }

                // Direct field
                const formatted = record[`${col}@OData.Community.Display.V1.FormattedValue`];
                if (formatted) return { value: String(formatted), label: col };
                if (record[col] != null) return { value: String(record[col]), label: col };

                return null;
            };

            // Process hop1 columns
            for (const col of hop1Columns) {
                // Skip hop1-only fields that are just intermediaries for hop2
                const isHop2Intermediary = hop2Paths.some(p => p.hop1Field === col) && !columns.includes(col);
                const result = readField(hop1Record, col);
                if (result && !isHop2Intermediary) {
                    results[col] = result;
                }
            }

            // Hop 2: group by the intermediary lookup, fetch each once with all needed columns
            if (hop2Paths.length > 0) {
                // Group hop2 paths by their hop1 field (the intermediary lookup)
                const hop2Groups: Record<string, { entityType: string; entityId: string; fields: string[]; originals: string[] }> = {};

                for (const path of hop2Paths) {
                    const hop1Result = readField(hop1Record, path.hop1Field);
                    if (!hop1Result?.lookupId) continue;

                    const key = `${hop1Result.lookupEntityType ?? path.hop1Field}:${hop1Result.lookupId}`;
                    if (!hop2Groups[key]) {
                        hop2Groups[key] = {
                            entityType: hop1Result.lookupEntityType ?? path.hop1Field,
                            entityId: hop1Result.lookupId,
                            fields: [],
                            originals: [],
                        };
                    }
                    if (!hop2Groups[key].fields.includes(path.hop2Field)) {
                        hop2Groups[key].fields.push(path.hop2Field);
                    }
                    hop2Groups[key].originals.push(path.original);
                }

                // Fetch each hop2 entity once with all needed columns
                for (const group of Object.values(hop2Groups)) {
                    try {
                        const hop2Select = group.fields.flatMap(f => [f, `_${f}_value`]).join(",");
                        const hop2Record = await context.webAPI.retrieveRecord(
                            group.entityType, group.entityId, `?$select=${hop2Select}`
                        );

                        // Map results back to original dotted paths
                        for (const path of hop2Paths) {
                            const pathKey = `${group.entityType}:${group.entityId}`;
                            const checkKey = `${readField(hop1Record, path.hop1Field)?.lookupEntityType ?? path.hop1Field}:${readField(hop1Record, path.hop1Field)?.lookupId}`;
                            if (pathKey !== checkKey) continue;

                            const result = readField(hop2Record, path.hop2Field);
                            if (result) {
                                results[path.original] = result;
                            }
                        }
                    } catch {
                        // Hop2 fetch failed — skip silently
                    }
                }
            }

            // Resolve display names from entity metadata
            try {
                const colsToResolve = Object.keys(results).filter(k => !k.startsWith("__") && !k.includes("."));
                if (colsToResolve.length > 0) {
                    const entityMeta = await context.utils.getEntityMetadata(entityType, colsToResolve);
                    if (entityMeta?.Attributes) {
                        for (const attr of entityMeta.Attributes.getAll()) {
                            const logicalName = attr.LogicalName;
                            const displayName = attr.DisplayName;
                            if (logicalName && displayName && results[logicalName]) {
                                results[logicalName].label = displayName;
                            }
                        }
                    }
                }
            } catch { /* metadata fetch is optional */ }

            return results;
        } catch (err) {
            // Return error info for debug diagnostics
            return {
                "__debug_error": {
                    value: String(err).substring(0, 150),
                    label: "error",
                },
            };
        }
    }

    // ────────────────────────────────────────
    // Duration formatting
    // ────────────────────────────────────────

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private isDurationField(attrs: Record<string, any>, raw: unknown, formatted: string | undefined): boolean {
        if (typeof raw !== "number") return false;
        // Check Dataverse Format attribute (may be "duration" or "Duration")
        const fmt = String(attrs?.Format ?? "").toLowerCase();
        if (fmt === "duration") return true;
        // Fallback: detect from the platform-formatted string
        if (formatted) {
            const f = formatted.toLowerCase().trim();
            if (f.endsWith("hours") || f.endsWith("hour") || f.endsWith("minutes") || f.endsWith("minute")) {
                return true;
            }
        }
        return false;
    }

    private formatDuration(minutes: number): string {
        if (minutes < 0) return String(minutes);
        const days = Math.floor(minutes / 1440);
        const hrs = Math.floor((minutes % 1440) / 60);
        const mins = minutes % 60;

        const parts: string[] = [];
        if (days > 0) parts.push(`${days}d`);
        if (hrs > 0) parts.push(`${hrs}h`);
        if (mins > 0) parts.push(`${mins}m`);

        return parts.length > 0 ? parts.join(" ") : "0m";
    }

    // ────────────────────────────────────────
    // Config resolution
    // ────────────────────────────────────────

    private resolveLayout(context: ComponentFramework.Context<IInputs>): LayoutMode {
        const raw = context.parameters.layout?.raw;
        if (raw === "smart" || raw === "contact" || raw === "compact") return raw;
        return "contact"; // default
    }

    private resolveTheme(context: ComponentFramework.Context<IInputs>): InfoCardTheme {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const fdl = (context as Record<string, any>).fluentDesignLanguage;
        if (fdl?.tokenTheme) {
            const t = fdl.tokenTheme;
            return {
                cardBg: t.colorNeutralBackground1 ?? defaultTheme.cardBg,
                textPrimary: t.colorNeutralForeground1 ?? defaultTheme.textPrimary,
                textSecondary: t.colorNeutralForeground3 ?? defaultTheme.textSecondary,
                textMuted: t.colorNeutralForeground4 ?? defaultTheme.textMuted,
                border: t.colorNeutralStroke1 ?? defaultTheme.border,
                borderLight: t.colorNeutralStroke2 ?? defaultTheme.borderLight,
                brand: t.colorBrandForeground1 ?? defaultTheme.brand,
                brandLight: t.colorBrandBackground2 ?? defaultTheme.brandLight,
                radius: t.borderRadiusMedium ?? defaultTheme.radius,
                shadow: t.shadow4 ?? defaultTheme.shadow,
                fontFamily: defaultTheme.fontFamily,
            };
        }
        return defaultTheme;
    }
}
