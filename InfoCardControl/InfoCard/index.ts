import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { InfoCardComponent, InfoCardData, InfoCardTheme, SlotField, LayoutMode, defaultTheme } from "./InfoCard";
import type { RelatedFieldMapping, BindingDiagnostic } from "./InfoCard";
import * as React from "react";

const CONTROL_VERSION = "3.9.8";

// Slot group definitions — order matters for rendering
const SUBTITLE_KEYS = ["subtitleField1", "subtitleField2", "subtitleField3"] as const;
const PHONE_KEYS = ["phoneField1", "phoneField2"] as const;
const DETAIL_KEYS = ["detailField1", "detailField2", "detailField3"] as const;
const GRID_KEYS = [
    "gridField1", "gridField2", "gridField3", "gridField4", "gridField5", "gridField6",
] as const;
const TAG_KEYS = ["tagField1", "tagField2", "tagField3"] as const;
const ALL_SLOT_KEYS = [
    "titleField", ...SUBTITLE_KEYS, ...PHONE_KEYS, "emailField", "webField",
    "addressField", "latitudeField", "longitudeField",
    ...DETAIL_KEYS, ...GRID_KEYS, ...TAG_KEYS,
];

// Friendly names for slot keys in diagnostics
const SLOT_LABELS: Record<string, string> = {
    titleField: "Title", subtitleField1: "Subtitle 1", subtitleField2: "Subtitle 2", subtitleField3: "Subtitle 3",
    phoneField1: "Phone 1", phoneField2: "Phone 2", emailField: "Email", webField: "Website",
    addressField: "Address", latitudeField: "Latitude", longitudeField: "Longitude",
    detailField1: "Detail 1", detailField2: "Detail 2", detailField3: "Detail 3",
    gridField1: "Grid 1", gridField2: "Grid 2", gridField3: "Grid 3",
    gridField4: "Grid 4", gridField5: "Grid 5", gridField6: "Grid 6",
    tagField1: "Tag 1", tagField2: "Tag 2", tagField3: "Tag 3",
};

interface ColumnMeta {
    displayName: string;
    logicalName: string;
    options?: Array<{ value: number; label: string; color?: string }>;
}

export class InfoCard implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private context: ComponentFramework.Context<IInputs>;

    // Stable references for React useEffect callbacks
    private readonly boundFetchRelatedData = (entityType: string, id: string, columns: string[]) => {
        const isCurrentRecord = entityType === this._formEntityName && id === this._formRecordId;
        return this.fetchRelatedFields(this.context, entityType, id, columns, isCurrentRecord);
    };
    private readonly boundResolveRecordFields = () => {
        return this.resolveRecordFieldsAsync(this.context);
    };

    // Form entity context
    private _formEntityName: string | null = null;
    private _formRecordId: string | null = null;

    // Entity metadata cache for bound field labels + option set resolution
    private _columnMetadata: Record<string, ColumnMeta> = {};
    private _slotToColumn: Record<string, string> = {};
    private _resolvedValues: Record<string, string> = {};
    // Colors resolved from lookup entity records
    private _resolvedColors: Record<string, string> = {};

    constructor() { }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
    ): void {
        this.context = context;
        console.log(`[InfoCard] v${CONTROL_VERSION} initialized`);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        // Detect form entity context for metadata resolution and @. syntax
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const modeAny = context.mode as Record<string, any>;
        if (!this._formEntityName && modeAny.contextInfo?.entityTypeName) {
            this._formEntityName = modeAny.contextInfo.entityTypeName;
            console.log("[InfoCard] Form entity:", this._formEntityName);
        }
        if (!this._formRecordId && modeAny.contextInfo?.entityId) {
            this._formRecordId = this.formatGuid(String(modeAny.contextInfo.entityId));
            console.log("[InfoCard] Form record ID:", this._formRecordId);
        }

        const data = this.collectData(context);
        const layout = this.resolveLayout(context);
        const hideEmpty = context.parameters.hideEmptyFields?.raw !== false; // default true
        const showBorder = context.parameters.showCardBorder?.raw !== false; // default true
        const showVersion = context.parameters.showVersionInfo?.raw === true; // default false
        const showTitle = context.parameters.showTitle?.raw !== false; // default true
        const startExpanded = context.parameters.startExpanded?.raw !== false; // default true
        const theme = this.resolveTheme(context);

        // Detect all related field mappings, split by source
        const allMappings = this.detectRelatedFields(context);
        const titleMappings = allMappings.filter(m => m.sourceSlot === "titleField");
        const currentRecordMappings = allMappings.filter(m => m.sourceSlot === "__currentRecord__");

        // Design-time detection: titleField is configured (valid type) but has no record data
        const titleParam = context.parameters.titleField;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const titleType = (titleParam as Record<string, any>)?.type;
        const designTime = titleParam != null
            && titleType !== "Unknown"
            && (titleParam.raw === null || titleParam.raw === undefined);

        // Build binding diagnostics for design-time panel (only warnings)
        const bindingDiagnostics = designTime ? this.buildBindingDiagnostics(context, allMappings) : undefined;

        const onOpenRecord = (entityType: string, id: string) => {
            context.navigation.openForm({
                entityName: entityType,
                entityId: id,
            });
        };

        // Record resolution is triggered by React useEffect via resolveRecordFields callback

        console.log("[InfoCard] title lookup:", data.title?.lookupEntityType, data.title?.lookupId,
            "| title mappings:", titleMappings.length,
            "| current-record mappings:", currentRecordMappings.length);

        return React.createElement(InfoCardComponent, {
            data,
            layout,
            hideEmpty,
            showBorder,
            showVersion,
            showTitle,
            startExpanded,
            designTime,
            theme,
            version: CONTROL_VERSION,
            relatedMappings: titleMappings,
            currentRecordMappings,
            currentRecordEntityType: this._formEntityName ?? undefined,
            currentRecordId: this._formRecordId ?? undefined,
            fetchRelatedData: this.boundFetchRelatedData,
            resolveRecordFields: (!designTime && this._formEntityName && this._formRecordId)
                ? this.boundResolveRecordFields : undefined,
            onOpenRecord,
            bindingDiagnostics,
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
        if (!param) return null;

        // Skip unconfigured properties (no type, no data, and no attributes)
        const hasAttributes = param.attributes && (param.attributes.DisplayName || param.attributes.LogicalName
            || param.attributes.displayName || param.attributes.logicalName);
        if (!param.type && !hasAttributes && (param.raw === null || param.raw === undefined)) return null;
        if (param.type === "Unknown") return null;
        // Attributes exist but have no identifying metadata and no type — field is not properly configured
        // Static input properties (static="true" in form XML) may have attributes={} but valid type and data
        if (param.attributes && !hasAttributes && !param.type) return null;

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const attrs = (param as Record<string, any>).attributes;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const paramAny = param as Record<string, any>;

        // ── Extract logical name ──
        const logicalName: string | null = attrs?.LogicalName ?? attrs?.logicalName ?? null;

        // Track slot → column mapping for entity metadata fetch
        if (logicalName) {
            this._slotToColumn[key] = logicalName;
        }

        // ── Extract display name from multiple sources ──
        let displayName: string | null = null;

        // 1. From entity metadata cache (most reliable, populated async)
        if (logicalName && this._columnMetadata[logicalName]) {
            displayName = this._columnMetadata[logicalName].displayName;
        }

        // 2. From param.attributes.DisplayName (various formats)
        if (!displayName) {
            displayName = this.extractDisplayNameFromMeta(attrs?.DisplayName ?? attrs?.displayName);
        }

        // 3. Param-level display name (some PCF hosts)
        if (!displayName && typeof paramAny.DisplayName === "string" && paramAny.DisplayName.length > 0) {
            displayName = paramAny.DisplayName;
        }

        // Check if we resolved the column name from the record fetch
        const resolvedColumn = this._slotToColumn[key] ?? logicalName;

        const raw = param.raw;

        // Label fallback chain: displayName → metadata cache → formatted column name → formatted slot key
        const cachedMeta = resolvedColumn ? this._columnMetadata[resolvedColumn] : null;
        const label = displayName
            ?? cachedMeta?.displayName
            ?? (resolvedColumn ? this.formatLogicalName(resolvedColumn) : null)
            ?? this.formatPropertyKey(key);

        const formatted = param.formatted;
        const resolvedValue = this._resolvedValues[key]; // from WebAPI record fetch

        let displayValue = "";
        let isEmpty = false;

        // @-prefixed values are related field placeholders — treat as empty
        // The actual value will be populated by the related field fetch
        if (typeof raw === "string" && raw.startsWith("@")) {
            isEmpty = true;
            // Show the raw reference as placeholder for developer visibility
            displayValue = raw;
        } else if (raw === null || raw === undefined) {
            isEmpty = true;
            displayValue = "---";
        } else if (resolvedValue) {
            // Use the formatted value from the WebAPI record (includes OptionSet labels, dates, etc.)
            displayValue = resolvedValue;
        } else if (formatted) {
            displayValue = formatted;
        } else if (raw instanceof Date) {
            // Date objects (DateTime fields when formatted is null)
            displayValue = raw.toLocaleDateString(undefined, {
                year: "numeric", month: "short", day: "numeric",
                hour: "2-digit", minute: "2-digit",
            });
        } else if (typeof raw === "object" && Array.isArray(raw) && raw.length > 0) {
            // Lookup
            displayValue = raw[0].name ?? String(raw[0].id ?? "---");
        } else if (typeof raw === "object" && !Array.isArray(raw)) {
            // Single EntityReference object: {Id: {_rawGuid, _formattedGuid}, Name, LogicalName}
            // or simple lookup object: {name, id, entityType}
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const obj = raw as Record<string, any>;
            displayValue = obj.name ?? obj.Name ?? "";
        } else {
            displayValue = String(raw);
        }

        if (displayValue === "" || displayValue === "---") {
            isEmpty = true;
        }

        // OptionSet: if displayValue is a raw number, resolve from metadata cache
        if (typeof raw === "number" && !formatted && !resolvedValue && resolvedColumn) {
            const cached = this._columnMetadata[resolvedColumn];
            if (cached?.options) {
                const opt = cached.options.find(o => o.value === raw);
                if (opt) {
                    displayValue = opt.label;
                    isEmpty = false;
                }
            }
        }

        // Format duration fields (stored as minutes in Dataverse)
        if (!isEmpty && attrs && this.isDurationField(attrs, raw, formatted)) {
            displayValue = this.formatDuration(Number(raw));
        }

        // Extract lookup entity/id for navigation and related field fetches
        const lookupRef = this.extractLookupRef(raw);
        let lookupEntityType = lookupRef?.entityType;
        const lookupId = lookupRef?.id;

        // Fallback: PCF may omit entityType — resolve from field metadata Targets
        if (lookupId && !lookupEntityType && attrs?.Targets) {
            const targets = Array.isArray(attrs.Targets) ? attrs.Targets
                : typeof attrs.Targets.getAll === "function" ? attrs.Targets.getAll() : [];
            if (targets.length > 0) lookupEntityType = targets[0];
        }

        // Resolve color: OptionSet from metadata cache, or lookup from resolved colors
        let optionColor: string | undefined = this._resolvedColors[key];
        if (!optionColor && typeof raw === "number" && resolvedColumn) {
            const cached = this._columnMetadata[resolvedColumn];
            const opt = cached?.options?.find(o => o.value === raw);
            if (opt?.color) optionColor = opt.color;
        }

        return {
            label,
            value: isEmpty ? "---" : displayValue,
            rawValue: raw,
            isEmpty,
            lookupEntityType,
            lookupId,
            optionColor,
        };
    }

    // ────────────────────────────────────────
    // Entity metadata (async label + option set resolution)
    // ────────────────────────────────────────

    /**
     * Fetch the current form record via WebAPI and match param values to columns.
     * This resolves labels and formatted values for bound fields where
     * param.attributes doesn't provide LogicalName (type-group properties).
     */
    // System/audit columns to exclude from value matching — these almost never
    // map to InfoCard slots and cause false positives (e.g., statuscode=1 vs bookingtype=1)
    private static readonly SYSTEM_COLUMNS = new Set([
        "statecode", "statuscode", "createdon", "modifiedon", "createdby", "modifiedby",
        "ownerid", "owningbusinessunit", "owningteam", "owninguser",
        "versionnumber", "timezoneruleversionnumber", "utcconversiontimezonecode",
        "importsequencenumber", "overriddencreatedon", "exchangerate",
    ]);

    /** Entities that store a color field on their record */
    private static readonly COLOR_FIELDS: Record<string, string> = {
        bookingstatus: "msdyn_statuscolor",
        msdyn_bookingstatus: "msdyn_statuscolor",
        msdyn_priority: "msdyn_prioritycolor",
    };

    /** Extract lookup entity type + ID from various PCF raw value formats */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private extractLookupRef(raw: any): { entityType?: string; id?: string } | null {
        if (Array.isArray(raw) && raw.length > 0 && raw[0].id) {
            return { entityType: raw[0].entityType ?? raw[0].etn, id: raw[0].id };
        }
        if (typeof raw === "object" && raw !== null && !Array.isArray(raw) && !(raw instanceof Date)) {
            const rawGuid = raw.Id?._formattedGuid ?? raw.Id?._rawGuid ?? raw.id;
            if (rawGuid) {
                return {
                    entityType: raw.LogicalName ?? raw.entityType,
                    id: this.formatGuid(String(rawGuid)),
                };
            }
        }
        return null;
    }

    /** Fetch a color field from a lookup entity record, normalizing hex format */
    private async fetchLookupColor(
        context: ComponentFramework.Context<IInputs>,
        entityType: string, id: string,
    ): Promise<string | undefined> {
        const colorField = InfoCard.COLOR_FIELDS[entityType];
        if (!colorField) return undefined;
        try {
            const rec = await context.webAPI.retrieveRecord(entityType, id, `?$select=${colorField}`);
            let color = rec[colorField];
            if (color && typeof color === "string" && !color.startsWith("#")) color = `#${color}`;
            if (color && typeof color === "string"
                && color.toLowerCase() !== "#ffffff" && color.toLowerCase() !== "#000000") {
                return color;
            }
        } catch { /* color fetch is optional */ }
        return undefined;
    }

    private async resolveRecordFieldsAsync(
        context: ComponentFramework.Context<IInputs>,
    ): Promise<Record<string, { label: string; value: string; color?: string }>> {
        if (!this._formEntityName || !this._formRecordId) return {};
        const overrides: Record<string, { label: string; value: string; color?: string }> = {};

        try {
            // Fetch full record — includes formatted value annotations for all columns
            const record = await context.webAPI.retrieveRecord(this._formEntityName, this._formRecordId);
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const rec = record as Record<string, any>;

            // Build column index: primitive columns only, excluding system fields
            const columnNames = Object.keys(rec).filter(k =>
                !k.includes("@") && !k.startsWith("_") && k !== "id"
                && !InfoCard.SYSTEM_COLUMNS.has(k)
            );

            // For each bound slot without a resolved column, match by value
            for (const key of ALL_SLOT_KEYS) {
                if (this._slotToColumn[key]) continue;
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const param = (context.parameters as Record<string, any>)[key];
                if (!param || param.type === "Unknown") continue;
                const raw = param.raw;
                if (raw === null || raw === undefined) continue;
                if (typeof raw === "string" && raw.startsWith("@")) continue;

                for (const colName of columnNames) {
                    const colVal = rec[colName];
                    if (colVal === null || colVal === undefined) continue;

                    if (this.valuesMatch(raw, colVal)) {
                        this._slotToColumn[key] = colName;
                        // Cache formatted value from WebAPI response
                        const fmtKey = `${colName}@OData.Community.Display.V1.FormattedValue`;
                        const fmtVal = rec[fmtKey]
                            ?? (typeof rec.getFormattedValue === "function" ? rec.getFormattedValue(colName) : null);
                        if (fmtVal) {
                            this._resolvedValues[key] = String(fmtVal);
                        }
                        break;
                    }
                }
            }

            console.log("[InfoCard] Record match:", Object.entries(this._slotToColumn)
                .map(([k, v]) => `${k}→${v}${this._resolvedValues[k] ? ` (${this._resolvedValues[k]})` : ""}`).join(", "));

            // Fetch metadata for all discovered columns (labels + option set defs)
            const allColumns = [...new Set(Object.values(this._slotToColumn))];
            if (allColumns.length > 0) {
                try {
                    const metadata = await context.utils.getEntityMetadata(this._formEntityName!, allColumns);
                    if (metadata?.Attributes) {
                        for (const attr of metadata.Attributes.getAll()) {
                            const logName = attr.LogicalName;
                            if (!logName) continue;
                            const entry: ColumnMeta = {
                                logicalName: logName,
                                displayName: this.extractDisplayNameFromMeta(attr.DisplayName) ?? this.formatLogicalName(logName),
                            };
                            // eslint-disable-next-line @typescript-eslint/no-explicit-any
                            const attrDesc = (attr as Record<string, any>).attributeDescriptor;
                            const optionSet = attrDesc?.OptionSet ?? attrDesc?.optionSet;
                            if (optionSet?.Options) {
                                entry.options = [];
                                for (const opt of optionSet.Options) {
                                    const optLabel = opt.Label?.UserLocalizedLabel?.Label
                                        ?? opt.Label?.userLocalizedLabel?.label
                                        ?? opt.Label ?? String(opt.Value);
                                    const optColor = opt.Color && opt.Color !== "#0000ff" ? opt.Color : undefined;
                                    entry.options.push({ value: opt.Value, label: String(optLabel), color: optColor });
                                }
                            }
                            this._columnMetadata[logName] = entry;
                        }
                    }
                    console.log("[InfoCard] Bound field labels:", Object.entries(this._slotToColumn)
                        .map(([k, col]) => `${k}→"${this._columnMetadata[col]?.displayName ?? col}"`).join(", "));
                } catch { /* metadata is optional */ }
            }

            // Resolve colors for lookup-based tag fields
            for (const key of TAG_KEYS) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const param = (context.parameters as Record<string, any>)[key];
                if (!param) continue;
                const raw = param.raw;
                if (!raw || (typeof raw === "string" && raw.startsWith("@"))) continue;
                const ref = this.extractLookupRef(raw);
                if (ref?.entityType && ref?.id) {
                    const color = await this.fetchLookupColor(context, ref.entityType, ref.id);
                    if (color) this._resolvedColors[key] = color;
                }
            }

            // Build overrides map for React to apply
            // Skip titleField — it's usage="bound" and already has correct formatted value
            for (const [slotKey, colName] of Object.entries(this._slotToColumn)) {
                if (slotKey === "titleField") continue;
                const meta = this._columnMetadata[colName];
                const fmtValue = this._resolvedValues[slotKey];
                const color = this._resolvedColors[slotKey];
                if (meta || fmtValue || color) {
                    overrides[slotKey] = {
                        label: meta?.displayName ?? this.formatLogicalName(colName),
                        value: fmtValue ?? "",
                        color,
                    };
                }
            }

            // Include tag colors even if the tag wasn't matched by value
            // (lookups like bookingstatus can't be matched via _column_value filtering)
            for (const key of TAG_KEYS) {
                const color = this._resolvedColors[key];
                if (color && !overrides[key]) {
                    overrides[key] = { label: "", value: "", color };
                } else if (color && overrides[key]) {
                    overrides[key].color = color;
                }
            }

            console.log("[InfoCard] Record overrides:", Object.entries(overrides)
                .map(([k, v]) => `${k}→"${v.label}": ${v.value}${v.color ? ` [${v.color}]` : ""}`).join(", "));
            return overrides;
        } catch (err) {
            console.warn("[InfoCard] resolveRecordFields failed:", err);
            return overrides;
        }
    }

    /** Match a PCF param raw value against a WebAPI record column value */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private valuesMatch(paramRaw: any, recordVal: any): boolean {
        // Number (OptionSet, integer, decimal)
        if (typeof paramRaw === "number" && typeof recordVal === "number") {
            return paramRaw === recordVal;
        }
        // Date: param is Date object, record is ISO string
        if (paramRaw instanceof Date && typeof recordVal === "string") {
            const recordDate = new Date(recordVal);
            return !isNaN(recordDate.getTime()) && Math.abs(paramRaw.getTime() - recordDate.getTime()) < 60000;
        }
        // String
        if (typeof paramRaw === "string" && typeof recordVal === "string") {
            return paramRaw === recordVal;
        }
        // Lookup array: compare ID
        if (Array.isArray(paramRaw) && paramRaw.length > 0 && paramRaw[0].id && typeof recordVal === "string") {
            return paramRaw[0].id.toLowerCase() === recordVal.toLowerCase();
        }
        return false;
    }

    private extractDisplayNameFromMeta(dn: unknown): string | null {
        if (typeof dn === "string" && dn.length > 0) return dn;
        if (dn && typeof dn === "object") {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const obj = dn as Record<string, any>;
            return obj.UserLocalizedLabel?.Label
                ?? obj.userLocalizedLabel?.label
                ?? obj.LocalizedLabels?.[0]?.Label
                ?? obj.localizedLabels?.[0]?.label
                ?? null;
        }
        return null;
    }

    // ────────────────────────────────────────
    // Related fields
    // ────────────────────────────────────────

    /**
     * Detect related field references using the @ convention.
     *
     * Two source types:
     *   @fieldName       → fetch from title field's lookup entity (work order)
     *   @nav.field       → fetch nav.field from title entity via $expand
     *   @.nav.field      → fetch from CURRENT form record via $expand
     *
     * The @. prefix navigates from the current record instead of the title entity.
     */
    private detectRelatedFields(context: ComponentFramework.Context<IInputs>): RelatedFieldMapping[] {
        const mappings: RelatedFieldMapping[] = [];
        const scanKeys = ALL_SLOT_KEYS.filter(k => k !== "titleField");

        for (const slotKey of scanKeys) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const param = (context.parameters as Record<string, any>)[slotKey];
            if (!param) continue;
            const raw = param.raw;
            if (typeof raw !== "string" || !raw.startsWith("@")) continue;

            let fetchField: string;
            let sourceSlot: string;

            if (raw.startsWith("@.")) {
                // @. prefix → navigate from current form record
                fetchField = raw.substring(2).trim();
                sourceSlot = "__currentRecord__";
            } else {
                // @ prefix → navigate from title entity
                fetchField = raw.substring(1).trim();
                sourceSlot = "titleField";
            }

            if (fetchField) {
                mappings.push({ sourceSlot, fetchField, targetSlot: slotKey });
            }
        }

        return mappings;
    }

    /**
     * Read a field value from a WebAPI record, handling lookups, direct fields, and images.
     * Tries OData annotation properties and getFormattedValue() method.
     */
    private readField(record: ComponentFramework.WebApi.Entity, col: string): {
        value: string; label: string; lookupId?: string; lookupEntityType?: string;
    } | null {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const rec = record as Record<string, any>;
        const getFormatted = typeof rec.getFormattedValue === "function"
            ? (key: string) => rec.getFormattedValue(key) as string | null
            : () => null;

        // Image columns
        if (col === "entityimage" || col.endsWith("image") || col.endsWith("_image")) {
            const allImageKeys = Object.keys(rec).filter(k =>
                k.toLowerCase().includes("image") && !k.includes("@") && !k.includes("timestamp"));
            for (const key of allImageKeys) {
                if (key.toLowerCase().endsWith("_url") && rec[key]) {
                    return { value: String(rec[key]), label: "image" };
                }
            }
            for (const key of allImageKeys) {
                const val = rec[key];
                if (val && typeof val === "string" && (val as string).length > 100 && !key.endsWith("_url")) {
                    return { value: `data:image/png;base64,${val}`, label: "image" };
                }
            }
            return null;
        }

        // Lookup pattern: _col_value
        const lookupKey = `_${col}_value`;
        const lookupFormatted = rec[`${lookupKey}@OData.Community.Display.V1.FormattedValue`]
            ?? rec[`${col}@OData.Community.Display.V1.FormattedValue`]
            ?? getFormatted(lookupKey)
            ?? getFormatted(col);
        const lookupVal = rec[lookupKey];
        if (lookupFormatted) {
            const etnKey = `${lookupKey}@Microsoft.Dynamics.CRM.lookuplogicalname`;
            return {
                value: String(lookupFormatted),
                label: col,
                lookupId: lookupVal ? String(lookupVal) : undefined,
                lookupEntityType: rec[etnKey] ? String(rec[etnKey]) : undefined,
            };
        }
        if (lookupVal != null) {
            return { value: String(lookupVal), label: col, lookupId: String(lookupVal) };
        }

        // Direct field with formatted value
        const formatted = rec[`${col}@OData.Community.Display.V1.FormattedValue`]
            ?? getFormatted(col);
        if (formatted) return { value: String(formatted), label: col };
        if (rec[col] != null) return { value: String(rec[col]), label: col };

        return null;
    }

    /**
     * Fetch related fields using OData with layered fallback.
     *
     * For title-entity fetches (@syntax): tries $expand for dotted paths, falls back to hop2.
     * For current-record fetches (@. syntax): skips $expand entirely (column logical names
     * aren't OData navigation properties) and goes straight to hop2.
     */
    private async fetchRelatedFields(
        context: ComponentFramework.Context<IInputs>,
        entityType: string,
        id: string,
        columns: string[],
        skipExpand = false,
    ): Promise<Record<string, { value: string; label: string; lookupId?: string; lookupEntityType?: string; color?: string }>> {
        try {
            console.log("[InfoCard] fetchRelatedFields:", entityType, id, columns, skipExpand ? "(hop2 only)" : "");
            const results: Record<string, { value: string; label: string; lookupId?: string; lookupEntityType?: string; color?: string }> = {};

            // ── 1. Classify columns ──────────────────────────────────────
            const directCols: string[] = [];
            const expandGroups: Record<string, string[]> = {};
            // For @. (current-record), dotted paths use hop2 not $expand
            const hop2Groups: Record<string, string[]> = {};

            for (const col of columns) {
                const dotIdx = col.indexOf(".");
                if (dotIdx > 0) {
                    const navProp = col.substring(0, dotIdx);
                    const field = col.substring(dotIdx + 1);
                    if (skipExpand) {
                        // Go straight to hop2 — column names aren't OData nav properties
                        if (!hop2Groups[navProp]) hop2Groups[navProp] = [];
                        if (!hop2Groups[navProp].includes(field)) hop2Groups[navProp].push(field);
                        // Don't add to directCols — hop2 fetches _navProp_value itself
                    } else {
                        if (!expandGroups[navProp]) expandGroups[navProp] = [];
                        if (!expandGroups[navProp].includes(field)) expandGroups[navProp].push(field);
                        if (!directCols.includes(navProp)) directCols.push(navProp);
                    }
                } else {
                    if (!directCols.includes(col)) directCols.push(col);
                }
            }

            const hasExpand = Object.keys(expandGroups).length > 0;

            // ── 2. Build combined query ──────────────────────────────────
            const selectParts = directCols;
            const expandParts = Object.entries(expandGroups).map(
                ([nav, fields]) => `${nav}($select=${fields.join(",")})`
            );
            let options = `?$select=${selectParts.join(",")}`;
            if (hasExpand) {
                options += `&$expand=${expandParts.join(",")}`;
            }

            // ── 3. Execute with layered fallback ─────────────────────────
            let record: ComponentFramework.WebApi.Entity = {};
            let usedFallback = false;

            try {
                record = await context.webAPI.retrieveRecord(entityType, id, options);
            } catch (expandErr) {
                console.warn("[InfoCard] Combined query failed — using fallback:", expandErr);
                usedFallback = true;
                record = {};

                // Fallback A: fetch direct columns per-column
                for (const col of directCols) {
                    // Try both primitive and lookup patterns
                    for (const selectCol of [col, `_${col}_value`]) {
                        try {
                            const r = await context.webAPI.retrieveRecord(entityType, id, `?$select=${selectCol}`);
                            for (const k of Object.keys(r)) { record[k] = r[k]; }
                        } catch { /* column doesn't exist in this form */ }
                    }
                }

                // Fallback B: try each $expand group individually
                for (const [nav, fields] of Object.entries(expandGroups)) {
                    try {
                        const r = await context.webAPI.retrieveRecord(
                            entityType, id, `?$expand=${nav}($select=${fields.join(",")})`
                        );
                        if (r[nav]) record[nav] = r[nav];
                    } catch {
                        // Fallback C: legacy hop2 — resolve lookup, then fetch from related entity
                        console.warn(`[InfoCard] $expand(${nav}) failed — trying hop2`);
                        const hop1 = this.readField(record, nav);
                        if (hop1?.lookupId && hop1?.lookupEntityType) {
                            try {
                                const hop2Record = await context.webAPI.retrieveRecord(
                                    hop1.lookupEntityType, hop1.lookupId, `?$select=${fields.join(",")}`
                                );
                                record[nav] = hop2Record;
                            } catch {
                                console.warn(`[InfoCard] Hop2 for ${nav} failed — skipping`);
                            }
                        }
                    }
                }
            }

            // ── 3b. Follow-up: ensure lookup primitives are fetched ──────
            if (!usedFallback) {
                const missingLookupCols = directCols.filter(col => {
                    const val = record[col];
                    const lookupVal = record[`_${col}_value`];
                    return lookupVal === undefined &&
                        (val === undefined || val === null || (typeof val === "object" && val !== null));
                });

                if (missingLookupCols.length > 0) {
                    const lookupSelect = missingLookupCols.map(col => `_${col}_value`).join(",");
                    try {
                        const lookupRecord = await context.webAPI.retrieveRecord(entityType, id, `?$select=${lookupSelect}`);
                        for (const k of Object.keys(lookupRecord)) { record[k] = lookupRecord[k]; }
                    } catch {
                        for (const col of missingLookupCols) {
                            try {
                                const r = await context.webAPI.retrieveRecord(entityType, id, `?$select=_${col}_value`);
                                for (const k of Object.keys(r)) { record[k] = r[k]; }
                            } catch { /* skip */ }
                        }
                    }
                }
            }

            // ── 4. Parse direct fields ───────────────────────────────────
            for (const col of directCols) {
                const isExpandIntermediary = expandGroups[col] && !columns.includes(col);
                const result = this.readField(record, col);
                if (result && !isExpandIntermediary) {
                    results[col] = result;
                }
            }

            // ── 5. Parse expanded (nested) fields ────────────────────────
            for (const [nav, fields] of Object.entries(expandGroups)) {
                const nested = record[nav];
                const nestedObj = (nested && typeof nested === "object" && !Array.isArray(nested))
                    ? nested as Record<string, unknown> : null;

                if (nestedObj) {
                    for (const field of fields) {
                        const result = this.readField(nestedObj as ComponentFramework.WebApi.Entity, field);
                        if (result) {
                            results[`${nav}.${field}`] = result;
                        }
                    }
                }

                // Hop2 fallback: if expanded parsing missed fields
                const missingFields = fields.filter(f => !results[`${nav}.${f}`]);
                if (missingFields.length > 0) {
                    const hop1 = this.readField(record, nav);
                    if (hop1?.lookupId && hop1?.lookupEntityType) {
                        try {
                            const hop2Record = await context.webAPI.retrieveRecord(
                                hop1.lookupEntityType, hop1.lookupId, `?$select=${missingFields.join(",")}`
                            );
                            for (const field of missingFields) {
                                const result = this.readField(hop2Record, field);
                                if (result) {
                                    results[`${nav}.${field}`] = result;
                                }
                            }
                        } catch { /* hop2 failed */ }
                    } else {
                        // No lookup primitive yet — try fetching it
                        try {
                            const lookupRecord = await context.webAPI.retrieveRecord(entityType, id, `?$select=_${nav}_value`);
                            const hop1b = this.readField(lookupRecord, nav);
                            if (hop1b?.lookupId && hop1b?.lookupEntityType) {
                                const hop2Record = await context.webAPI.retrieveRecord(
                                    hop1b.lookupEntityType, hop1b.lookupId, `?$select=${missingFields.join(",")}`
                                );
                                for (const field of missingFields) {
                                    const result = this.readField(hop2Record, field);
                                    if (result) {
                                        results[`${nav}.${field}`] = result;
                                    }
                                }
                            }
                        } catch { /* hop2b failed */ }
                    }
                }
            }

            // ── 5b. Direct hop2 for @. dotted paths (skip $expand) ────────
            for (const [nav, fields] of Object.entries(hop2Groups)) {
                // Resolve the lookup column to get target entity + ID
                const hop1 = this.readField(record, nav);
                if (hop1?.lookupId && hop1?.lookupEntityType) {
                    try {
                        const hop2Record = await context.webAPI.retrieveRecord(
                            hop1.lookupEntityType, hop1.lookupId, `?$select=${fields.join(",")}`
                        );
                        for (const field of fields) {
                            const result = this.readField(hop2Record, field);
                            if (result) {
                                results[`${nav}.${field}`] = result;
                            }
                        }
                    } catch { /* hop2 failed */ }
                } else {
                    // Lookup primitive not in record yet — fetch it
                    try {
                        const lookupRecord = await context.webAPI.retrieveRecord(entityType, id, `?$select=_${nav}_value`);
                        // Merge into record so step 6b can resolve the target entity
                        for (const k of Object.keys(lookupRecord)) { record[k] = lookupRecord[k]; }
                        const hop1b = this.readField(lookupRecord, nav);
                        if (hop1b?.lookupId && hop1b?.lookupEntityType) {
                            const hop2Record = await context.webAPI.retrieveRecord(
                                hop1b.lookupEntityType, hop1b.lookupId, `?$select=${fields.join(",")}`
                            );
                            for (const field of fields) {
                                const result = this.readField(hop2Record, field);
                                if (result) {
                                    results[`${nav}.${field}`] = result;
                                }
                            }
                        }
                    } catch { /* hop2b failed */ }
                }
            }

            // ── 6. Resolve display names from entity metadata ────────────
            // 6a. Direct fields on the source entity
            try {
                const colsToResolve = Object.keys(results).filter(k => !k.startsWith("__") && !k.includes("."));
                if (colsToResolve.length > 0) {
                    const entityMeta = await context.utils.getEntityMetadata(entityType, colsToResolve);
                    if (entityMeta?.Attributes) {
                        for (const attr of entityMeta.Attributes.getAll()) {
                            const logicalName = attr.LogicalName;
                            const dn = this.extractDisplayNameFromMeta(attr.DisplayName);
                            if (logicalName && dn && results[logicalName]) {
                                results[logicalName].label = dn;
                            }
                        }
                    }
                }
            } catch { /* metadata fetch is optional */ }

            // 6b. Dotted-path fields: resolve from the target entity
            const allDottedGroups = { ...expandGroups, ...hop2Groups };
            for (const [nav, fields] of Object.entries(allDottedGroups)) {
                const parsedFields = fields.filter(f => results[`${nav}.${f}`]);
                if (parsedFields.length === 0) continue;
                // Determine target entity from the nav property's lookup metadata
                const hop1 = this.readField(record, nav);
                const targetEntity = hop1?.lookupEntityType;
                if (!targetEntity) continue;
                try {
                    const navMeta = await context.utils.getEntityMetadata(targetEntity, parsedFields);
                    if (navMeta?.Attributes) {
                        for (const attr of navMeta.Attributes.getAll()) {
                            const logicalName = attr.LogicalName;
                            const dn = this.extractDisplayNameFromMeta(attr.DisplayName);
                            if (logicalName && dn && results[`${nav}.${logicalName}`]) {
                                results[`${nav}.${logicalName}`].label = dn;
                            }
                        }
                    }
                } catch { /* metadata fetch is optional */ }
            }

            // ── 7. Resolve colors for lookup results from known color entities ──
            for (const key of Object.keys(results)) {
                const val = results[key];
                if (!val.lookupId || !val.lookupEntityType) continue;
                const color = await this.fetchLookupColor(context, val.lookupEntityType, val.lookupId);
                if (color) results[key] = { ...val, color };
            }

            const resultSummary = Object.entries(results).map(([k, v]) => `${k}="${v.value}" (${v.label})`).join(", ");
            console.log(`[InfoCard] fetchRelatedFields done: ${resultSummary}`);

            return results;
        } catch (err) {
            console.error("[InfoCard] fetchRelatedFields ERROR:", err);
            return {};
        }
    }

    // ────────────────────────────────────────
    // Binding diagnostics (design-time panel)
    // ────────────────────────────────────────

    private buildBindingDiagnostics(
        context: ComponentFramework.Context<IInputs>,
        mappings: RelatedFieldMapping[],
    ): BindingDiagnostic[] {
        const diagnostics: BindingDiagnostic[] = [];

        for (const key of ALL_SLOT_KEYS) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const param = (context.parameters as Record<string, any>)[key];
            if (!param) continue;

            const paramType = param.type;
            const raw = param.raw;

            // Skip truly unconfigured slots
            if (!paramType && (raw === null || raw === undefined)) continue;
            if (paramType === "Unknown") continue;

            const entry: BindingDiagnostic = {
                slotKey: key,
                slotLabel: SLOT_LABELS[key] ?? key,
                bindingType: "bound",
                rawExpression: "",
            };

            if (typeof raw === "string" && raw.startsWith("@.")) {
                entry.bindingType = "current-related";
                entry.rawExpression = raw;
                const path = raw.substring(2).trim();
                if (!path) {
                    entry.warning = "Empty path after @. — expected @.navProp.field";
                } else if (!path.includes(".")) {
                    entry.warning = `Direct field '${path}' on current record — use column binding instead of @. for single-hop`;
                }
            } else if (typeof raw === "string" && raw.startsWith("@")) {
                entry.bindingType = "title-related";
                entry.rawExpression = raw;
                const path = raw.substring(1).trim();
                if (!path) {
                    entry.warning = "Empty path after @ — expected @fieldName";
                }
            } else {
                entry.bindingType = "bound";
                const logicalName = param.attributes?.LogicalName ?? param.attributes?.logicalName;
                entry.rawExpression = logicalName ?? `[${paramType ?? "unknown type"}]`;
            }

            // Only include entries with warnings (syntax issues)
            if (entry.warning) {
                diagnostics.push(entry);
            }
        }

        return diagnostics;
    }

    // ────────────────────────────────────────
    // Duration formatting
    // ────────────────────────────────────────

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private isDurationField(attrs: Record<string, any>, raw: unknown, formatted: string | undefined): boolean {
        if (typeof raw !== "number") return false;
        const fmt = String(attrs?.Format ?? "").toLowerCase();
        if (fmt === "duration") return true;
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
    // Helpers
    // ────────────────────────────────────────

    /** Format a compact GUID (no hyphens) into standard 8-4-4-4-12 format */
    private formatGuid(raw: string): string {
        const hex = raw.replace(/[{}-]/g, "");
        if (hex.length !== 32) return raw;
        return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
    }

    /** Convert property key like "gridField1" to "Grid 1" */
    private formatPropertyKey(key: string): string {
        return key
            .replace(/Field(\d*)$/, " $1")
            .replace(/([a-z])([A-Z])/g, "$1 $2")
            .replace(/^\w/, c => c.toUpperCase())
            .trim();
    }

    /** Convert Dataverse logical name to readable label: "msdyn_workordertype" → "Work Order Type" */
    private formatLogicalName(logicalName: string): string {
        return logicalName
            .replace(/^(msdyn_|new_|cr[a-z0-9]+_|ukf_|jdp_|cli_)/, "")
            .replace(/_/g, " ")
            .replace(/([a-z])([A-Z])/g, "$1 $2")
            .replace(/\b\w/g, c => c.toUpperCase())
            .trim();
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
